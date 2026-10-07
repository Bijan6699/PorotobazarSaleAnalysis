# -*- coding: utf-8 -*-
"""داشبورد تحلیلی فروش، مشتریان، کالا و پورسانت — تک‌فایلی Streamlit"""
from __future__ import annotations
import io, re, xml.etree.ElementTree as ET
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="داشبورد تحلیلی فروش | کیان‌پرداز", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
NS = "{urn:schemas-microsoft-com:office:spreadsheet}"

st.markdown(r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;600;700;800&display=swap');
:root{--ink:#172554;--muted:#64748b;--line:#e2e8f0;--bg:#f5f8fc;--blue:#2563eb;--teal:#0f9f8f}
html,body,[class*="css"],.stApp{font-family:'Vazirmatn',Tahoma,sans-serif!important;direction:rtl;text-align:right}
.stApp{background:linear-gradient(180deg,#f7faff 0%,#fff 34%)}
section[data-testid="stSidebar"]{border-left:1px solid var(--line);border-right:0}
.block-container{padding-top:1.7rem;padding-bottom:3rem;max-width:1500px}
.hero{background:linear-gradient(120deg,#102a5e,#2563eb 62%,#14b8a6);color:white;padding:26px 30px;border-radius:20px;margin:4px 0 22px;box-shadow:0 12px 30px #1e3a8a22}
.hero h1{font-size:1.75rem;margin:0 0 6px;font-weight:800}.hero p{opacity:.9;margin:0}
[data-testid="stMetric"]{background:white;border:1px solid #e6edf7;border-radius:16px;padding:16px 18px;box-shadow:0 5px 18px #0f172a08;min-width:0;overflow:hidden}
[data-testid="stMetricValue"]{font-size:clamp(0.9rem,1.4vw,1.3rem)!important;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;direction:ltr;text-align:right;max-width:100%;word-break:keep-all}
[data-testid="stMetricLabel"]{color:#64748b;font-size:clamp(.78rem,1vw,.9rem);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
[data-testid="stMetricLabel"]{color:#64748b}.stTabs [data-baseweb="tab-list"]{gap:6px;border-bottom:1px solid var(--line)}
.stTabs [data-baseweb="tab"]{border-radius:11px 11px 0 0;padding:10px 14px}
div[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:12px;overflow:hidden}
.small-note{color:#64748b;font-size:.88rem}
</style>""", unsafe_allow_html=True)

# ---------- file readers ----------
def read_upload(upload):
    raw=upload.getvalue()
    if raw.lstrip().startswith(b"<?xml") or b"urn:schemas-microsoft-com:office:spreadsheet" in raw[:5000]:
        root=ET.fromstring(raw); rows=[]
        for row in root.iter(NS+"Row"):
            # SpreadsheetML ss:Index is an absolute, 1-based column position.
            # Expand every omitted cell so each row shares the same true grid.
            values=[]; col=1
            for cell in row.findall(NS+"Cell"):
                idx=cell.attrib.get(NS+"Index")
                if idx is not None:
                    target=int(idx)
                    if target < col:
                        raise ValueError(f"Invalid SpreadsheetML ss:Index={target} (current column {col})")
                    values.extend([None] * (target-col)); col=target
                data=cell.find(NS+"Data")
                values.append(data.text if data is not None else None); col+=1
            if any(v is not None and str(v).strip() for v in values): rows.append(values)
        width=max((len(r) for r in rows),default=0)
        return pd.DataFrame([r+[None]*(width-len(r)) for r in rows],columns=range(1,width+1)),"xml"
    try:
        book=pd.ExcelFile(io.BytesIO(raw))
        # read first sheet as a raw grid; preserves report layouts without trusting headers
        return pd.read_excel(book, sheet_name=0, header=None),"excel"
    except Exception as exc:
        raise ValueError(f"قالب فایل پشتیبانی نشد: {exc}") from exc

def number(v):
    if pd.isna(v): return np.nan
    if isinstance(v,(int,float,np.number)): return float(v)
    s=str(v).strip().replace("٬","").replace(",","").replace("،","").replace("٫",".")
    trans=str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩","01234567890123456789"); s=s.translate(trans)
    try:return float(s)
    except:return np.nan

def clean_text(v, default=""):
    if pd.isna(v): return default
    return re.sub(r"\s+"," ",str(v).replace("\xa0"," ")).strip()

def brand_of(name):
    found=re.findall(r"[\(（]([^()（）]+)[\)）]",clean_text(name))
    return found[-1].strip() if found else "بدون برند / متفرقه"

def detect_column_map(raw):
    """Find the Persian header row and map logical fields to column positions (1-based XML / 0-based excel)."""
    import numpy as _np
    candidates = {
        "کد کالا": ("کد کالا", "کدکالا"),
        "کالا": ("عنوان کالا", "نام کالا", "کالا", "نام محصول"),
        "مشتری": ("نام طرف حساب", "نام طرف حساب2", "مشتری", "نام مشتری"),
        "نوع سند": ("نوع فاکتور", "نوع سند", "نوع"),
        "تاریخ": ("تاریخ", "تاريخ"),
        "شماره فاکتور": ("فاکتور", "شماره فاکتور", "شماره فاکتور2"),
        "تعداد": ("تعداد",),
        "قیمت واحد": ("قیمت واحد", "في واحد", "قيمت واحد"),
        "تخفیف": ("تخفیف واحد", "تخفیف"),
        "مبلغ": ("جمع", "جمع کل", "مبلغ کل", "مبلغ"),
        "ویزیتور": ("نام ویزیتور", "ویزیتور", "نام فروشنده", "فروشنده", "عامل فروش"),
    }
    hits = []
    for i in range(min(30, len(raw))):
        row = [clean_text(v) for v in raw.iloc[i].tolist()]
        matched = {}
        for logical, names in candidates.items():
            for j, v in enumerate(row):
                if v in names:
                    matched[logical] = raw.columns[j]
                    break
        if len(matched) >= 4 and ("کالا" in matched or "مشتری" in matched):
            hits.append((i, len(matched), matched))
    if not hits:
        return None
    hits.sort(key=lambda t: -t[1])
    best_i, _, best = hits[0]
    return best_i, best


def headerless_fallback(raw):
    """When no Persian header exists (old compact layout), guess by column content."""
    ncols = raw.shape[1]
    scores = {}
    text_cols = {}
    for c in raw.columns:
        vals = raw[c].dropna().astype(str).head(300)
        text_cols[c] = vals
        scores[c] = {
            "has_farsi_name": sum(1 for v in vals if re.search(r"فروش|برگشت", v)),
            "has_date": sum(1 for v in vals if re.match(r"^14\d{2}/", v.strip())),
        }
    doc_col = max(scores, key=lambda c: scores[c]["has_farsi_name"]) if scores else None
    date_col = max(scores, key=lambda c: scores[c]["has_date"]) if scores else None
    # doc type usually just left of the date; item right of it
    order = list(raw.columns)
    if doc_col is None:
        return None
    di = order.index(date_col) if date_col is not None and date_col is not doc_col else order.index(doc_col)
    return {
        "نوع سند": order.index(doc_col) + 1,
        "تاریخ": order.index(date_col) + 1 if date_col is not None else None,
        "شماره فاکتور": order[di - 1] + 1 if di - 1 >= 0 else None,
        "کالا": order[di + 1] + 1 if di + 1 < len(order) else None,
        "کد کالا": order[di + 2] + 1 if di + 2 < len(order) else None,
        "مشتری": order[di + 3] + 1 if di + 3 < len(order) else None,
        "تعداد": order[di - 3] + 1 if di - 3 >= 0 else None,
        "قیمت واحد": order[di - 2] + 1 if di - 2 >= 0 else None,
        "تخفیف": order[di - 4] + 1 if di - 4 >= 0 else None,
        "مبلغ": None,
    }


def sales_from_raw(raw, fname):
    """Parse Kian RepFroshDetail reports: headered SpreadsheetML (new) and compact headerless (old)."""
    if raw.empty:
        return pd.DataFrame()
    detected = detect_column_map(raw)
    if detected:
        header_i, mp = detected
        # Some FastReport SpreadsheetML exports place a heading in the first
        # cell of a merged pair, while values are written to the next cell.
        # Shift only when the mapped column is largely empty and its right
        # neighbor consistently contains data; this leaves normal/legacy grids intact.
        if raw.shape[1] and isinstance(raw.columns[0], (int, np.integer)):
            for logical in ("کالا", "کد کالا"):
                col=mp.get(logical)
                if col is None or col not in raw.columns: continue
                nxt=col+1
                if nxt not in raw.columns: continue
                body=raw.iloc[header_i+1:]
                here=body[col].map(lambda v: bool(clean_text(v)))
                there=body[nxt].map(lambda v: bool(clean_text(v)))
                if len(body) and there.sum() >= 3 and here.sum() <= max(1, int(there.sum()*0.2)):
                    mp[logical]=nxt
        # guard: doc-type column must actually contain sale/return values
        if "نوع سند" in mp:
            col = mp["نوع سند"]
            vals = set(clean_text(v) for v in raw[col].iloc[1:200].dropna() if str(v).strip())
            if not vals & {"فروش", "برگشت از فروش"}:
                detected = None
                mp = None
    if not detected:
        mp = headerless_fallback(raw)
        if not mp:
            return pd.DataFrame()
        if "مبلغ" not in mp or mp.get("مبلغ") is None:
            mp["مبلغ"] = 1  # old compact layout: first column is the net total
        mp.setdefault("کد کالا", None); mp.setdefault("کالا", None); mp.setdefault("مشتری", None)
        mp.setdefault("نوع سند", None); mp.setdefault("تاریخ", None); mp.setdefault("شماره فاکتور", None)
        mp.setdefault("تعداد", None); mp.setdefault("قیمت واحد", None); mp.setdefault("تخفیف", None)
        mp.setdefault("ویزیتور", None)

    def cell(r, key):
        idx = mp.get(key)
        return clean_text(r.get(idx)) if idx is not None else ""

    out = []
    for _, r in raw.iterrows():
        typ = cell(r, "نوع سند")
        if typ not in ("فروش", "برگشت از فروش"):
            continue
        item = cell(r, "کالا")
        customer = cell(r, "مشتری") or "نامشخص"
        if not item and customer == "نامشخص":
            continue
        out.append({
            "نوع سند": typ,
            "کد کالا": cell(r, "کد کالا"),
            "کالا": item or "نامشخص",
            "برند": brand_of(item),
            "مشتری": customer,
            "تاریخ": cell(r, "تاریخ"),
            "شماره فاکتور": cell(r, "شماره فاکتور"),
            "تعداد": number(r.get(mp["تعداد"])) if mp.get("تعداد") is not None else np.nan,
            "قیمت واحد": number(r.get(mp["قیمت واحد"])) if mp.get("قیمت واحد") is not None else np.nan,
            "تخفیف": number(r.get(mp["تخفیف"])) if mp.get("تخفیف") is not None else 0,
            "مبلغ": number(r.get(mp["مبلغ"])) if mp.get("مبلغ") is not None else np.nan,
            "ویزیتور": cell(r, "ویزیتور"),
            "منبع": fname,
        })
    return pd.DataFrame(out)


def parse_aux(raw):
    # Return a display-friendly frame; identify meaningful header where possible.
    if raw.empty:return raw
    for i in range(min(25,len(raw))):
        vals=[clean_text(x) for x in raw.iloc[i].tolist()]
        if any("کد کالا" in x for x in vals) and any("نام کالا" in x for x in vals):
            d=raw.iloc[i+1:].copy(); d.columns=[v or f"ستون {j+1}" for j,v in enumerate(vals)]
            return d.dropna(how="all").reset_index(drop=True)
    return raw.dropna(how="all").reset_index(drop=True)

def fmt(n): return f"{n:,.0f}" if pd.notna(n) else "—"
def money_columns(df, cols):
    return df.style.format({c:"{:,.0f}" for c in cols if c in df.columns}, na_rep="—")

st.sidebar.markdown("## 📁 بارگذاری گزارش‌ها")
st.sidebar.caption("می‌توانید چند فایل را هم‌زمان انتخاب کنید؛ نوع گزارش از نام فایل تشخیص داده می‌شود.")
files=st.sidebar.file_uploader("فایل‌های فروش، پورسانت یا مقایسه ویزیتورها",type=["xls","xlsx"],accept_multiple_files=True)
rep_name=st.sidebar.text_input("نام ویزیتور / منطقه",value="",help="اگر گزارش ستون ویزیتور نداشته باشد یا مقادیر آن «نامشخص» باشد، این نام به‌صورت دستی درج می‌شود. برای چند فایل، نام‌ها را با ویرگول جدا کنید تا به ترتیب بارگذاری اختصاص یابند؛ یا نام هر فایل را در کادر زیر آن وارد کنید.")
manual_visitors=[s.strip() for s in re.split(r"[،,\n;]+",rep_name) if s.strip()]
file_visitor={}
for _fi,_f in enumerate(files or []):
    _lbl=f"ویزیتور برای «{_f.name}»"
    _dv=manual_visitors[_fi] if _fi<len(manual_visitors) else (manual_visitors[0] if len(manual_visitors)==1 else "")
    file_visitor[_f.name]=st.sidebar.text_input(_lbl,value=_dv,key=f"vis_{_f.name}").strip()
if files:
    st.sidebar.caption("نام واردشده فقط وقتی استفاده می‌شود که گزارش، ستون ویزیتور معتبر نداشته باشد.")
monthly_target=st.sidebar.number_input("تارگت ماهانه فروش خالص",min_value=0.0,value=100000000.0,step=10000000.0,format="%.0f",help="مبنای نمودار تحقق هدف؛ واحد مطابق مبلغ گزارش فروش است.")
recent_periods=st.sidebar.number_input("تعداد تاریخ‌های اخیر برای هشدار عدم فعالیت",min_value=1,max_value=30,value=3,step=1)
previous=st.sidebar.file_uploader("فایل فروش قبلی برای تحلیل مشتریان (اختیاری)",type=["xls","xlsx"],key="previous")

if not files:
    st.markdown('<div class="hero"><h1>📊 داشبورد تحلیلی فروش کیان‌پرداز</h1><p>تحلیل یکپارچه فروش، مشتریان، برندها، کالاها و پورسانت با خروجی پاک‌سازی‌شده</p></div>',unsafe_allow_html=True)
    st.info("از نوار کناری یک یا چند فایل گزارش را بارگذاری کنید. فایل‌های SpreadsheetML با پسوند XLS، اکسل XLS/XLSX و فایل‌های دوره قبل پشتیبانی می‌شوند.")
    st.markdown("**گزارش‌های قابل پشتیبانی** · ریز فروش RepFroshDetail · فروش و پورسانت کالا · مقایسه کالایی ویزیتورها")
    st.stop()

sales_parts=[]; aux={}; errors=[]
for f in files:
    try:
        raw,kind=read_upload(f); name=f.name.lower()
        if "frosh" in name or "فروش" in name or ("compare" not in name and "مقایسه" not in name and not any(k in name for k in ["porsant","پورسانت"])):
            parsed=sales_from_raw(raw,f.name)
            if not parsed.empty:sales_parts.append(parsed)
            else: aux[f.name]=parse_aux(raw)
        else: aux[f.name]=parse_aux(raw)
    except Exception as e: errors.append(f"{f.name}: {e}")
for err in errors: st.warning(err)
sales=pd.concat(sales_parts,ignore_index=True) if sales_parts else pd.DataFrame()
if sales.empty:
    st.markdown('<div class="hero"><h1>📊 داشبورد تحلیلی فروش کیان‌پرداز</h1><p>فایل بارگذاری شد؛ برای نمایش تحلیل، گزارش ریز فروش معتبر لازم است.</p></div>',unsafe_allow_html=True)
    st.warning("هیچ ردیف فروش یا برگشت فروش قابل شناسایی نبود. فایل RepFroshDetail را انتخاب کنید؛ فایل‌های پورسانت به تنهایی برای تحلیل فروش کافی نیستند.")
    for name,df in aux.items():
        with st.expander(name): st.dataframe(df.head(50),use_container_width=True)
    st.stop()

for c in ["مبلغ","تعداد","تخفیف","قیمت واحد"]: sales[c]=pd.to_numeric(sales[c],errors="coerce").fillna(0)
sales["تاریخ"]=sales["تاریخ"].fillna("").astype(str).str.strip()
sales["مشتری"]=sales["مشتری"].replace("","نامشخص").fillna("نامشخص")
sales["کالا"]=sales["کالا"].replace("","نامشخص").fillna("نامشخص")
sales["برند"]=sales["برند"].fillna("بدون برند / متفرقه")
# --- ویزیتور دستی: اگر گزارش ستون ویزیتور معتبر ندارد، نام واردشده در نوار کناری استفاده می‌شود ---
def _visitor_is_known(s: pd.Series) -> bool:
    v=s.dropna().astype(str).str.strip()
    v=v[v.ne("") & v.ne("نامشخص")]
    return not v.empty
if "ویزیتور" not in sales.columns: sales["ویزیتور"]=""
if "منبع" not in sales.columns: sales["منبع"]=""
_sales_files=list(dict.fromkeys(sales["منبع"].astype(str)))
_fallback={}
for _i,_n in enumerate(_sales_files):
    _name=file_visitor.get(_n,"")
    if not _name and manual_visitors:
        _name=manual_visitors[_i] if _i<len(manual_visitors) else manual_visitors[0]
    _fallback[_n]=_name
_v=sales["ویزیتور"].fillna("").astype(str).str.strip()
_unknown=_v.isin(["","نامشخص"])
_manual=sales["منبع"].astype(str).map(lambda n:_fallback.get(n,""))
sales["ویزیتور"]=_v.where(~_unknown,_manual).replace("","نامشخص").fillna("نامشخص")
# magnitude in reports can be represented as positive return amount
sale=sales[sales["نوع سند"]=="فروش"]; returns=sales[sales["نوع سند"]=="برگشت از فروش"]
gross=float(sale["مبلغ"].sum()); ret=float(returns["مبلغ"].sum()); net=gross-ret
ret_rate=ret/gross*100 if gross else 0
invoices=sales.loc[sales["شماره فاکتور"].astype(str)!="","شماره فاکتور"].nunique()
brand=sales.pivot_table(index="برند",columns="نوع سند",values="مبلغ",aggfunc="sum",fill_value=0).reset_index()
for c in ["فروش","برگشت از فروش"]:
    if c not in brand:brand[c]=0
brand["فروش خالص"]=brand["فروش"]-brand["برگشت از فروش"]; brand["درصد سهم"]=brand["فروش خالص"].div(net if net else 1).mul(100); brand=brand.sort_values("فروش خالص",ascending=False)
cust=sales.pivot_table(index="مشتری",columns="نوع سند",values="مبلغ",aggfunc="sum",fill_value=0).reset_index()
for c in ["فروش","برگشت از فروش"]:
    if c not in cust:cust[c]=0
cust["فروش خالص"]=cust["فروش"]-cust["برگشت از فروش"]; cust["نرخ مرجوعی (%)"]=np.where(cust["فروش"]>0,cust["برگشت از فروش"]/cust["فروش"]*100,0); cust=cust.sort_values("فروش خالص",ascending=False)
prod=sales.pivot_table(index=["کالا","برند"],columns="نوع سند",values="مبلغ",aggfunc="sum",fill_value=0).reset_index()
for c in ["فروش","برگشت از فروش"]:
    if c not in prod:prod[c]=0
prod["فروش خالص"]=prod["فروش"]-prod["برگشت از فروش"]; prod=prod.sort_values("فروش خالص",ascending=False)
dates=sorted([d for d in sales["تاریخ"].unique() if d])
# Reusable customer churn analysis: compare chronological halves and detect inactivity.
_sale=sales[sales["نوع سند"]=="فروش"].copy()
_period_dates=sorted(_sale.loc[_sale["تاریخ"]!="","تاریخ"].unique())
_split=max(1,len(_period_dates)//2)
_first_dates=set(_period_dates[:_split]); _second_dates=set(_period_dates[_split:])
_churn_rows=[]
for _customer,_group in _sale.groupby("مشتری",dropna=False):
    _old=_group[_group["تاریخ"].isin(_first_dates)]; _new=_group[_group["تاریخ"].isin(_second_dates)]
    _old_spend=float(_old["مبلغ"].sum()); _new_spend=float(_new["مبلغ"].sum())
    _old_freq=_old["شماره فاکتور"].replace("",np.nan).nunique() or len(_old)
    _new_freq=_new["شماره فاکتور"].replace("",np.nan).nunique() or len(_new)
    _reason=[]
    if _old_spend>0 and _new_spend<_old_spend*.5:
        _reason.append(f"افت {100*(1-_new_spend/_old_spend):.0f} درصدی فروش")
    if _old_freq>0 and _new_freq<_old_freq*.5:
        _reason.append("کاهش معنادار دفعات خرید")
    _last=_group["تاریخ"].replace("",np.nan).dropna().max() if len(_group) else ""
    _inactive=bool(_period_dates and _last not in set(_period_dates[-int(recent_periods):]))
    if _inactive:_reason.append(f"عدم سفارش در {int(recent_periods)} تاریخ اخیر")
    _level="ریسک بالا" if (_inactive and _reason) or (len(_reason)>=2) else ("متوسط" if _reason else "فعال")
    _churn_rows.append({"مشتری":_customer,"وضعیت":_level,"دلایل": "، ".join(_reason) if _reason else "خرید بدون افت معنادار",
        "فروش نیمه اول":_old_spend,"فروش نیمه دوم":_new_spend,"تعداد فاکتور نیمه اول":int(_old_freq),"تعداد فاکتور نیمه دوم":int(_new_freq),"آخرین تاریخ خرید":_last})
churn=pd.DataFrame(_churn_rows)
if not churn.empty:
    _order={"ریسک بالا":0,"متوسط":1,"فعال":2}; churn["_sort"]=churn["وضعیت"].map(_order); churn=churn.sort_values(["_sort","فروش نیمه دوم"],ascending=[True,True]).drop(columns="_sort")
# Invoice-level product co-occurrence; confidence is P(B|A), support is share of invoices.
_invoice_col=sales["شماره فاکتور"].astype(str).str.strip()
_basket=sales[(sales["نوع سند"]=="فروش")&(_invoice_col!="")].copy(); _basket["_invoice"]=_invoice_col.loc[_basket.index]
_invoice_items=_basket.groupby("_invoice")["کالا"].agg(lambda x: sorted(set(v for v in x if v and v!="نامشخص")))
_pair_counts={}; _product_invoice_counts={}; _invoice_count=len(_invoice_items)
for _items in _invoice_items:
    for _item in _items:_product_invoice_counts[_item]=_product_invoice_counts.get(_item,0)+1
    for _i in range(len(_items)):
        for _j in range(_i+1,len(_items)):
            _pair=(_items[_i],_items[_j]); _pair_counts[_pair]=_pair_counts.get(_pair,0)+1
_basket_rows=[]
for (_a,_b),_cnt in _pair_counts.items():
    _basket_rows.append({"کالای A":_a,"کالای مکمل B":_b,"فاکتورهای مشترک":_cnt,
        "درصد هم‌وقوعی از فاکتورها":100*_cnt/_invoice_count if _invoice_count else 0,
        "شانس خرید B در فاکتورهای A (%)":100*_cnt/_product_invoice_counts[_a] if _product_invoice_counts.get(_a) else 0})
basket=pd.DataFrame(_basket_rows,columns=["کالای A","کالای مکمل B","فاکتورهای مشترک","درصد هم‌وقوعی از فاکتورها","شانس خرید B در فاکتورهای A (%)"])
if not basket.empty:basket=basket.sort_values(["فاکتورهای مشترک","شانس خرید B در فاکتورهای A (%)"],ascending=False)
# Balanced visitor/team quality score, each KPI normalized to the strongest group.
_rep_sales=sales.copy(); _rep_sales["_rep"]=_rep_sales["ویزیتور"].replace("",np.nan).fillna("نامشخص")
if rep_name and "ویزیتور" in sales and sales["ویزیتور"].astype(str).str.strip().ne("").any():
    _known=_rep_sales["_rep"].astype(str).str.contains(re.escape(rep_name),case=False,na=False)
    if _known.any():_rep_sales=_rep_sales[_known]
_rep_rows=[]
for _rep,_group in _rep_sales.groupby("_rep"):
    _ss=_group[_group["نوع سند"]=="فروش"]; _rr=_group[_group["نوع سند"]=="برگشت از فروش"]
    _rep_rows.append({"ویزیتور":_rep,"فروش خالص":float(_ss["مبلغ"].sum()-_rr["مبلغ"].sum()),
        "نرخ مرجوعی (%)":100*float(_rr["مبلغ"].sum())/float(_ss["مبلغ"].sum()) if _ss["مبلغ"].sum()>0 else 0,
        "مشتری فعال":int(_ss["مشتری"].nunique()),"تعداد فاکتور":int(_ss.loc[_ss["شماره فاکتور"].astype(str)!="","شماره فاکتور"].nunique()),
        "تنوع کالا":int(_ss["کالا"].nunique())})
scorecard=pd.DataFrame(_rep_rows)
if not scorecard.empty:
    for _metric in ["فروش خالص","مشتری فعال","تعداد فاکتور","تنوع کالا"]:
        _max=float(scorecard[_metric].max()); scorecard["_"+_metric]=scorecard[_metric].clip(lower=0)/_max*100 if _max>0 else 0
    scorecard["_مرجوعی"]=100-scorecard["نرخ مرجوعی (%)"].clip(0,100)
    scorecard["امتیاز کیفی از ۱۰۰"]=scorecard[["_فروش خالص","_مشتری فعال","_تعداد فاکتور","_تنوع کالا","_مرجوعی"]].mean(axis=1).round(1)
    scorecard["رتبه"]=np.select([scorecard["امتیاز کیفی از ۱۰۰"]>=80,scorecard["امتیاز کیفی از ۱۰۰"]>=60],["عالی","خوب"],default="نیازمند بهبود")
    scorecard=scorecard.drop(columns=[c for c in scorecard if c.startswith("_")]).sort_values("امتیاز کیفی از ۱۰۰",ascending=False)
st.markdown(f'<div class="hero"><h1>📊 داشبورد جامع فروش{(" · "+rep_name) if rep_name else ""}</h1><p>گزارش‌های بارگذاری‌شده: {len(files)} · رکوردهای معتبر: {len(sales):,} · بازه: {dates[0] if dates else "—"} تا {dates[-1] if dates else "—"}</p></div>',unsafe_allow_html=True)
cols=st.columns(5)
for col,label,value,delta in zip(cols,["فروش خالص","فروش ناخالص","مرجوعی ریالی","مشتری فعال","فاکتور"],[fmt(net),fmt(gross),fmt(ret),f"{sales['مشتری'].nunique():,}",f"{invoices:,}"],[f"نرخ مرجوعی {ret_rate:.1f}%","قبل از کسر مرجوعی","مبالغ گزارش‌شده","مشتری یکتا","شماره فاکتور یکتا"]): col.metric(label,value,delta)

# Build each tab exactly once, after its full ordered key list is known.
aux_comm={k:v for k,v in aux.items() if any(x in k.lower() for x in ["porsant","پورسانت","commission"])}
aux_comp={k:v for k,v in aux.items() if any(x in k.lower() for x in ["compare","مقایسه"])}
tab_titles=["داشبورد جامع","تحلیل مشتریان","تحلیل برندها","تحلیل کالاها و مرجوعی‌ها","ریسک ریزش مشتریان","تحلیل سبد خرید و کالاهای مکمل","امتیاز کیفی ویزیتور","هدف ماهانه"]
if aux_comm:tab_titles.append("پورسانت و تخفیفات")
if aux_comp:tab_titles.append("مقایسه ویزیتورها")
tab_titles.append("دانلود گزارش اکسل")
tabs=st.tabs(tab_titles); ti={n:t for n,t in zip(tab_titles,tabs)}
with ti["ریسک ریزش مشتریان"]:
    st.subheader("هشدار افت خرید و ریسک ریزش مشتریان (Churn Risk Alert)")
    st.caption("شناسایی خودکار کاهش فراوانی یا مبلغ خرید در نیمه دوم دوره و عدم سفارش در تاریخ‌های اخیر؛ قابل تنظیم از نوار کناری.")
    if churn.empty: st.info("داده کافی برای تحلیل ریزش وجود ندارد.")
    else:
        a,b,c=st.columns(3); a.metric("ریسک بالا",int((churn['وضعیت']=='ریسک بالا').sum())); b.metric("ریسک متوسط",int((churn['وضعیت']=='متوسط').sum())); c.metric("فعال",int((churn['وضعیت']=='فعال').sum()))
        _badge={"ریسک بالا":"🔴 ریسک بالا","متوسط":"🟠 ریسک متوسط","فعال":"🟢 فعال"}
        _show=churn.copy(); _show["وضعیت"]=_show["وضعیت"].map(_badge)
        st.dataframe(money_columns(_show[_show["وضعیت"]!="🟢 فعال"],["فروش نیمه اول","فروش نیمه دوم"]),use_container_width=True,height=420,hide_index=True)
        st.dataframe(money_columns(churn[churn["وضعیت"]=="فعال"],["فروش نیمه اول","فروش نیمه دوم"]),use_container_width=True,height=260,hide_index=True)
with ti["تحلیل سبد خرید و کالاهای مکمل"]:
    st.subheader("کالاهای مکمل و تحلیل هم‌فاکتوری (Basket / Cross-Selling)")
    st.caption("کالاهایی که در یک شماره فاکتور با هم خریداری می‌شوند؛ برای پیشنهاد فروش مکمل به ویزیتورها.")
    if basket.empty: st.info("ستون شماره فاکتور معتبر یا چند کالا در هر فاکتور یافت نشد.")
    else:
        _bshow=basket.head(30).copy()
        _bshow["پیشنهاد فروش"]=_bshow["کالای A"]+" همراه با "+_bshow["کالای مکمل B"]
        _bshow["درصد هم‌وقوعی از فاکتورها"]=_bshow["درصد هم‌وقوعی از فاکتورها"].round(1); _bshow["شانس خرید B در فاکتورهای A (%)"]=_bshow["شانس خرید B در فاکتورهای A (%)"].round(1)
        st.dataframe(_bshow[["پیشنهاد فروش","کالای A","کالای مکمل B","فاکتورهای مشترک","درصد هم‌وقوعی از فاکتورها","شانس خرید B در فاکتورهای A (%)"]],use_container_width=True,height=480,hide_index=True)
with ti["امتیاز کیفی ویزیتور"]:
    st.subheader("کارت امتیازی کیفی ویزیتور / عملکرد فروش")
    st.caption("ترکیب فروش خالص، نرخ مرجوعی، مشتری فعال، تعداد فاکتور و تنوع سبد کالایی؛ امتیاز نسبی از ۱۰۰ نسبت به بهترین عملکرد گروه.")
    if scorecard.empty: st.info("داده کافی برای امتیازدهی وجود ندارد.")
    else:
        a,b=st.columns(2)
        with a: st.plotly_chart(px.bar(scorecard,x="ویزیتور",y="امتیاز کیفی از ۱۰۰",color="رتبه",title="امتیاز کیفی از ۱۰۰",color_discrete_map={"عالی":"#0f9f8f","خوب":"#2563eb","نیازمند بهبود":"#f97373"}),use_container_width=True)
        with b:
            figq=px.bar(scorecard.melt(id_vars=["ویزیتور"],value_vars=["نرخ مرجوعی (%)","مشتری فعال","تعداد فاکتور","تنوع کالا"]),x="ویزیتور",y="value",color="variable",barmode="group",title="شاخص‌های عملکرد")
            figq.update_layout(template="plotly_white"); st.plotly_chart(figq,use_container_width=True)
        st.dataframe(scorecard.style.format({"فروش خالص":"{:,.0f}","نرخ مرجوعی (%)":"{:.1f}","امتیاز کیفی از ۱۰۰":"{:.1f}"}),use_container_width=True,hide_index=True)
with ti["هدف ماهانه"]:
    st.subheader("سیستم هدف‌گذاری و تحقق تارگت ماهانه")
    _target2=float(monthly_target) if monthly_target>0 else 0.0
    if _target2<=0: st.info("یک تارگت مثبت در نوار کناری وارد کنید.")
    else:
        _pct2=min(100.0,net/_target2*100)
        m1,m2,m3,m4=st.columns(4)
        m1.metric("فروش خالص فعلی",fmt(net)); m2.metric("تارگت ماهانه",fmt(_target2)); m3.metric("درصد تحقق",f"{_pct2:.1f}%"); m4.metric("فاصله تا هدف",fmt(max(_target2-net,0)))
        figg2=go.Figure(go.Indicator(mode="gauge+number",value=_pct2,number=dict(suffix="%",font=dict(size=42)),
        title=dict(text="درصد تحقق تارگت فروش خالص",font=dict(size=15)),
        gauge=dict(axis=dict(range=[0,100]),bar=dict(color="#2563eb"),
        steps=[dict(range=[0,60],color="#fee2e2"),dict(range=[60,90],color="#fef3c7"),dict(range=[90,100],color="#d1fae5")],threshold=dict(line=dict(color="#dc2626",width=3),value=100))))
        figg2.update_layout(height=350,margin=dict(l=25,r=25,t=45,b=20),template="plotly_white")
        st.plotly_chart(figg2,use_container_width=True)
        st.progress(_pct2/100)
with ti["داشبورد جامع"]:
    st.subheader("روند فروش و مرجوعی")
    daily=sales.pivot_table(index="تاریخ",columns="نوع سند",values="مبلغ",aggfunc="sum",fill_value=0).reset_index()
    for c in ["فروش","برگشت از فروش"]:
        if c not in daily:daily[c]=0
    daily["فروش خالص"]=daily["فروش"]-daily["برگشت از فروش"]
    fig=go.Figure()
    fig.add_bar(x=daily["تاریخ"],y=daily["فروش"],name="فروش ناخالص",marker_color="#2563eb")
    fig.add_bar(x=daily["تاریخ"],y=daily["برگشت از فروش"],name="مرجوعی",marker_color="#f97373")
    fig.add_scatter(x=daily["تاریخ"],y=daily["فروش خالص"],name="فروش خالص",mode="lines+markers",line=dict(color="#0f9f8f",width=3))
    fig.update_layout(template="plotly_white",barmode="group",height=390,legend=dict(orientation="h",y=1.12),margin=dict(l=10,r=10,t=25,b=10))
    st.plotly_chart(fig,use_container_width=True)
    a,b=st.columns(2)
    with a: st.markdown("#### 🏆 برندهای پیشرو"); st.dataframe(money_columns(brand[["برند","فروش خالص","درصد سهم"]].head(8),["فروش خالص"]),use_container_width=True,hide_index=True)
    with b: st.markdown("#### 👥 مشتریان برتر"); st.dataframe(money_columns(cust[["مشتری","فروش خالص"]].head(8),["فروش خالص"]),use_container_width=True,hide_index=True)
    _target=float(monthly_target) if monthly_target>0 else 0.0
    _pct=min(100.0,net/_target*100) if _target>0 else 0.0
    figg=go.Figure(go.Indicator(mode="gauge+number+delta",value=_pct,number=dict(suffix="% تحقق",font=dict(size=30)),
        delta=dict(reference=100,suffix="%",increasing=dict(color="#0f9f8f"),decreasing=dict(color="#f97373")),
        title=dict(text=f"تحقق تارگت فروش خالص<br><sub>فروش: {fmt(net)} از {fmt(_target)} · مانده: {fmt(max(_target-net,0))}</sub>",font=dict(size=13)),
        gauge=dict(axis=dict(range=[0,100],tickfont=dict(size=11)),bar=dict(color="#2563eb"),
        steps=[dict(range=[0,60],color="#fee2e2"),dict(range=[60,90],color="#fef3c7"),dict(range=[90,100],color="#d1fae5")],threshold=dict(line=dict(color="#dc2626",width=3),value=100))))
    figg.update_layout(height=270,margin=dict(l=20,r=20,t=40,b=15),template="plotly_white")
    with b: st.plotly_chart(figg,use_container_width=True)
with ti["تحلیل مشتریان"]:
    st.subheader("وفادار، از دست‌رفته و مشتریان جدید")
    prev_df=pd.DataFrame()
    if previous:
        try:
            pr,_=read_upload(previous); prev_df=sales_from_raw(pr,previous.name)
        except Exception as e:st.warning(f"خواندن فایل قبلی ممکن نشد: {e}")
    if previous and not prev_df.empty:
        old=set(prev_df.loc[prev_df["نوع سند"]=="فروش","مشتری"].dropna()); curr=set(sale["مشتری"].dropna())
        new=curr-old; lost=old-curr; loyal=old&curr
        k=st.columns(3); k[0].metric("مشتریان وفادار",len(loyal)); k[1].metric("مشتریان جدید",len(new)); k[2].metric("مشتریان ازدست‌رفته",len(lost))
        cl,cr=st.columns(2)
        with cl: st.write("**مشتریان جدید**");st.dataframe(pd.DataFrame({"مشتری":sorted(new)}),use_container_width=True,hide_index=True)
        with cr: st.write("**مشتریان ازدست‌رفته**");st.dataframe(pd.DataFrame({"مشتری":sorted(lost)}),use_container_width=True,hide_index=True)
        st.caption("وفادار یعنی مشتری دارای فروش در هر دو فایل. دسته‌بندی بر مبنای شناسه/نام مشتری انجام شده است.")
    elif len(dates)>1:
        st.markdown("#### انتخاب دو بازه از فایل جاری")
        x,y=st.columns(2)
        with x: p1=st.multiselect("بازه اول (مقایسه پایه)",dates,default=dates[:max(1,len(dates)//2)],key="p1")
        with y: p2=st.multiselect("بازه دوم (جاری)",dates,default=dates[max(1,len(dates)//2):],key="p2")
        old=set(sales.loc[sales["تاریخ"].isin(p1)&(sales["نوع سند"]=="فروش"),"مشتری"]); curr=set(sales.loc[sales["تاریخ"].isin(p2)&(sales["نوع سند"]=="فروش"),"مشتری"])
        k=st.columns(3); k[0].metric("وفادار",len(old&curr));k[1].metric("جدید",len(curr-old));k[2].metric("از دست‌رفته",len(old-curr))
        st.caption("در صورت نیاز به مقایسه دقیق دو دوره، از گزینه فایل فروش قبلی در نوار کناری استفاده کنید.")
    else:st.info("برای تحلیل وفادار / جدید / ازدست‌رفته، فایل دوره قبل را بارگذاری کنید یا گزارشی با چند تاریخ انتخاب نمایید.")
    st.markdown("#### ارزش خرید و نرخ مرجوعی مشتریان")
    left,right=st.columns(2)
    with left:
        figc=px.bar(cust.head(15).sort_values("فروش خالص"),x="فروش خالص",y="مشتری",orientation="h",color="فروش خالص",color_continuous_scale="Blues",title="۱۵ مشتری برتر بر اساس فروش خالص");figc.update_layout(template="plotly_white",height=480);st.plotly_chart(figc,use_container_width=True)
    with right: st.dataframe(money_columns(cust.sort_values("برگشت از فروش",ascending=False).head(20),["فروش","برگشت از فروش","فروش خالص"]),use_container_width=True,height=480,hide_index=True)
with ti["تحلیل برندها"]:
    st.subheader("سهم برندها از فروش")
    a,b=st.columns([1,1.2]); top=brand[brand["فروش خالص"]>0].head(10)
    with a:
        if not top.empty: st.plotly_chart(px.pie(top,names="برند",values="فروش خالص",hole=.48,title="سهم برندهای برتر"),use_container_width=True)
        else:st.info("داده مثبت برای نمودار سهم وجود ندارد.")
    with b: st.plotly_chart(px.bar(brand.head(15).sort_values("فروش خالص"),x="فروش خالص",y="برند",orientation="h",color="درصد سهم",color_continuous_scale="Teal",title="فروش خالص و سهم برندها"),use_container_width=True)
    st.dataframe(brand.style.format({"فروش":"{:,.0f}","برگشت از فروش":"{:,.0f}","فروش خالص":"{:,.0f}","درصد سهم":"{:.2f}%"}),use_container_width=True,hide_index=True)
with ti["تحلیل کالاها و مرجوعی‌ها"]:
    st.subheader("پرفروش‌ترین و پرمرجوعی‌ترین کالاها")
    a,b=st.columns(2)
    with a:st.markdown("#### پرفروش‌ها (فروش خالص)");st.dataframe(money_columns(prod.head(20),["فروش","برگشت از فروش","فروش خالص"]),use_container_width=True,hide_index=True)
    with b:st.markdown("#### بیشترین مرجوعی ریالی");st.dataframe(money_columns(prod.sort_values("برگشت از فروش",ascending=False).head(20),["فروش","برگشت از فروش","فروش خالص"]),use_container_width=True,hide_index=True)
    qty=sales.pivot_table(index=["کالا","برند"],columns="نوع سند",values="تعداد",aggfunc="sum",fill_value=0).reset_index()
    for c in ["فروش","برگشت از فروش"]:
        if c not in qty:qty[c]=0
    qty["تعداد خالص"]=qty["فروش"]-qty["برگشت از فروش"]
    st.markdown("#### تحلیل مقداری");st.dataframe(qty.sort_values("تعداد خالص",ascending=False).head(30),use_container_width=True,hide_index=True)
if aux_comm:
    with ti["پورسانت و تخفیفات"]:
        st.subheader("گزارش پورسانت / تخفیف")
        for name,df in aux_comm.items():
            st.markdown(f"**{name}**");st.dataframe(df,use_container_width=True,height=430,hide_index=True)
        st.caption("ستون‌ها مطابق گزارش منبع نمایش داده شده‌اند؛ تفسیر محاسبات پورسانت وابسته به ساختار خروجی فایل است.")
if aux_comp:
    with ti["مقایسه ویزیتورها"]:
        st.subheader("مقایسه کالا به تفکیک ویزیتور")
        for name,df in aux_comp.items():
            st.markdown(f"**{name}**");st.dataframe(df,use_container_width=True,height=500,hide_index=True)
with ti["دانلود گزارش اکسل"]:
    st.subheader("دانلود گزارش استاندارد و پاک‌سازی‌شده")
    st.caption("شامل ریز فروش، خلاصه برند، خلاصه مشتری، کالاها و در صورت وجود جداول پورسانت/مقایسه.")
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        sales.to_excel(writer,sheet_name="ریز فروش",index=False)
        brand.to_excel(writer,sheet_name="خلاصه برند",index=False)
        cust.to_excel(writer,sheet_name="خلاصه مشتری",index=False)
        prod.to_excel(writer,sheet_name="خلاصه کالا",index=False)
        for name,df in aux.items():
            sheet=re.sub(r"[\\/*?:\[\]]","_",Path(name).stem)[:31] or "گزارش"
            if sheet in writer.book.sheetnames:sheet=sheet[:26]+"_2"
            df.to_excel(writer,sheet_name=sheet,index=False)
    st.download_button("📥 دانلود گزارش اکسل",data=out.getvalue(),file_name="گزارش_تحلیلی_فروش.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    st.download_button("دانلود ریز فروش CSV",data=sales.to_csv(index=False).encode("utf-8-sig"),file_name="ریز_فروش.csv",mime="text/csv")
