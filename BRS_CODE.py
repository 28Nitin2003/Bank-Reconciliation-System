import streamlit as st
import os, re, warnings
import pandas as pd
from io import BytesIO
from difflib import SequenceMatcher
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
warnings.filterwarnings("ignore")

# ----------------------------------------------------
# Streamlit Page Setup
# ----------------------------------------------------
st.set_page_config(page_title="Bank Reconciliation System", page_icon="🏦", layout="wide")

# ----------------------------------------------------
# 🌊 Deep Navy & Aqua Accent Theme
# ----------------------------------------------------
st.markdown("""
<style>
html, body, .stApp {
    background-color:#f8fafc;
    font-family:"Segoe UI",sans-serif;
    color:#1e293b;
}

/* Header */
.app-header {
    background:linear-gradient(90deg,#06b6d4,#0891b2);
    padding:26px;border-radius:14px;text-align:center;
    box-shadow:0 3px 10px rgba(0,0,0,0.1);
}
.app-header h2 {color:#ffffff;margin:0;font-weight:600;}
.app-header p {color:#e0f7fa;margin-top:6px;font-size:15px;}

/* Sidebar – deep navy gradient */
section[data-testid="stSidebar"] {
    background:linear-gradient(180deg,#0f172a 0%,#1e293b 100%);
    box-shadow:2px 0 8px rgba(0,0,0,0.25);
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p {color:#f8fafc !important;}
section[data-testid="stSidebar"] .stMarkdown {color:#f1f5f9 !important;}

/* File uploaders */
[data-testid="stFileUploader"] {
    background-color:#ffffff;
    border:2px dashed #06b6d4;
    border-radius:10px;
    padding:16px;
    box-shadow:0 1px 5px rgba(0,0,0,0.1);
}

/* Buttons */
div.stButton > button, .stDownloadButton>button {
    background:linear-gradient(90deg,#06b6d4,#0891b2);
    color:white;font-weight:600;border:none;
    border-radius:8px;padding:10px 22px;
    transition:all .3s ease;font-size:15px;
}
div.stButton > button:hover, .stDownloadButton>button:hover {
    background:linear-gradient(90deg,#0e7490,#06b6d4);
    transform:translateY(-2px);
}

/* Metrics */
[data-testid="stMetricValue"] {color:#0891b2 !important;font-weight:700;}
[data-testid="stMetricLabel"] {color:#475569 !important;}

/* Info & Success boxes */
.stInfo, .stSuccess {
    border:none;border-radius:8px;padding:12px 15px;
}
.stInfo {background:#e0f2f1;color:#004d40;}
.stSuccess {background:#d1fae5;color:#065f46;}

/* DataFrame */
.stDataFrame {
    border:1px solid #e2e8f0;
    border-radius:10px;
    background:white;
    box-shadow:0 1px 6px rgba(0,0,0,0.06);
}

/* Legend */
.legend-box {
    background:#e0f2f1;
    padding:10px 15px;
    border-radius:8px;
    margin-top:10px;
    color:#134e4a;
    font-size:14px;
}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# Header
# ----------------------------------------------------
st.markdown("""
<div class='app-header'>
  <h2>🏦 Bank Reconciliation System</h2>
  <p>Compare SAP and Bank data — generate clean Excel reconciliation reports instantly</p>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# Sidebar Upload Section
# ----------------------------------------------------
st.sidebar.header("📂 Upload Your Files")
bank_file = st.sidebar.file_uploader("🏦 Upload Bank Excel File", type=["xlsx","xls","csv"])
bank_sheet = st.sidebar.text_input("📄 Bank File Sheet Name (optional)")
sap_file  = st.sidebar.file_uploader("💼 Upload SAP Excel File",  type=["xlsx","xls","csv"])

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Matching Settings")
fuzzy_threshold = st.sidebar.slider("Fuzzy Match Threshold (%)",50,100,60,5)

os.makedirs("uploaded_files",exist_ok=True)
sap_path=bank_path=None
if sap_file:
    sap_path="uploaded_files/SAP.xlsx"
    with open(sap_path,"wb") as f:f.write(sap_file.getbuffer())
    st.sidebar.success("✅ SAP File Uploaded")
if bank_file:
    bank_path="uploaded_files/Bank.xlsx"
    with open(bank_path,"wb") as f:f.write(bank_file.getbuffer())
    st.sidebar.success("✅ Bank File Uploaded")

# ----------------------------------------------------
# Account Type
# ----------------------------------------------------
st.markdown("### 🏦 Select Account Type")
acct_type=st.selectbox("Account Type",["G/L Account","BRS Account"],index=1)

# ----------------------------------------------------
# Processor Class
# ----------------------------------------------------
class Processor:
    def __init__(self,b,s,f,a):
        self.b,self.s,self.f,self.a=b,s,f,a

    def clean(self,t):
        return "" if pd.isna(t) else re.sub(r"[^A-Z\s]","",str(t).upper())

    def load_bank(self,sh=None):
        if not self.b: return False
        data=pd.read_excel(self.b,header=None,sheet_name=sh or None)
        if isinstance(data,dict):
            frames=[]
            for nm,df in data.items():
                frames.append(self._prep(df))
            self.df=pd.concat(frames,ignore_index=True)
        else:
            self.df=self._prep(data)
        return True

    def _prep(self,df):
        header_idx=None
        for i in range(len(df)):
            vals=df.iloc[i,0:4].astype(str).str.lower()
            if any(v in["date","txn date","transaction date"] for v in vals):
                header_idx=i;break
        if header_idx is None: return pd.DataFrame()
        df.columns=df.iloc[header_idx]
        df=df[header_idx+1:]
        df.columns=[str(c).title() for c in df.columns]
        for w in["Withdrawal","Withdrawals","Debit","Dr Amount"]:
            if w in df.columns:
                df.rename(columns={w:"Withdrawals"},inplace=True)
        df=df[df["Withdrawals"].notna()]
        return df

    def load_sap(self):
        self.df2=pd.read_excel(self.s)
        col="Amount in LC" if self.a=="BRS Account" else "Amount in Local Currency"
        if col not in self.df2.columns:
            st.error(f"❌ Missing column '{col}' in SAP file.");return False
        self.df2[col]=pd.to_numeric(self.df2[col],errors="coerce")
        return True

    def match(self):
        bank=self.df.copy(); sap=self.df2.copy()
        bank["Withdrawals"]=pd.to_numeric(bank["Withdrawals"],errors="coerce")
        col="Amount in LC" if self.a=="BRS Account" else "Amount in Local Currency"
        sap["status"]="Not Matched"
        for i,row in sap.iterrows():
            amt=row[col]
            if pd.isna(amt): continue
            match_rows=bank[bank["Withdrawals"]==amt]
            if len(match_rows)==1:
                sap.at[i,"status"]="100% Matched"
            elif len(match_rows)>1:
                sap.at[i,"status"]="Multiple Matches"
        self.final=sap

    def excel(self):
        buf=BytesIO()
        self.final.to_excel(buf,index=False)
        buf.seek(0)
        wb=load_workbook(buf)
        ws=wb.active
        colors={"100%":"90EE90","Multiple":"FFB347","Not":"FF9999"}
        col_idx=list(ws[1]).index([c for c in ws[1] if c.value=="status"][0])+1
        for r in range(2,ws.max_row+1):
            val=str(ws.cell(r,col_idx).value or "").lower()
            for k,v in colors.items():
                if k.lower() in val:
                    ws.cell(r,col_idx).fill=PatternFill(start_color=v,end_color=v,fill_type="solid")
        out=BytesIO();wb.save(out);out.seek(0)
        return out

# ----------------------------------------------------
# Processing Section
# ----------------------------------------------------
st.markdown("### 🚀 Process Files")

if st.button("Generate Final Statement"):
    if not sap_file or not bank_file:
        st.error("❌ Please upload both SAP and Bank files.")
    else:
        with st.spinner("Processing your files..."):
            p=Processor(bank_path,sap_path,fuzzy_threshold,acct_type)
            if p.load_bank(bank_sheet) and p.load_sap():
                p.match()
                excel_data=p.excel()

                st.success("✅ Final statement generated successfully!")
                matched=(p.final["status"]=="100% Matched").sum()
                multiple=(p.final["status"]=="Multiple Matches").sum()
                unmatched=(p.final["status"]=="Not Matched").sum()

                col1,col2,col3=st.columns(3)
                col1.metric("✅ 100% Matched",matched)
                col2.metric("🟠 Multiple",multiple)
                col3.metric("❌ Not Matched",unmatched)

                st.markdown("### 📊 Preview (First 10 Rows)")
                st.dataframe(p.final.head(10))

                st.download_button(
                    label="⬇️ Download Final Statement",
                    data=excel_data,
                    file_name=f"{'BRS' if acct_type=='BRS Account' else 'GL'}_Match_Status.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.markdown("""
                <div class="legend-box">
                <b>Legend:</b> 🟩 Matched (Green) &nbsp;&nbsp; 🟧 Multiple Matches (Orange) &nbsp;&nbsp; 🟥 Not Matched (Red)
                </div>
                """, unsafe_allow_html=True)
