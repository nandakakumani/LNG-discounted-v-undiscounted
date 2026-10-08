from io import BytesIO
import re
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DEFAULT_YEARS=list(range(2026,2033))
DIV={"NET GAS (TBTU)":1e6,"NET OIL (MM BBL)":1e6,"HH":1e6}
REQ={"PnL Year","Delta","Final Index","Group","ViewGroup"}

def clean(x): return "" if pd.isna(x) else re.sub(r"\s+"," ",str(x).strip()).upper()
def years(s):
    n=pd.to_numeric(s,errors="coerce"); out=pd.Series(pd.NA,index=s.index,dtype="Int64")
    y=n.between(1900,2200); out.loc[y]=n.loc[y].astype("Int64")
    serial=n.between(20000,100000)&~y
    out.loc[serial]=pd.to_datetime(n.loc[serial],unit="D",origin="1899-12-30",errors="coerce").dt.year.astype("Int64")
    rem=out.isna(); out.loc[rem]=pd.to_datetime(s.loc[rem],errors="coerce").dt.year.astype("Int64")
    return out

def classify(r):
    g,i=clean(r.get("Group")),clean(r.get("Final Index"))
    if i=="HH" or i.startswith("HENRY HUB"): return "HH"
    if "EUA" in g or "EUA" in i or "FREIGHT" in g or "FREIGHT" in i: return "EXCLUDED"
    if g=="OIL" or any(x in i for x in ["BRENT","JCC","DUBAI","WTI"]): return "NET OIL (MM BBL)"
    if g in ["NATURAL GAS & LNG","UNSOLD HOUSE CURVES"] or "LNG" in g or "GAS" in g or i in ["TTF","NBP","PEG","PVB","THE","TVB","ZTP","JKM","PEG - DNK","ZTP - DNK"] or "LNG" in i: return "NET GAS (TBTU)"
    return "UNCLASSIFIED"

def load(file):
    x=pd.ExcelFile(file,engine="openpyxl"); names=[n for n in x.sheet_names if n.strip().casefold()=="raw data"]
    if not names: raise ValueError('Workbook must contain "Raw Data".')
    d=pd.read_excel(x,sheet_name=names[0]); d.columns=[str(c).strip() for c in d.columns]
    miss=REQ-set(d.columns)
    if miss: raise ValueError("Missing columns: "+", ".join(sorted(miss)))
    d["Delta"]=pd.to_numeric(d["Delta"],errors="coerce").fillna(0); d["IFRS Year"]=years(d["PnL Year"])
    d["Category"]=d.apply(classify,axis=1); d["View"]=d["ViewGroup"].map(clean); return d

def table(d,view,ys):
    v=d[d.View.isin([view.upper(),"BOTH"])]
    rows=[]
    for cat in DIV:
        vals=[v.loc[(v.Category==cat)&(v["IFRS Year"]==y),"Delta"].sum()/DIV[cat] for y in ys]
        rows.append([cat]+vals+[sum(vals)])
    return pd.DataFrame(rows,columns=["LNG PORTFOLIO EXPOSURES"]+[f"IFRS {y}" for y in ys]+["NET"]).set_index("LNG PORTFOLIO EXPOSURES")

def section(ws,row,title,t,ys,refs=None):
    thin=Side(style="thin",color="000000"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.merge_cells(start_row=row,start_column=3,end_row=row,end_column=12); c=ws.cell(row,3,title); c.font=Font(name="Arial",size=12,bold=True); c.alignment=Alignment(horizontal="center"); c.fill=PatternFill("solid",fgColor="F2F2F2")
    for col in range(3,13): ws.cell(row,col).border=Border(top=thin,bottom=thin)
    hr=row+2
    for col,val in zip(range(3,4+len(ys)),["LNG PORTFOLIO EXPOSURES"]+[f"IFRS {y}" for y in ys]):
        x=ws.cell(hr,col,val); x.font=Font(name="Arial",size=11,bold=True); x.alignment=Alignment(horizontal="center"); x.border=border
    x=ws.cell(hr,12,"NET"); x.font=Font(bold=True); x.alignment=Alignment(horizontal="center"); x.border=border
    mp={}
    for i,label in enumerate(t.index):
        r=hr+1+i; mp[label]=r; x=ws.cell(r,3,label); x.font=Font(bold=True); x.alignment=Alignment(horizontal="center"); x.border=border
        for j,col in enumerate(range(4,4+len(ys))):
            x=ws.cell(r,col); x.value=f"={get_column_letter(col)}{refs[label][0]}-{get_column_letter(col)}{refs[label][1]}" if refs else float(t.iloc[i,j]); x.number_format='0.00;[Red]-0.00;0.00'; x.alignment=Alignment(horizontal="center"); x.border=border
        x=ws.cell(r,12); x.value=f"=SUM(D{r}:{get_column_letter(3+len(ys))}{r})" if refs else float(t.iloc[i,-1]); x.font=Font(bold=True); x.number_format='0.00;[Red]-0.00;0.00'; x.alignment=Alignment(horizontal="center"); x.border=border
    return mp

def add_sheet(wb,name,df):
    ws=wb.create_sheet(name); ws.sheet_view.showGridLines=False
    for c,n in enumerate(df.columns,1):
        x=ws.cell(1,c,str(n)); x.font=Font(bold=True,color="FFFFFF"); x.fill=PatternFill("solid",fgColor="1F4E78")
    for r,vals in enumerate(df.itertuples(index=False,name=None),2):
        for c,v in enumerate(vals,1):
            if isinstance(v,np.generic): v=v.item()
            ws.cell(r,c,None if pd.isna(v) else v)
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for c,n in enumerate(df.columns,1): ws.column_dimensions[get_column_letter(c)].width=min(35,max(12,len(str(n))+2))

def output(d,u,di,imp,ys):
    wb=Workbook(); ws=wb.active; ws.title="Summary"; ws.column_dimensions["C"].width=34
    for c in "DEFGHIJ": ws.column_dimensions[c].width=12.5
    ws.column_dimensions["K"].width=8; ws.column_dimensions["L"].width=12
    ur=section(ws,2,"OVERALL PORTFOLIO BALANCE BASE CASE (UNDISCOUNTED)",u,ys); dr=section(ws,9,"OVERALL PORTFOLIO BALANCE BASE CASE (DISCOUNTED)",di,ys)
    section(ws,16,"DISCOUNTING IMPACT ON HEDGES",imp,ys,{k:(dr[k],ur[k]) for k in ur})
    inc=d[d.Category.isin(DIV)]
    parts=[]
    for view in ["Undiscounted","Discounted"]:
        z=inc[inc.View.isin([view.upper(),"BOTH"])].groupby(["Category","Final Index","IFRS Year"],dropna=False).Delta.sum().reset_index(); z["Exposure"]=z.apply(lambda r:r.Delta/DIV[r.Category],axis=1); z.insert(0,"View",view); parts.append(z.drop(columns="Delta"))
    add_sheet(wb,"Product Detail",pd.concat(parts,ignore_index=True)); add_sheet(wb,"Classification Audit",d.groupby(["Category","Group","Final Index"],dropna=False).agg(Rows=("Delta","size"),Raw_Delta=("Delta","sum")).reset_index()); add_sheet(wb,"Filtered Raw Data",d)
    wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; out=BytesIO(); wb.save(out); return out.getvalue()

def app():
    st.set_page_config(page_title="LNG IFRS Exposure Builder",layout="wide"); st.title("LNG IFRS Exposure Builder")
    f=st.file_uploader('Upload workbook containing the "Raw Data" tab',type=["xlsx"])
    if not f:return
    try:d=load(f)
    except Exception as e:st.error(str(e));return
    with st.sidebar:
        st.header("Controls")
        opts=sorted(int(x) for x in d["IFRS Year"].dropna().unique() if 2026<=int(x)<=2032); ys=sorted(st.multiselect("IFRS years",opts,default=[y for y in DEFAULT_YEARS if y in opts]))
        po=sorted(d.Portfolio.dropna().astype(str).unique()) if "Portfolio" in d else []; ps=st.multiselect("Portfolios",po,default=po)
        bo=sorted(d.Book.dropna().astype(str).unique()) if "Book" in d else []; bs=st.multiselect("Books",bo,default=bo)
        with st.expander("Unit divisors"):
            DIV["NET GAS (TBTU)"]=st.number_input("Gas divisor",1.0,value=float(DIV["NET GAS (TBTU)"]),step=100000.0); DIV["NET OIL (MM BBL)"]=st.number_input("Oil divisor",1.0,value=float(DIV["NET OIL (MM BBL)"]),step=100000.0); DIV["HH"]=st.number_input("HH divisor",1.0,value=float(DIV["HH"]),step=100000.0)
    if not ys or (po and not ps) or (bo and not bs): st.warning("Select at least one IFRS year, portfolio and book.");return
    z=d[d["IFRS Year"].isin(ys)].copy()
    if po:z=z[z.Portfolio.astype(str).isin(ps)]
    if bo:z=z[z.Book.astype(str).isin(bs)]
    u=table(z,"Undiscounted",ys); di=table(z,"Discounted",ys); imp=di-u
    for title,t in [("Undiscounted",u),("Discounted",di),("Discounting impact",imp)]: st.subheader(title); st.dataframe(t.style.format("{:.2f}"),use_container_width=True)
    st.download_button("Download formatted Excel output",output(z,u,di,imp,ys),"IFRS_Exposure_Output.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
if __name__=="__main__":app()
