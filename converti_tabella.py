#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convertitore Tabella Sinottica Widiba -> catalogo.json per la web-app "Tabella Sinottica 2.0".
USO (ufficio):
    pip install pandas openpyxl
    python converti_tabella.py Tabella_Sinottica_Widiba.xlsx catalogo.json
Poi pubblicare 'catalogo.json' SULLO STESSO SITO dell'app (stesso dominio) e
impostarne l'indirizzo nell'app, sezione Dati -> Sorgente condivisa.
"""
import sys, json, datetime
import pandas as pd, numpy as np, openpyxl

def main(src, dst):
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    ws = wb["Catalogo"]
    r1 = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    snap = next((v.strftime("%Y-%m-%d") for v in r1 if isinstance(v, datetime.datetime)), None)

    df = pd.read_excel(src, sheet_name="Catalogo", header=1)
    df = df[df.iloc[:,0].notna()].reset_index(drop=True)

    def norm_status(x):
        if not isinstance(x,str): return "Altro"
        s=x.strip(); return {"Full Closure":"Full closure","Soft Closure":"Soft closure"}.get(s,s)
    def norm_distrib(x):
        if not isinstance(x,str): return None
        s=x.strip().upper().replace(" ","")
        if s.startswith("ACC"): return "ACC"
        if s=="DIS": return "DIS"
        for suf in ["A","Q","M","H"]:
            if s.startswith("DIS") and s.endswith(suf): return "DIS-"+suf
        return "DIS" if s.startswith("DIS") else x.strip()
    def numcol(i): return pd.to_numeric(df.iloc[:,i],errors="coerce").to_numpy()
    def cleanv(v):
        if v is None or (isinstance(v,float) and np.isnan(v)): return None
        if isinstance(v,float): return round(v) if abs(v-round(v))<1e-9 else round(v,4)
        return v
    def numlist(i,mult=1.0):
        a=numcol(i); return [cleanv(float(x)*mult) if not np.isnan(x) else None for x in a]
    strip=lambda s: s.str.strip()

    status=df.iloc[:,1].map(norm_status); distrib=df.iloc[:,10].map(norm_distrib)
    fac_src={"casa":strip(df.iloc[:,3].astype("string")),"macro":strip(df.iloc[:,12].astype("string")),
     "wise":strip(df.iloc[:,13].astype("string")),"ms":strip(df.iloc[:,14].astype("string")),
     "valuta":strip(df.iloc[:,7].astype("string")),"rischio":strip(df.iloc[:,17].astype("string")),
     "classe":strip(df.iloc[:,6].astype("string")),"status":status,"distrib":distrib}
    dicts={}; codes={}
    for k,s in fac_src.items():
        cats=sorted([c for c in s.dropna().unique()]); idx={c:i for i,c in enumerate(cats)}
        dicts[k]=cats; codes[k]=[idx[v] if isinstance(v,str) and v in idx else None for v in s]
    txt=lambda i:[v.strip() if isinstance(v,str) else None for v in df.iloc[:,i]]
    gamma,nome,kid=txt(4),txt(5),txt(64)
    anno=[v.year if isinstance(v,(datetime.datetime,datetime.date)) else None for v in df.iloc[:,11]]
    cols={'ratingMS':numlist(15),'esg':numlist(16),'omdF':numlist(18),'omdP':numlist(19),
     'minIniz':numlist(28),'commSott':numlist(32,100),'commGest':numlist(33,100),'payUp':numlist(38,100),
     'payCont':numlist(39,100),'ratingFZ':numlist(40),'p1m':numlist(43),'p3m':numlist(44),'p6m':numlist(45),
     'p1y':numlist(46),'p3y':numlist(47),'p5y':numlist(48),'vol1y':numlist(49),'vol3y':numlist(50),
     'sharpe1y':numlist(51),'sharpe3y':numlist(52),'ir1y':numlist(53),'alpha1y':numlist(55),'tev1y':numlist(57),
     'maxdd':numlist(59),'recov':numlist(60),'distR1y':numlist(61),'cedola':numlist(62)}
    isin=[cleanv(v) for v in df.iloc[:,0]]
    schema=["isin","status","casa","gamma","nome","classe","valuta","distrib","anno","macro","wise","ms",
     "ratingMS","esg","rischio","omdF","omdP","minIniz","commSott","commGest","payUp","payCont","ratingFZ",
     "p1m","p3m","p6m","p1y","p3y","p5y","vol1y","vol3y","sharpe1y","sharpe3y","ir1y","alpha1y","tev1y",
     "maxdd","recov","distR1y","cedola","kid"]
    N=len(df); rows=[]
    for r in range(N):
        rows.append([isin[r],codes["status"][r],codes["casa"][r],gamma[r],nome[r],codes["classe"][r],
         codes["valuta"][r],codes["distrib"][r],anno[r],codes["macro"][r],codes["wise"][r],codes["ms"][r],
         cols['ratingMS'][r],cols['esg'][r],codes["rischio"][r],cols['omdF'][r],cols['omdP'][r],cols['minIniz'][r],
         cols['commSott'][r],cols['commGest'][r],cols['payUp'][r],cols['payCont'][r],cols['ratingFZ'][r],
         cols['p1m'][r],cols['p3m'][r],cols['p6m'][r],cols['p1y'][r],cols['p3y'][r],cols['p5y'][r],cols['vol1y'][r],
         cols['vol3y'][r],cols['sharpe1y'][r],cols['sharpe3y'][r],cols['ir1y'][r],cols['alpha1y'][r],cols['tev1y'][r],
         cols['maxdd'][r],cols['recov'][r],cols['distR1y'][r],cols['cedola'][r],kid[r]])
    payload={"v":snap,"schema":schema,"dicts":dicts,"rows":rows,"n":N}
    with open(dst,"w",encoding="utf-8") as f:
        json.dump(payload,f,separators=(",",":"),ensure_ascii=False)
    print(f"OK · {N} fondi · versione {snap} · scritto {dst}")

if __name__=="__main__":
    if len(sys.argv)<2:
        print("Uso: python converti_tabella.py <input.xlsx> [output.json]"); sys.exit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv)>2 else "catalogo.json")
