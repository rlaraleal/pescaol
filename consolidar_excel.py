import pandas as pd, re, unicodedata, hashlib, json, sys
from pathlib import Path
from difflib import SequenceMatcher

def txt(v): return '' if pd.isna(v) else re.sub(r'\s+',' ',str(v)).strip()
def rut(v): return re.sub(r'[^0-9Kk]','',txt(v)).upper()
def norm(s):
 s=unicodedata.normalize('NFD',txt(s).upper());s=''.join(c for c in s if unicodedata.category(c)!='Mn')
 s=re.sub(r'\bAV\.?\b','AVENIDA',s);s=re.sub(r'\bPJE\.?\b','PASAJE',s);s=re.sub(r'\bCAM\.?\b','CAMINO',s)
 s=re.sub(r'\bN(?:RO|UMERO)?\.?\s*[°º#]?\s*',' ',s);s=re.sub(r'[^A-Z0-9 ]',' ',s)
 return re.sub(r'\s+',' ',s).strip()
def score(addr,comuna,count,last):
 a=norm(addr); q=0
 q+=25 if re.search(r'\d',a) else 0;q+=12 if any(w in a for w in ['CALLE','AVENIDA','PASAJE','CAMINO','RUTA','KM']) else 0
 q+=15 if norm(comuna) in a else 0;q+=min(25,count*5);q+=10 if len(a)>=10 else 0;q+=min(13,last/86400/365)
 return q
def main(xlsx,outdir='salida_pescaol'):
 d=pd.read_excel(xlsx,engine='openpyxl');out=Path(outdir);out.mkdir(exist_ok=True)
 d['_rut']=d['Rut Destinatario'].map(rut);d['_addr']=d['Dirección de Destino'].map(txt);d['_norm']=d['_addr'].map(norm)
 d['_fecha']=pd.to_datetime(d['Fecha Emision'],errors='coerce');mx=d['_fecha'].max()
 rows=[];variants=[]
 for r,g in d[d['_rut']!=''].groupby('_rut'):
  name=g['Nombre Destinatario'].dropna().astype(str).str.strip().mode();name=name.iloc[0] if len(name) else ''
  comuna=g['Ciudad de Destino'].dropna().astype(str).str.strip().mode();comuna=comuna.iloc[0] if len(comuna) else ''
  cand=[]
  for n,a in g.groupby('_norm'):
   if not n: continue
   original=a['_addr'].mode().iloc[0];last=a['_fecha'].max();cnt=len(a)
   sc=score(original,comuna,cnt,(last-pd.Timestamp('2000-01-01')).total_seconds() if pd.notna(last) else 0)
   cand.append((sc,last,cnt,original,n))
  cand.sort(key=lambda x:(x[0],x[1]),reverse=True)
  best=cand[0] if cand else (0,pd.NaT,0,'','')
  public='pub-'+hashlib.sha256(('PescAOL|'+r).encode()).hexdigest()[:16]
  rows.append({'establecimientoId':public,'nombre':name,'comuna':comuna,'direccionConsolidada':best[3],'direccionNormalizada':best[4],'puntajeDireccion':round(best[0],1),'variantes':len(cand),'estadoGeocodificacion':'PENDIENTE','latitud':'','longitud':'','precision':'','direccionFormateada':'','fechaUltimoMovimiento':g['_fecha'].max()})
  for sc,last,cnt,original,n in cand:variants.append({'establecimientoId':public,'direccionOriginal':original,'direccionNormalizada':n,'frecuencia':cnt,'ultimaFecha':last,'puntaje':round(sc,1),'seleccionada':n==best[4]})
 master=pd.DataFrame(rows);master.to_csv(out/'establecimientos_maestros.csv',index=False,encoding='utf-8-sig')
 pd.DataFrame(variants).to_csv(out/'variantes_direcciones.csv',index=False,encoding='utf-8-sig')
 queue=master[['establecimientoId','nombre','direccionConsolidada','comuna']].copy();queue['consultaGeocodificacion']=queue['direccionConsolidada']+', '+queue['comuna']+', Región de Ñuble, Chile';queue.to_csv(out/'cola_geocodificacion.csv',index=False,encoding='utf-8-sig')
 print(f'{len(master)} destinatarios consolidados en {out}')
if __name__=='__main__':main(sys.argv[1],sys.argv[2] if len(sys.argv)>2 else 'salida_pescaol')
