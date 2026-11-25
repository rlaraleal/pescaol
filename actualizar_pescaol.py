#!/usr/bin/env python3
import json,re,hashlib,subprocess,urllib.parse,unicodedata
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from difflib import SequenceMatcher
import pandas as pd
ROOT=Path(__file__).resolve().parent; XLSX=ROOT/'datos'/'aol_actual.xlsx'; INDEX=ROOT/'index.html'
CENTERS={'Chillán':(-36.6066,-72.1034),'Chillán Viejo':(-36.6227,-72.1310),'Bulnes':(-36.7422,-72.2986),'Yungay':(-37.1211,-72.0137),'Quillón':(-36.7388,-72.4722),'Coelemu':(-36.4874,-72.7020),'Quirihue':(-36.2804,-72.5412),'San Carlos':(-36.4246,-71.9580),'Coihueco':(-36.6271,-71.8308),'Pinto':(-36.7042,-71.8920),'El Carmen':(-36.8995,-72.0252),'Pemuco':(-36.9770,-72.0988),'San Ignacio':(-36.7997,-72.0303),'Ránquil':(-36.6508,-72.5490),'Trehuaco':(-36.4309,-72.6682),'Cobquecura':(-36.1332,-72.7917),'Portezuelo':(-36.5355,-72.4280),'Ninhue':(-36.3938,-72.3970),'San Fabián':(-36.5561,-71.5499),'San Nicolás':(-36.4990,-72.2133),'Ñiquén':(-36.2860,-71.9010)}
def clean(v): return '' if pd.isna(v) else re.sub(r'\s+',' ',str(v)).strip()
def norm_rut(v): return re.sub(r'[^0-9K]','',clean(v).upper())
def norm(v):
 s=''.join(c for c in unicodedata.normalize('NFD',clean(v).upper()) if unicodedata.category(c)!='Mn')
 s=re.sub(r'\bAVDA?\.?\b','AVENIDA',s);s=re.sub(r'\bPJE\.?\b','PASAJE',s);s=re.sub(r'\bCAM\.?\b','CAMINO',s);s=re.sub(r'\bN[°ºO]?\.?\s*','N ',s)
 return re.sub(r'[^A-Z0-9]+',' ',s).strip()
def show_addr(v): return norm(v).title().replace(' N ',' N° ')
def iso(v):
 d=pd.to_datetime(v,errors='coerce');return '' if pd.isna(d) else d.strftime('%Y-%m-%dT%H:%M:%S')
def offset(k):
 n=int(hashlib.sha256(k.encode()).hexdigest()[:12],16);return ((n%101)-50)/100000,(((n//101)%101)-50)/100000
def data_block(html):
 m=re.search(r'\bconst\s+DATA\s*=\s*',html)
 if not m:return [],None,None
 obj,n=json.JSONDecoder().raw_decode(html[m.end():]);return obj,m.end(),m.end()+n
def find_sheets():
 book=pd.ExcelFile(XLSX,engine='openpyxl')
 data_sheet=None;links_sheet=None
 required={'Fecha Emision','Rut Destinatario','Nombre Destinatario','Dirección de Destino'}
 for sh in book.sheet_names:
  probe=pd.read_excel(XLSX,sheet_name=sh,engine='openpyxl',nrows=5)
  if required.issubset(set(probe.columns)): data_sheet=sh
  raw2=pd.read_excel(XLSX,sheet_name=sh,header=None,engine='openpyxl',usecols='A:B',nrows=700)
  if raw2.apply(lambda col: col.astype(str).str.startswith(('http://','https://')).any()).any(): links_sheet=sh
 if not data_sheet: raise SystemExit('ERROR: no se encontró la hoja con datos AOL')
 if not links_sheet: raise SystemExit('ERROR: no se encontró la hoja con enlaces de Google Maps')
 print(f'OK: hoja AOL={data_sheet}; hoja Maps={links_sheet}')
 return data_sheet,links_sheet

def links_from_pivot(raw,links_sheet):
 piv=pd.read_excel(XLSX,sheet_name=links_sheet,header=None,engine='openpyxl',usecols='A:B')
 names={norm(x) for x in raw['Nombre Destinatario'].dropna()};addrs={norm(x) for x in raw['Dirección de Destino'].dropna()};bn={};ba={}
 for _,r in piv.iterrows():
  label=norm(r.iloc[0]);url=clean(r.iloc[1])
  if not label or not url.startswith('http'):continue
  if label in addrs or any(c.isdigit() for c in label):ba[label]=url;continue
  if label in names:bn[label]=url;continue
  best=max(names,key=lambda x:SequenceMatcher(None,label,x).ratio(),default='')
  if best and SequenceMatcher(None,label,best).ratio()>=.88:bn[best]=url
 return bn,ba
def extract(text):
 text=urllib.parse.unquote(str(text or '')).replace('\\u003d','=').replace('\\u0026','&').replace('\\/','/')
 pats=[r'@(-?\d{1,2}\.\d+),(-?\d{1,3}\.\d+)',r'!3d(-?\d{1,2}\.\d+)!4d(-?\d{1,3}\.\d+)',r'[?&](?:q|query|center)=(-?\d{1,2}\.\d+)[,%]2?C?\s*(-?\d{1,3}\.\d+)',r'"latitude"\s*:\s*(-?\d{1,2}\.\d+).*?"longitude"\s*:\s*(-?\d{1,3}\.\d+)']
 for pat in pats:
  for m in re.finditer(pat,text,re.I|re.S):
   lat,lng=map(float,m.groups())
   if -38.8<lat<-35 and -73.5<lng<-70.5:return lat,lng
 return None
CACHE={}
def resolve(url):
 if not url:return None
 if url in CACHE:return CACHE[url]
 try:
  final=subprocess.check_output(['curl','-A','Mozilla/5.0','-LsS','--retry','3','--retry-delay','2','--max-time','40','-o','/dev/null','-w','%{url_effective}',url],text=True).strip()
  c=extract(final)
  if not c:
   body=subprocess.check_output(['curl','-A','Mozilla/5.0','-LsS','--retry','3','--retry-delay','2','--max-time','40',url],text=True,errors='ignore')
   c=extract(body)
  CACHE[url]=c;return c
 except Exception as e:print('AVISO:',url,e);CACHE[url]=None;return None
def main():
 if not XLSX.exists():raise SystemExit('ERROR: falta datos/aol_actual.xlsx')
 data_sheet,links_sheet=find_sheets()
 raw=pd.read_excel(XLSX,sheet_name=data_sheet,engine='openpyxl')
 sender='Rut_Remitente_Visa'
 if sender not in raw.columns:raise SystemExit('ERROR: falta Rut_Remitente_Visa')
 senders=set(raw[sender].map(norm_rut));senders.discard('')
 df=raw[raw['Region Destino'].map(clean).str.upper().str.contains(r'XVI|NUBLE|ÑUBLE',regex=True,na=False)].copy()
 if 'Estado_Visacion' in df:df=df[~df['Estado_Visacion'].map(clean).str.upper().str.contains('ANULADA',na=False)].copy()
 df['_fecha']=pd.to_datetime(df['Fecha Emision'],errors='coerce');df=df[df['_fecha'].notna()];maxdate=df['_fecha'].max();df=df[df['_fecha']>=maxdate-pd.Timedelta(days=15)]
 before=len(df);df=df[~df['Rut Destinatario'].map(norm_rut).isin(senders)].copy();excluded=before-len(df)
 byname,byaddr=links_from_pivot(raw,links_sheet);html=INDEX.read_text(encoding='utf-8');old,start,end=data_block(html)
 if start is None:raise SystemExit('ERROR: no se encontró const DATA=')
 saved={norm_rut(p.get('rut'))+'|'+norm(p.get('comuna')):(p.get('lat'),p.get('lng'),p.get('precision',''),p.get('mapsUrl','')) for p in old}
 places={}
 for _,r in df.iterrows():
  rk=norm_rut(r['Rut Destinatario']);com=clean(r['Ciudad de Destino']);nk=norm(r['Nombre Destinatario']);ak=norm(r['Dirección de Destino']);key=rk+'|'+norm(com)
  if not rk:continue
  p=places.setdefault(key,{'id':key,'rut':clean(r['Rut Destinatario']),'nombre':clean(r['Nombre Destinatario']),'comuna':com,'direccion':show_addr(r['Dirección de Destino']),'items':[],'_a':{},'_n':nk})
  if ak:p['_a'][ak]=p['_a'].get(ak,0)+1
  q=pd.to_numeric(r['Toneladas Visadas'],errors='coerce')
  p['items'].append({'emision':iso(r['Fecha Emision']),'tipo':clean(r['Producto/Recurso']),'especie':clean(r['Especie Recurso o Producto']),'descripcion':clean(r.get('Nombre Recurso o Producto','')),'ton':0 if pd.isna(q) else float(q),'objeto':clean(r.get('Objeto del Destino','')),'estado':clean(r.get('Estado_Visacion',''))})
 result=[];unresolved=[];verified=centers=0
 for p in places.values():
  best=max(p['_a'],key=p['_a'].get) if p['_a'] else ''
  if best:p['direccion']=show_addr(best)
  url=byaddr.get(best) or byname.get(p['_n'],'');coord=resolve(url) if url else None
  if coord:p['lat'],p['lng']=coord;p['precision']='Google Maps verificado';p['mapsUrl']=url;verified+=1
  elif url:unresolved.append((p['nombre'],p['comuna'],p['direccion'],url))
  else:
   lat,lng=CENTERS.get(p['comuna'],(-36.6066,-72.1034));dy,dx=offset(p['id']);p['lat'],p['lng']=lat+dy,lng+dx;p['precision']='Centro comunal aproximado';p['mapsUrl']='';centers+=1
  p.pop('_a',None);p.pop('_n',None);result.append(p)
 if unresolved:
  print('ERROR: enlaces Maps no resueltos:')
  for x in unresolved:print(' | '.join(x))
  raise SystemExit(f'ERROR: {len(unresolved)} enlaces sin coordenadas. No se publicó para evitar puntos incorrectos.')
 html=html[:start]+json.dumps(result,ensure_ascii=False,separators=(',',':'))+html[end:]
 stamp=maxdate.strftime('%Y-%m-%dT%H:%M:%S');html,n=re.subn(r"const\s+MAXDATE\s*=\s*new Date\('[^']*'\)",f"const MAXDATE=new Date('{stamp}')",html,count=1)
 html=re.sub(r"const\s+SOURCE_FILE\s*=\s*'[^']*'","const SOURCE_FILE='aol_actual.xlsx'",html,count=1);now=datetime.now(ZoneInfo('America/Santiago')).strftime('%d-%m-%Y %H:%M');html=re.sub(r"const\s+SOURCE_UPDATED\s*=\s*'[^']*'",f"const SOURCE_UPDATED='{now}'",html,count=1)
 INDEX.write_text(html,encoding='utf-8');print(f'OK: {excluded} filas remitentes excluidas; {len(result)} locales; {verified} Maps; {centers} centro comunal')
if __name__=='__main__':main()
