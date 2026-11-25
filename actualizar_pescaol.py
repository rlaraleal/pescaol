#!/usr/bin/env python3
import json,re,hashlib,subprocess,urllib.parse,unicodedata
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from difflib import SequenceMatcher
import pandas as pd
ROOT=Path(__file__).resolve().parent; XLSX=ROOT/'datos'/'aol_actual.xlsx'; INDEX=ROOT/'index.html'
CENTERS={'Chillán':(-36.6066,-72.1034),'Chillán Viejo':(-36.6227,-72.1310),'Bulnes':(-36.7422,-72.2986),'Yungay':(-37.1211,-72.0137),'Quillón':(-36.7388,-72.4722),'Coelemu':(-36.4874,-72.7020),'Quirihue':(-36.2804,-72.5412),'San Carlos':(-36.4246,-71.9580),'Coihueco':(-36.6271,-71.8308),'Pinto':(-36.7042,-71.8920),'El Carmen':(-36.8995,-72.0252),'Pemuco':(-36.9770,-72.0988),'San Ignacio':(-36.7997,-72.0303),'Ránquil':(-36.6508,-72.5490),'Trehuaco':(-36.4309,-72.6682),'Cobquecura':(-36.1332,-72.7917),'Portezuelo':(-36.5355,-72.4280),'Ninhue':(-36.3938,-72.3970),'San Fabián':(-36.5561,-71.5499),'San Nicolás':(-36.4990,-72.2133),'Ñiquén':(-36.2860,-71.9010)}
COMMUNES={}
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
 book=pd.ExcelFile(XLSX,engine='openpyxl');data_sheet=links_sheet=None
 required={'Fecha Emision','Rut Destinatario','Nombre Destinatario','Dirección de Destino'}
 for sh in book.sheet_names:
  probe=pd.read_excel(XLSX,sheet_name=sh,engine='openpyxl',nrows=5)
  if required.issubset(set(probe.columns)):data_sheet=sh
  raw2=pd.read_excel(XLSX,sheet_name=sh,header=None,engine='openpyxl',usecols='A:B',nrows=800)
  if raw2.apply(lambda c:c.astype(str).str.startswith(('http://','https://')).any()).any():links_sheet=sh
 if not data_sheet or not links_sheet:raise SystemExit(f'ERROR: hojas no encontradas. Disponibles: {book.sheet_names}')
 print(f'OK: hoja AOL={data_sheet}; hoja Maps={links_sheet}');return data_sheet,links_sheet

def parse_pivot(raw,links_sheet):
 """Construye coincidencias por razón social + comuna y por dirección.
 Evita aplicar un enlace de una sucursal a otra comuna o dirección."""
 piv=pd.read_excel(XLSX,sheet_name=links_sheet,header=None,engine='openpyxl',usecols='A:B')
 names={norm(x) for x in raw['Nombre Destinatario'].dropna()};communes={norm(x) for x in raw['Ciudad de Destino'].dropna()};addresses={norm(x) for x in raw['Dirección de Destino'].dropna()}
 by_name_commune={};by_addr={};pending=[]
 for _,r in piv.iterrows():
  label=norm(r.iloc[0]);url=clean(r.iloc[1])
  if not label:continue
  if url.startswith('http'):
   # La fila URL normalmente cierra el bloque de comuna/direcciones de esa razón social.
   name=label if label in names else max(names,key=lambda n:SequenceMatcher(None,label,n).ratio(),default='')
   if name and SequenceMatcher(None,label,name).ratio()>=.82:
    block_communes=[x for x in pending if x in communes]
    block_addresses=[x for x in pending if x in addresses or any(c.isdigit() for c in x)]
    for c in block_communes:by_name_commune[(name,c)]=url
    # Un enlace colocado directamente en una dirección tiene prioridad absoluta.
    if label in addresses or any(c.isdigit() for c in label):by_addr[label]=url
    elif block_addresses and len(block_communes)==1:
     for a in block_addresses:by_addr.setdefault(a,url)
   pending=[]
  else:pending.append(label)
 return by_name_commune,by_addr

def extract(text):
 text=urllib.parse.unquote(str(text or '')).replace('\\u003d','=').replace('\\u0026','&').replace('\\/','/')
 for pat in [r'@(-?\d{1,2}\.\d+),(-?\d{1,3}\.\d+)',r'!3d(-?\d{1,2}\.\d+)!4d(-?\d{1,3}\.\d+)',r'[?&](?:q|query|center)=(-?\d{1,2}\.\d+)[,%]2?C?\s*(-?\d{1,3}\.\d+)',r'"latitude"\s*:\s*(-?\d{1,2}\.\d+).*?"longitude"\s*:\s*(-?\d{1,3}\.\d+)']:
  for m in re.finditer(pat,text,re.I|re.S):
   lat,lng=map(float,m.groups())
   if -38.8<lat<-35 and -73.5<lng<-70.5:return lat,lng
 return None
CACHE={}
def resolve(url):
 if not url:return None
 if url in CACHE:return CACHE[url]
 try:
  final=subprocess.check_output(['curl','-A','Mozilla/5.0','-LsS','--retry','2','--max-time','30','-o','/dev/null','-w','%{url_effective}',url],text=True).strip();c=extract(final)
  if not c:c=extract(subprocess.check_output(['curl','-A','Mozilla/5.0','-LsS','--retry','2','--max-time','30',url],text=True,errors='ignore'))
  CACHE[url]=c;return c
 except Exception:CACHE[url]=None;return None

def main():
 if not XLSX.exists():raise SystemExit('ERROR: falta datos/aol_actual.xlsx')
 data_sheet,links_sheet=find_sheets();raw=pd.read_excel(XLSX,sheet_name=data_sheet,engine='openpyxl')
 if 'Rut_Remitente_Visa' not in raw.columns:raise SystemExit('ERROR: falta Rut_Remitente_Visa')
 senders=set(raw['Rut_Remitente_Visa'].map(norm_rut));senders.discard('')
 df=raw[raw['Region Destino'].map(clean).str.upper().str.contains(r'XVI|NUBLE|ÑUBLE',regex=True,na=False)].copy()
 if 'Estado_Visacion' in df:df=df[~df['Estado_Visacion'].map(clean).str.upper().str.contains('ANULADA',na=False)].copy()
 df['_fecha']=pd.to_datetime(df['Fecha Emision'],errors='coerce');df=df[df['_fecha'].notna()];maxdate=df['_fecha'].max();df=df[df['_fecha']>=maxdate-pd.Timedelta(days=15)].copy()
 before=len(df);df=df[~df['Rut Destinatario'].map(norm_rut).isin(senders)].copy();excluded=before-len(df)
 # Corrección solicitada: no publicar el registro erróneo de Fernando Maureira en Chillán.
 bad=(df['Nombre Destinatario'].map(norm).eq('FERNANDO MAUREIRA') & df['Ciudad de Destino'].map(norm).eq('CHILLAN'))
 removed_fernando=int(bad.sum());df=df[~bad].copy()
 bync,byaddr=parse_pivot(raw,links_sheet);html=INDEX.read_text(encoding='utf-8');old,start,end=data_block(html)
 if start is None:raise SystemExit('ERROR: no se encontró const DATA=')
 saved={p.get('id'):(p.get('lat'),p.get('lng'),p.get('precision',''),p.get('mapsUrl','')) for p in old}
 places={}
 for _,r in df.iterrows():
  rk=norm_rut(r['Rut Destinatario']);com=clean(r['Ciudad de Destino']);nk=norm(r['Nombre Destinatario']);ak=norm(r['Dirección de Destino'])
  if not rk:continue
  # Un RUT conserva puntos separados solo cuando corresponde a otra comuna.
  key=rk+'|'+norm(com)
  p=places.setdefault(key,{'id':key,'rut':clean(r['Rut Destinatario']),'nombre':clean(r['Nombre Destinatario']),'comuna':com,'direccion':show_addr(r['Dirección de Destino']),'items':[],'_a':{},'_n':nk})
  if ak:p['_a'][ak]=p['_a'].get(ak,0)+1
  q=pd.to_numeric(r['Toneladas Visadas'],errors='coerce')
  p['items'].append({'emision':iso(r['Fecha Emision']),'tipo':clean(r['Producto/Recurso']),'especie':clean(r['Especie Recurso o Producto']),'descripcion':clean(r.get('Nombre Recurso o Producto','')),'ton':0 if pd.isna(q) else float(q),'objeto':clean(r.get('Objeto del Destino','')),'estado':clean(r.get('Estado_Visacion',''))})
 result=[];unresolved=[];verified=centers=market=0
 for p in places.values():
  best=max(p['_a'],key=p['_a'].get) if p['_a'] else ''
  if best:p['direccion']=show_addr(best)
  url=byaddr.get(best) or bync.get((p['_n'],norm(p['comuna'])),'');coord=resolve(url) if url else None
  if coord:
   p['lat'],p['lng']=coord;p['precision']='Google Maps verificado';p['mapsUrl']=url;verified+=1
  elif url and any(x in norm(p['direccion']) for x in ('MERCADO','PLAZA SARGENTO ALDEA','PLAZA 5 DE ABRIL','PLAZA CINCO DE ABRIL')):
   n=int(hashlib.sha256(p['id'].encode()).hexdigest()[:10],16);p['lat']=-36.60755+((n%7)-3)*.00016;p['lng']=-72.10355+(((n//7)%7)-3)*.00016;p['precision']='Cuadrante Mercado de Chillán';p['mapsUrl']=url;market+=1
  else:
   # No se pierde el punto: si Maps no responde, se conserva enlace y se usa centro comunal.
   lat,lng=CENTERS.get(p['comuna'],(-36.6066,-72.1034));dy,dx=offset(p['id']);p['lat'],p['lng']=lat+dy,lng+dx;p['precision']='Centro comunal aproximado';p['mapsUrl']=url;centers+=1
   if url:unresolved.append((p['nombre'],p['comuna'],p['direccion'],url))
  p.pop('_a',None);p.pop('_n',None);result.append(p)
 html=html[:start]+json.dumps(result,ensure_ascii=False,separators=(',',':'))+html[end:]
 stamp=maxdate.strftime('%Y-%m-%dT%H:%M:%S');html,n=re.subn(r"const\s+MAXDATE\s*=\s*new Date\('[^']*'\)",f"const MAXDATE=new Date('{stamp}')",html,count=1)
 html=re.sub(r"const\s+SOURCE_FILE\s*=\s*'[^']*'","const SOURCE_FILE='aol_actual.xlsx'",html,count=1);now=datetime.now(ZoneInfo('America/Santiago')).strftime('%d-%m-%Y %H:%M');html=re.sub(r"const\s+SOURCE_UPDATED\s*=\s*'[^']*'",f"const SOURCE_UPDATED='{now}'",html,count=1)
 INDEX.write_text(html,encoding='utf-8')
 communes=sorted({p['comuna'] for p in result});print(f'OK: {excluded} filas remitentes excluidas; Fernando Chillán eliminado={removed_fernando}')
 print(f'OK: {len(result)} locales; {verified} Maps exactos; {market} Mercado; {centers} centro comunal')
 print('OK: comunas publicadas:',', '.join(communes))
 if unresolved:
  print(f'AVISO: {len(unresolved)} enlaces no resolvieron coordenadas; se conservaron el enlace y el punto comunal:')
  for x in unresolved:print(' | '.join(x))
if __name__=='__main__':main()
