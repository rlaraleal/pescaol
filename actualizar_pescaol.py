#!/usr/bin/env python3
import json,re,math,time,urllib.request,urllib.parse,unicodedata
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from difflib import SequenceMatcher
import pandas as pd
ROOT=Path(__file__).resolve().parent; XLSX=ROOT/'datos'/'aol_actual.xlsx'; INDEX=ROOT/'index.html'; AUDIT=ROOT/'datos'/'auditoria_ubicaciones.csv'; CACHE=ROOT/'datos'/'geocache.json'
CENTERS={'Chillán':(-36.6066,-72.1034),'Chillán Viejo':(-36.6227,-72.1310),'Bulnes':(-36.7422,-72.2986),'Yungay':(-37.1211,-72.0137),'Quillón':(-36.7388,-72.4722),'Coelemu':(-36.4874,-72.7020),'Quirihue':(-36.2804,-72.5412),'San Carlos':(-36.4246,-71.9580),'Coihueco':(-36.6271,-71.8308),'Pinto':(-36.7042,-71.8920),'El Carmen':(-36.8995,-72.0252),'Pemuco':(-36.9770,-72.0988),'San Ignacio':(-36.7997,-72.0303),'Ránquil':(-36.6508,-72.5490),'Trehuaco':(-36.4309,-72.6682),'Cobquecura':(-36.1332,-72.7917),'Portezuelo':(-36.5355,-72.4280),'Ninhue':(-36.3938,-72.3970),'San Fabián':(-36.5561,-71.5499),'San Nicolás':(-36.4990,-72.2133),'Ñiquén':(-36.2860,-71.9010)}
OVERRIDES={('884904005','CHILLAN','COCHARCAS 815'):(-36.6130975,-72.1014415,'Coordenada maestra verificada')}
def clean(v):return '' if pd.isna(v) else re.sub(r'\s+',' ',str(v)).strip()
def norm_rut(v):return re.sub(r'[^0-9K]','',clean(v).upper())
def norm(v):
 s=''.join(c for c in unicodedata.normalize('NFD',clean(v).upper()) if unicodedata.category(c)!='Mn');s=re.sub(r'\bAVDA?\.?\b','AVENIDA',s);s=re.sub(r'\bPJE\.?\b','PASAJE',s);s=re.sub(r'\bCAM\.?\b','CAMINO',s);s=re.sub(r'\bN[°ºO]?\.?\s*','N ',s);return re.sub(r'[^A-Z0-9]+',' ',s).strip()
def display(v):return norm(v).title().replace(' N ',' N° ')
def iso(v):
 d=pd.to_datetime(v,errors='coerce');return '' if pd.isna(d) else d.strftime('%Y-%m-%dT%H:%M:%S')
def hav(a,b):
 R=6371000;p1,p2=map(math.radians,(a[0],b[0]));dp=math.radians(b[0]-a[0]);dl=math.radians(b[1]-a[1]);x=math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2;return 2*R*math.asin(math.sqrt(x))
def is_market_quadrant(name,address):
 t=norm(str(name)+' '+str(address))
 return any(x in t for x in ('MERCADO DE CHILLAN','MERCADO MUNICIPAL','PLAZA SARGENTO ALDEA','PLAZA 5 DE ABRIL','PLAZA CINCO DE ABRIL'))
def market_coord(key):
 # Stable grid strictly within the Mercado de Chillán quadrant, not Plaza de Armas.
 n=int(__import__('hashlib').sha256(key.encode()).hexdigest()[:10],16)
 row=(n%7)-3; col=((n//7)%7)-3
 return (-36.60755+row*0.00016,-72.10355+col*0.00016)
def data_block(html):
 m=re.search(r'\bconst\s+DATA\s*=\s*',html)
 if not m:return [],None,None
 obj,n=json.JSONDecoder().raw_decode(html[m.end():]);return obj,m.end(),m.end()+n
def sheets():
 book=pd.ExcelFile(XLSX,engine='openpyxl');data=links=None;req={'Fecha Emision','Rut Destinatario','Nombre Destinatario','Dirección de Destino'}
 for sh in book.sheet_names:
  probe=pd.read_excel(XLSX,sheet_name=sh,engine='openpyxl',nrows=5)
  if req.issubset(probe.columns):data=sh
  raw=pd.read_excel(XLSX,sheet_name=sh,header=None,engine='openpyxl',usecols='A:B',nrows=800)
  if raw.apply(lambda c:c.astype(str).str.startswith(('http://','https://')).any()).any():links=sh
 if not data or not links:raise SystemExit(f'ERROR: no se detectaron ambas hojas. Hojas: {book.sheet_names}')
 return data,links
def map_links(raw,sh):
 piv=pd.read_excel(XLSX,sheet_name=sh,header=None,engine='openpyxl',usecols='A:B');names={norm(x) for x in raw['Nombre Destinatario'].dropna()};addrs={norm(x) for x in raw['Dirección de Destino'].dropna()};bn={};ba={}
 for _,r in piv.iterrows():
  label=norm(r.iloc[0]);url=clean(r.iloc[1])
  if not label or not url.startswith('http'):continue
  if label in addrs or any(c.isdigit() for c in label):ba[label]=url;continue
  if label in names:bn[label]=url;continue
  best=max(names,key=lambda n:SequenceMatcher(None,label,n).ratio(),default='')
  if best and SequenceMatcher(None,label,best).ratio()>=.9:bn[best]=url
 return bn,ba
def extract_maps(url):
 if not url:return None
 try:
  req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0'})
  with urllib.request.urlopen(req,timeout=20) as r:txt=r.geturl()+' '+r.read(500000).decode('utf-8','ignore')
  txt=urllib.parse.unquote(txt)
  for pat in [r'@(-?\d{1,2}\.\d+),(-?\d{1,3}\.\d+)',r'!3d(-?\d{1,2}\.\d+)!4d(-?\d{1,3}\.\d+)',r'[?&](?:q|query|center)=(-?\d{1,2}\.\d+)[,%]2?C?\s*(-?\d{1,3}\.\d+)']:
   m=re.search(pat,txt)
   if m:
    c=tuple(map(float,m.groups()))
    if -38.8<c[0]<-35 and -73.5<c[1]<-70.5:return c
 except Exception:return None
 return None
def load_cache():
 try:return json.loads(CACHE.read_text(encoding='utf-8'))
 except:return {}
GEO=load_cache()
def geocode(address,commune):
 key=norm(address)+'|'+norm(commune)
 if key in GEO:return tuple(GEO[key]) if GEO[key] else None
 query=f'{address}, {commune}, Región de Ñuble, Chile';url='https://nominatim.openstreetmap.org/search?'+urllib.parse.urlencode({'q':query,'format':'jsonv2','limit':1,'countrycodes':'cl'})
 try:
  req=urllib.request.Request(url,headers={'User-Agent':'PescAOL/1.0 contacto: SERNAPESCA Nuble','Accept-Language':'es'})
  with urllib.request.urlopen(req,timeout=25) as r:data=json.loads(r.read().decode())
  c=(float(data[0]['lat']),float(data[0]['lon'])) if data else None
 except Exception:c=None
 GEO[key]=list(c) if c else None;time.sleep(1.05);return c
def main():
 if not XLSX.exists():raise SystemExit('ERROR: falta datos/aol_actual.xlsx')
 ds,ls=sheets();raw=pd.read_excel(XLSX,sheet_name=ds,engine='openpyxl')
 if 'Rut_Remitente_Visa' not in raw.columns:raise SystemExit('ERROR: falta Rut_Remitente_Visa')
 senders=set(raw['Rut_Remitente_Visa'].map(norm_rut));senders.discard('');df=raw[raw['Region Destino'].map(clean).str.upper().str.contains(r'XVI|NUBLE|ÑUBLE',regex=True,na=False)].copy()
 if 'Estado_Visacion' in df:df=df[~df['Estado_Visacion'].map(clean).str.upper().str.contains('ANULADA',na=False)].copy()
 df['_fecha']=pd.to_datetime(df['Fecha Emision'],errors='coerce');df=df[df['_fecha'].notna()];maxdate=df['_fecha'].max();df=df[df['_fecha']>=maxdate-pd.Timedelta(days=15)].copy();before=len(df);df=df[~df['Rut Destinatario'].map(norm_rut).isin(senders)].copy();excluded=before-len(df)
 # Error conocido: Fernando Maureira no se publica como Chillán.
 df=df[~(df['Nombre Destinatario'].map(norm).str.contains('FERNANDO MAUREIRA')&df['Ciudad de Destino'].map(norm).eq('CHILLAN'))].copy()
 bn,ba=map_links(raw,ls);html=INDEX.read_text(encoding='utf-8');_,start,end=data_block(html)
 if start is None:raise SystemExit('ERROR: falta const DATA=')
 # Cada establecimiento es RUT + comuna + dirección normalizada. No se fusionan sucursales.
 places={}
 for _,r in df.iterrows():
  rk=norm_rut(r['Rut Destinatario']);com=clean(r['Ciudad de Destino']);addr=norm(r['Dirección de Destino']);name=clean(r['Nombre Destinatario']);nk=norm(name)
  if not rk:continue
  key='|'.join((rk,norm(com),addr));p=places.setdefault(key,{'id':key,'rut':clean(r['Rut Destinatario']),'nombre':name,'comuna':com,'direccion':display(addr),'items':[],'_nk':nk,'_addr':addr})
  q=pd.to_numeric(r['Toneladas Visadas'],errors='coerce');p['items'].append({'emision':iso(r['Fecha Emision']),'tipo':clean(r['Producto/Recurso']),'especie':clean(r['Especie Recurso o Producto']),'descripcion':clean(r.get('Nombre Recurso o Producto','')),'ton':0 if pd.isna(q) else float(q),'objeto':clean(r.get('Objeto del Destino','')),'estado':clean(r.get('Estado_Visacion',''))})
 # Razón social solo se usa cuando tiene una única dirección elegible; evita mezclar sucursales.
 byname_count={}
 for p in places.values():byname_count[p['_nk']]=byname_count.get(p['_nk'],0)+1
 result=[];audit=[]
 for p in places.values():
  url=ba.get(p['_addr']) or (bn.get(p['_nk'],'') if byname_count[p['_nk']]==1 else '')
  override=OVERRIDES.get((norm_rut(p['rut']),norm(p['comuna']),p['_addr']));osm=geocode(p['direccion'],p['comuna']);gm=extract_maps(url);chosen=None;status=''
  if is_market_quadrant(p['nombre'],p['direccion']):
   chosen=market_coord(p['id']);status='Cuadrante Mercado de Chillán verificado'
  elif override:chosen=override[:2];status=override[2]
  elif gm and osm and hav(gm,osm)<=250:chosen=gm;status='Google Maps validado contra dirección'
  elif osm:chosen=osm;status='Dirección geocodificada y validada'
  elif gm:chosen=gm;status='Google Maps sin validación de dirección'
  else:chosen=CENTERS.get(p['comuna'],(-36.6066,-72.1034));status='Centro comunal aproximado'
  discrepancy=round(hav(gm,osm)) if gm and osm else ''
  p['lat'],p['lng']=chosen;p['precision']=status;p['mapsUrl']=url
  audit.append({'RUT':p['rut'],'Razón social':p['nombre'],'Comuna':p['comuna'],'Dirección':p['direccion'],'Latitud':chosen[0],'Longitud':chosen[1],'Estado':status,'Distancia Maps vs dirección (m)':discrepancy,'Enlace Maps':url})
  p.pop('_nk');p.pop('_addr');result.append(p)
 pd.DataFrame(audit).to_csv(AUDIT,index=False,encoding='utf-8-sig');CACHE.write_text(json.dumps(GEO,ensure_ascii=False),encoding='utf-8')
 html=html[:start]+json.dumps(result,ensure_ascii=False,separators=(',',':'))+html[end:];stamp=maxdate.strftime('%Y-%m-%dT%H:%M:%S');html,n=re.subn(r"const\s+MAXDATE\s*=\s*new Date\('[^']*'\)",f"const MAXDATE=new Date('{stamp}')",html,count=1);now=datetime.now(ZoneInfo('America/Santiago')).strftime('%d-%m-%Y %H:%M');html=re.sub(r"const\s+SOURCE_FILE\s*=\s*'[^']*'","const SOURCE_FILE='aol_actual.xlsx'",html,count=1);html=re.sub(r"const\s+SOURCE_UPDATED\s*=\s*'[^']*'",f"const SOURCE_UPDATED='{now}'",html,count=1);INDEX.write_text(html,encoding='utf-8')
 print(f'OK: {excluded} filas de remitentes excluidas; {len(result)} establecimientos finales')
 print(pd.Series([a['Estado'] for a in audit]).value_counts().to_string());print(f'OK: auditoría guardada en {AUDIT.relative_to(ROOT)}')
if __name__=='__main__':main()
