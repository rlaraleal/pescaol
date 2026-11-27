#!/usr/bin/env python3
import json,re,math,unicodedata
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd

ROOT=Path(__file__).resolve().parent
AOL=ROOT/'datos'/'aol_actual.xlsx'
COORD=ROOT/'datos'/'Coordenadas.xlsx'
INDEX=ROOT/'index.html'
AUDIT=ROOT/'datos'/'auditoria_ubicaciones.csv'
CENTERS={'Chillán':(-36.6066,-72.1034),'Chillán Viejo':(-36.6227,-72.1310),'Bulnes':(-36.7422,-72.2986),'Yungay':(-37.1211,-72.0137),'Quillón':(-36.7388,-72.4722),'Coelemu':(-36.4874,-72.7020),'Quirihue':(-36.2804,-72.5412),'San Carlos':(-36.4246,-71.9580),'Coihueco':(-36.6271,-71.8308),'Pinto':(-36.7042,-71.8920),'El Carmen':(-36.8995,-72.0252),'Pemuco':(-36.9770,-72.0988),'San Ignacio':(-36.7997,-72.0303),'Ránquil':(-36.6508,-72.5490),'Trehuaco':(-36.4309,-72.6682),'Cobquecura':(-36.1332,-72.7917),'Portezuelo':(-36.5355,-72.4280),'Ninhue':(-36.3938,-72.3970),'San Fabián':(-36.5561,-71.5499),'San Nicolás':(-36.4990,-72.2133),'Ñiquén':(-36.2860,-71.9010)}

def clean(v): return '' if pd.isna(v) else re.sub(r'\s+',' ',str(v)).strip()
def norm_rut(v): return re.sub(r'[^0-9K]','',clean(v).upper())
def norm(v):
 s=''.join(c for c in unicodedata.normalize('NFD',clean(v).upper()) if unicodedata.category(c)!='Mn')
 s=re.sub(r'\bAVDA?\.?\b','AVENIDA',s);s=re.sub(r'\bPJE\.?\b','PASAJE',s);s=re.sub(r'\bCAM\.?\b','CAMINO',s);s=re.sub(r'\bN[°ºO]?\.?\s*','N ',s)
 return re.sub(r'[^A-Z0-9]+',' ',s).strip()
def display(v): return norm(v).title().replace(' N ',' N° ')
def iso(v):
 d=pd.to_datetime(v,errors='coerce',dayfirst=True);return '' if pd.isna(d) else d.strftime('%Y-%m-%dT%H:%M:%S')
def dist(a,b):
 R=6371;p1,p2=map(math.radians,(a[0],b[0]));dp=math.radians(b[0]-a[0]);dl=math.radians(b[1]-a[1]);x=math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2;return 2*R*math.asin(math.sqrt(x))
def data_block(html):
 m=re.search(r'\bconst\s+DATA\s*=\s*',html)
 if not m:return [],None,None
 obj,n=json.JSONDecoder().raw_decode(html[m.end():]);return obj,m.end(),m.end()+n

def coordinate_catalog(raw):
 c=pd.read_excel(COORD,sheet_name=0,header=None,engine='openpyxl')
 names={norm(x) for x in raw['Nombre Destinatario'].dropna()};addresses={norm(x) for x in raw['Dirección de Destino'].dropna()}
 by_name={};by_addr={};invalid=[]
 for idx,r in c.iterrows():
  label=clean(r.iloc[0] if len(r)>0 else '');url=clean(r.iloc[1] if len(r)>1 else '');txt=clean(r.iloc[5] if len(r)>5 else '')
  m=re.search(r'(-?\d{1,2}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)',txt)
  if not m:continue
  point=(float(m.group(1)),float(m.group(2)))
  if not (-38.8<point[0]<-35 and -73.5<point[1]<-70.5):invalid.append((idx+1,label,txt));continue
  k=norm(label);rec={'point':point,'url':url,'label':label,'row':idx+1}
  if k in addresses or any(ch.isdigit() for ch in k):by_addr.setdefault(k,[]).append(rec)
  elif k in names:by_name.setdefault(k,[]).append(rec)
  elif k and k!='EN BLANCO':invalid.append((idx+1,label,'Etiqueta no reconocida'))
 return by_name,by_addr,invalid

def main():
 if not AOL.exists():raise SystemExit('ERROR: falta datos/aol_actual.xlsx')
 if not COORD.exists():raise SystemExit('ERROR: falta datos/Coordenadas.xlsx')
 raw=pd.read_excel(AOL,sheet_name=0,engine='openpyxl')
 required={'Fecha Emision','Rut_Remitente_Visa','Rut Destinatario','Nombre Destinatario','Region Destino','Ciudad de Destino','Dirección de Destino','Producto/Recurso','Especie Recurso o Producto','Toneladas Visadas'}
 missing=sorted(required-set(raw.columns))
 if missing:raise SystemExit('ERROR: faltan columnas AOL: '+', '.join(missing))
 all_dates=pd.to_datetime(raw['Fecha Emision'],errors='coerce',dayfirst=True);file_max=all_dates.max()
 senders=set(raw['Rut_Remitente_Visa'].map(norm_rut));senders.discard('')
 df=raw[raw['Region Destino'].map(clean).str.upper().str.contains(r'XVI|NUBLE|ÑUBLE',regex=True,na=False)].copy()
 if 'Estado_Visacion' in df.columns:df=df[~df['Estado_Visacion'].map(clean).str.upper().str.contains('ANULADA',na=False)].copy()
 df['_fecha']=pd.to_datetime(df['Fecha Emision'],errors='coerce',dayfirst=True);df=df[df['_fecha'].notna()].copy();eligible_max=df['_fecha'].max();df=df[df['_fecha']>=eligible_max-pd.Timedelta(days=15)].copy()
 before=len(df);df=df[~df['Rut Destinatario'].map(norm_rut).isin(senders)].copy();excluded=before-len(df)
 # Corrección solicitada: Fernando Maureira no se publica como Chillán.
 df=df[~(df['Nombre Destinatario'].map(norm).str.contains('FERNANDO MAUREIRA')&df['Ciudad de Destino'].map(norm).eq('CHILLAN'))].copy()
 by_name,by_addr,invalid=coordinate_catalog(raw)
 # Establecimientos AOL por RUT + comuna + dirección normalizada.
 base={}
 for _,r in df.iterrows():
  rk=norm_rut(r['Rut Destinatario']);com=clean(r['Ciudad de Destino']);addr=norm(r['Dirección de Destino']);nk=norm(r['Nombre Destinatario'])
  if not rk:continue
  key='|'.join((rk,norm(com),addr));p=base.setdefault(key,{'id':key,'rut':clean(r['Rut Destinatario']),'nombre':clean(r['Nombre Destinatario']),'comuna':com,'direccion':display(addr),'items':[],'_nk':nk,'_addr':addr})
  q=pd.to_numeric(r['Toneladas Visadas'],errors='coerce')
  p['items'].append({'emision':iso(r['Fecha Emision']),'tipo':clean(r['Producto/Recurso']),'especie':clean(r['Especie Recurso o Producto']),'descripcion':clean(r.get('Nombre Recurso o Producto','')),'ton':0 if pd.isna(q) else float(q),'objeto':clean(r.get('Objeto del Destino','')),'estado':clean(r.get('Estado_Visacion',''))})
 # Locales posibles por razón social para evitar asignar la misma coordenada a sucursales distintas.
 name_places={}
 for p in base.values():name_places.setdefault(p['_nk'],[]).append(p)
 assigned=[]
 for p in base.values():
  candidates=by_addr.get(p['_addr'],[]);method=''
  if candidates:
   rec=min(candidates,key=lambda x:dist(x['point'],CENTERS.get(p['comuna'],x['point'])));method='Coordenada por dirección'
  else:
   name_candidates=by_name.get(p['_nk'],[]);rec=None
   if name_candidates:
    # La coordenada de razón social se asigna solo a la sucursal/comuna geográficamente compatible.
    ranked=sorted(((dist(c['point'],CENTERS.get(p['comuna'],c['point'])),c) for c in name_candidates),key=lambda x:x[0])
    if ranked and ranked[0][0]<=45:
     rec=ranked[0][1];method='Coordenada por razón social y comuna'
  if rec:
   point=rec['point'];url=rec['url'];coord_key=(norm_rut(p['rut']),round(point[0],6),round(point[1],6))
  else:
   point=CENTERS.get(p['comuna'],(-36.6066,-72.1034));url='';method='Centro comunal aproximado';coord_key=(norm_rut(p['rut']),round(point[0],6),round(point[1],6))
  p['lat'],p['lng']=point;p['precision']=method;p['mapsUrl']=url;p['_coord_key']=coord_key;assigned.append(p)
 # Fusiona variantes de dirección del mismo RUT que terminaron en la misma coordenada exacta.
 merged={}
 for p in assigned:
  k=p['_coord_key']
  if k not in merged:merged[k]=p
  else:
   merged[k]['items'].extend(p['items'])
   if len(p['direccion'])>len(merged[k]['direccion']):merged[k]['direccion']=p['direccion']
 result=[];audit=[]
 for p in merged.values():
  p.pop('_nk',None);p.pop('_addr',None);p.pop('_coord_key',None);result.append(p)
  audit.append({'RUT':p['rut'],'Razón social':p['nombre'],'Comuna':p['comuna'],'Dirección':p['direccion'],'Latitud':p['lat'],'Longitud':p['lng'],'Estado':p['precision'],'Enlace Maps':p['mapsUrl']})
 html=INDEX.read_text(encoding='utf-8');_,start,end=data_block(html)
 if start is None:raise SystemExit('ERROR: no se encontró const DATA= en index.html')
 html=html[:start]+json.dumps(result,ensure_ascii=False,separators=(',',':'))+html[end:]
 stamp=eligible_max.strftime('%Y-%m-%dT%H:%M:%S');html,n=re.subn(r"const\s+MAXDATE\s*=\s*new Date\('[^']*'\)",f"const MAXDATE=new Date('{stamp}')",html,count=1)
 now=datetime.now(ZoneInfo('America/Santiago')).strftime('%d-%m-%Y %H:%M');html=re.sub(r"const\s+SOURCE_FILE\s*=\s*'[^']*'","const SOURCE_FILE='aol_actual.xlsx + Coordenadas.xlsx'",html,count=1);html=re.sub(r"const\s+SOURCE_UPDATED\s*=\s*'[^']*'",f"const SOURCE_UPDATED='{now}'",html,count=1)
 INDEX.write_text(html,encoding='utf-8');pd.DataFrame(audit).to_csv(AUDIT,index=False,encoding='utf-8-sig')
 exact=sum(p['precision']!='Centro comunal aproximado' for p in result);centers=len(result)-exact
 print(f'OK: última AOL del archivo {file_max}; última AOL publicable {eligible_max}')
 print(f'OK: {excluded} filas excluidas porque el destinatario también es remitente')
 print(f'OK: {len(result)} locales: {exact} con coordenadas maestras, {centers} en centro comunal')
 print(f'OK: {sum(len(v) for v in by_name.values())} coordenadas por razón social y {sum(len(v) for v in by_addr.values())} por dirección')
 if invalid:print(f'AVISO: {len(invalid)} filas de Coordenadas.xlsx no se usaron; revisar auditoría o estructura')
if __name__=='__main__':main()
