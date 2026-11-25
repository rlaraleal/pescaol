#!/usr/bin/env python3
import json, re, hashlib, urllib.request, urllib.parse
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
from difflib import SequenceMatcher
import pandas as pd

ROOT=Path(__file__).resolve().parent
DATOS=ROOT/'datos'; INDEX=ROOT/'index.html'; XLSX=DATOS/'aol_actual.xlsx'
CENTERS={'Chillán':(-36.6066,-72.1034),'Chillán Viejo':(-36.6227,-72.1310),'Bulnes':(-36.7422,-72.2986),'Yungay':(-37.1211,-72.0137),'Quillón':(-36.7388,-72.4722),'Coelemu':(-36.4874,-72.7020),'Quirihue':(-36.2804,-72.5412),'San Carlos':(-36.4246,-71.9580),'Coihueco':(-36.6271,-71.8308),'Pinto':(-36.7042,-71.8920),'El Carmen':(-36.8995,-72.0252),'Pemuco':(-36.9770,-72.0988),'San Ignacio':(-36.7997,-72.0303),'Ránquil':(-36.6508,-72.5490),'Trehuaco':(-36.4309,-72.6682),'Cobquecura':(-36.1332,-72.7917),'Portezuelo':(-36.5355,-72.4280),'Ninhue':(-36.3938,-72.3970),'San Fabián':(-36.5561,-71.5499),'San Nicolás':(-36.4990,-72.2133),'Ñiquén':(-36.2860,-71.9010)}
REQUIRED={'Fecha Emision','Producto/Recurso','Especie Recurso o Producto','Toneladas Visadas','Rut Destinatario','Nombre Destinatario','Region Destino','Ciudad de Destino','Dirección de Destino'}

def clean(v): return '' if pd.isna(v) else re.sub(r'\s+',' ',str(v)).strip()
def norm_rut(v): return re.sub(r'[^0-9K]','',clean(v).upper())
def norm_text(v):
    s=clean(v).upper(); s=''.join(c for c in __import__('unicodedata').normalize('NFD',s) if __import__('unicodedata').category(c)!='Mn')
    s=re.sub(r'\bAVDA?\.?\b','AVENIDA',s); s=re.sub(r'\bPJE\.?\b','PASAJE',s); s=re.sub(r'\bCAM\.?\b','CAMINO',s)
    s=re.sub(r'\bN[°ºO]?\.?\s*','N ',s); return re.sub(r'[^A-Z0-9]+',' ',s).strip()
def display_address(v): return norm_text(v).title().replace(' N ',' N° ')
def iso(v):
    d=pd.to_datetime(v,errors='coerce'); return '' if pd.isna(d) else d.strftime('%Y-%m-%dT%H:%M:%S')
def stable_offset(key):
    n=int(hashlib.sha256(key.encode()).hexdigest()[:12],16); return ((n%101)-50)/100000,(((n//101)%101)-50)/100000

def current_data(html):
    m=re.search(r'\bconst\s+DATA\s*=\s*',html)
    if not m:return [],None,None
    val,n=json.JSONDecoder().raw_decode(html[m.end():]); return val,m.end(),m.end()+n

def coord_lookup(cur):
    out={}
    for p in cur if isinstance(cur,list) else []:
        if 'lat' not in p or 'lng' not in p:continue
        key=norm_rut(p.get('rut'))+'|'+norm_text(p.get('comuna'))
        out[key]=(p['lat'],p['lng'],p.get('precision','Ubicación conservada'),p.get('mapsUrl',''))
    return out

def read_map_links(book,raw):
    try: piv=pd.read_excel(book,sheet_name='Hoja1',header=None,engine='openpyxl',usecols='A:B')
    except Exception:return {},{}
    names={norm_text(x) for x in raw['Nombre Destinatario'].dropna()}; addrs={norm_text(x) for x in raw['Dirección de Destino'].dropna()}
    by_name={}; by_addr={}
    for _,r in piv.iterrows():
        label=norm_text(r.iloc[0] if len(r)>0 else ''); url=clean(r.iloc[1] if len(r)>1 else '')
        if not label or not url.startswith('http'):continue
        if label in addrs or any(c.isdigit() for c in label): by_addr[label]=url; continue
        if label in names: by_name[label]=url; continue
        match=max(names,key=lambda n:SequenceMatcher(None,label,n).ratio(),default='')
        if match and SequenceMatcher(None,label,match).ratio()>=0.88: by_name[match]=url
    return by_name,by_addr

def resolve_maps(url):
    if not url:return None
    try:
        req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0'})
        with urllib.request.urlopen(req,timeout=15) as r: final=r.geturl(); body=r.read(250000).decode('utf-8','ignore')
        texts=[final,urllib.parse.unquote(final),body]
        patterns=[r'@(-?\d{1,2}\.\d+),(-?\d{1,3}\.\d+)',r'!3d(-?\d{1,2}\.\d+)!4d(-?\d{1,3}\.\d+)',r'center=(-?\d{1,2}\.\d+)%2C(-?\d{1,3}\.\d+)']
        for t in texts:
            for pat in patterns:
                m=re.search(pat,t)
                if m:
                    lat,lng=map(float,m.groups())
                    if -38.8<lat<-35 and -73.5<lng<-70.5:return lat,lng
    except Exception as e: print(f'AVISO: no se pudo resolver {url}: {e}')
    return None

def main():
    if not XLSX.exists():raise SystemExit('ERROR: falta datos/aol_actual.xlsx')
    raw=pd.read_excel(XLSX,sheet_name='Sheet1',engine='openpyxl')
    missing=sorted(REQUIRED-set(raw.columns))
    if missing:raise SystemExit('ERROR: faltan columnas: '+', '.join(missing))
    sender_col=next((c for c in ('Rut_Remitente_Visa','RUT_REMITENTE_VISA','Rut Remitente Visa','Rut Remitente','RUT Remitente') if c in raw.columns),None)
    if not sender_col:raise SystemExit('ERROR: no se encontró Rut_Remitente_Visa; se detiene para no publicar agentes intermedios')
    senders=set(raw[sender_col].map(norm_rut)); senders.discard('')
    region=raw['Region Destino'].map(clean).str.upper(); df=raw[region.str.contains(r'XVI|NUBLE|ÑUBLE',regex=True,na=False)].copy()
    if 'Estado_Visacion' in df.columns: df=df[~df['Estado_Visacion'].map(clean).str.upper().str.contains('ANULADA',na=False)].copy()
    df['_fecha']=pd.to_datetime(df['Fecha Emision'],errors='coerce'); df=df[df['_fecha'].notna()].copy(); maxdate=df['_fecha'].max(); df=df[df['_fecha']>=maxdate-pd.Timedelta(days=15)].copy()
    before=len(df); df=df[~df['Rut Destinatario'].map(norm_rut).isin(senders)].copy(); excluded=before-len(df)
    by_name,by_addr=read_map_links(XLSX,raw)
    html=INDEX.read_text(encoding='utf-8'); cur,start,end=current_data(html)
    if start is None:raise SystemExit('ERROR: no se encontró const DATA= en index.html')
    saved=coord_lookup(cur); places={}
    for _,r in df.iterrows():
        rk=norm_rut(r['Rut Destinatario']); comuna=clean(r['Ciudad de Destino']); nk=norm_text(r['Nombre Destinatario']); ak=norm_text(r['Dirección de Destino'])
        if not rk:continue
        key=rk+'|'+norm_text(comuna) # un RUT puede tener puntos distintos en comunas distintas
        p=places.setdefault(key,{'id':key,'rut':clean(r['Rut Destinatario']),'nombre':clean(r['Nombre Destinatario']),'comuna':comuna,'direccion':display_address(r['Dirección de Destino']),'items':[],'_addresses':{},'_namekey':nk})
        if ak:p['_addresses'][ak]=p['_addresses'].get(ak,0)+1
        qty=pd.to_numeric(r['Toneladas Visadas'],errors='coerce')
        p['items'].append({'emision':iso(r['Fecha Emision']),'tipo':clean(r['Producto/Recurso']),'especie':clean(r['Especie Recurso o Producto']),'descripcion':clean(r['Nombre Recurso o Producto']) if 'Nombre Recurso o Producto' in df.columns else clean(r['Especie Recurso o Producto']),'ton':0 if pd.isna(qty) else float(qty),'objeto':clean(r['Objeto del Destino']) if 'Objeto del Destino' in df.columns else '','estado':clean(r['Estado_Visacion']) if 'Estado_Visacion' in df.columns else ''})
    verified=0; centers=0; result=[]
    for p in places.values():
        best=max(p['_addresses'],key=p['_addresses'].get) if p['_addresses'] else ''
        if best:p['direccion']=display_address(best)
        mapurl=by_addr.get(best) or by_name.get(p['_namekey'],'')
        coord=resolve_maps(mapurl) if mapurl else None
        if coord:
            p['lat'],p['lng']=coord; p['precision']='Google Maps verificado'; p['mapsUrl']=mapurl; verified+=1
        elif p['id'] in saved and not re.search('aproximad|centro comunal',saved[p['id']][2],re.I):
            p['lat'],p['lng'],p['precision'],oldurl=saved[p['id']]; p['mapsUrl']=mapurl or oldurl; verified+=1
        else:
            lat,lng=CENTERS.get(p['comuna'],(-36.6066,-72.1034)); dy,dx=stable_offset(p['id']); p['lat'],p['lng']=lat+dy,lng+dx; p['precision']='Centro comunal aproximado'; p['mapsUrl']=mapurl; centers+=1
        p.pop('_addresses',None);p.pop('_namekey',None);result.append(p)
    html=html[:start]+json.dumps(result,ensure_ascii=False,separators=(',',':'))+html[end:]
    stamp=maxdate.strftime('%Y-%m-%dT%H:%M:%S'); html,n=re.subn(r"const\s+MAXDATE\s*=\s*new Date\('[^']*'\)",f"const MAXDATE=new Date('{stamp}')",html,count=1)
    if n!=1:raise SystemExit('ERROR: falta MAXDATE')
    processed=datetime.now(ZoneInfo('America/Santiago')).strftime('%d-%m-%Y %H:%M')
    html=re.sub(r"const\s+SOURCE_FILE\s*=\s*'[^']*'","const SOURCE_FILE='aol_actual.xlsx'",html,count=1)
    html=re.sub(r"const\s+SOURCE_UPDATED\s*=\s*'[^']*'",f"const SOURCE_UPDATED='{processed}'",html,count=1)
    INDEX.write_text(html,encoding='utf-8')
    print(f'OK: {excluded} filas excluidas porque el destinatario también es remitente ({sender_col})')
    print(f'OK: {len(result)} locales finales; {verified} con coordenadas verificadas y {centers} en centro comunal')
    print(f'OK: {len(by_name)} enlaces por razón social y {len(by_addr)} enlaces por dirección detectados en Hoja1')

if __name__=='__main__':main()
