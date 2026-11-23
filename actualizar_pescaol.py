#!/usr/bin/env python3
import json, re, hashlib
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent
DATOS = ROOT / 'datos'
INDEX = ROOT / 'index.html'
REQUIRED = {
    'Fecha Emision','Producto/Recurso','Especie Recurso o Producto',
    'Toneladas Visadas','Rut Destinatario','Nombre Destinatario',
    'Region Destino','Ciudad de Destino','Dirección de Destino'
}
CENTERS = {
 'Chillán':(-36.6066,-72.1034),'Chillán Viejo':(-36.6227,-72.1310),
 'Bulnes':(-36.7422,-72.2986),'Yungay':(-37.1211,-72.0137),
 'Quillón':(-36.7388,-72.4722),'Coelemu':(-36.4874,-72.7020),
 'Quirihue':(-36.2804,-72.5412),'San Carlos':(-36.4246,-71.9580),
 'Coihueco':(-36.6271,-71.8308),'Pinto':(-36.7042,-71.8920),
 'El Carmen':(-36.8995,-72.0252),'Pemuco':(-36.9770,-72.0988),
 'San Ignacio':(-36.7997,-72.0303),'Ránquil':(-36.6508,-72.5490),
 'Trehuaco':(-36.4309,-72.6682),'Cobquecura':(-36.1332,-72.7917),
 'Portezuelo':(-36.5355,-72.4280),'Ninhue':(-36.3938,-72.3970),
 'San Fabián':(-36.5561,-71.5499),'San Nicolás':(-36.4990,-72.2133),
 'Ñiquén':(-36.2860,-71.9010)
}

def clean(v):
    return '' if pd.isna(v) else re.sub(r'\s+', ' ', str(v)).strip()

def iso_date(v):
    d = pd.to_datetime(v, errors='coerce')
    return '' if pd.isna(d) else d.strftime('%Y-%m-%dT%H:%M:%S')

def stable_offset(key):
    n=int(hashlib.sha256(key.encode('utf-8')).hexdigest()[:12],16)
    return ((n%101)-50)/100000, (((n//101)%101)-50)/100000

def extract_current_data(html):
    m=re.search(r'\bconst\s+DATA\s*=\s*',html)
    if not m: return [],None,None
    start=m.end()
    dec=json.JSONDecoder()
    try:
        value,length=dec.raw_decode(html[start:])
    except json.JSONDecodeError as e:
        raise SystemExit(f'ERROR: el bloque DATA existe, pero no contiene JSON válido: {e}')
    return value,start,start+length

def coord_lookup(current):
    out={}
    for p in current if isinstance(current,list) else []:
        if not isinstance(p,dict) or 'lat' not in p or 'lng' not in p: continue
        exact='|'.join(clean(p.get(k,'')) for k in ('rut','nombre','comuna','direccion'))
        out[exact]=(p['lat'],p['lng'],p.get('precision','Ubicación conservada'))
        rut=clean(p.get('rut',''))
        if rut and rut not in out:
            out[rut]=(p['lat'],p['lng'],p.get('precision','Ubicación conservada'))
    return out

def main():
    files=sorted(DATOS.glob('*.xlsx'),key=lambda p:p.stat().st_mtime,reverse=True)
    if not files: raise SystemExit('ERROR: no hay archivos .xlsx en datos/')
    xlsx=files[0]
    df=pd.read_excel(xlsx,engine='openpyxl')
    missing=sorted(REQUIRED-set(df.columns))
    if missing: raise SystemExit('ERROR: faltan columnas obligatorias: '+', '.join(missing))
    if df.empty: raise SystemExit('ERROR: el Excel no contiene filas')

    region=df['Region Destino'].map(clean).str.upper()
    df=df[region.str.contains(r'XVI|NUBLE|ÑUBLE',regex=True,na=False)].copy()
    if 'Estado_Visacion' in df.columns:
        estado=df['Estado_Visacion'].map(clean).str.upper()
        df=df[~estado.str.contains('ANULADA',na=False)].copy()
    df['_fecha']=pd.to_datetime(df['Fecha Emision'],errors='coerce')
    df=df[df['_fecha'].notna()].copy()
    if df.empty: raise SystemExit('ERROR: no hay filas válidas para la Región de Ñuble')
    maxdate=df['_fecha'].max()
    # El HTML puede ofrecer hasta 15 días; no se publican filas más antiguas.
    df=df[df['_fecha']>=maxdate-pd.Timedelta(days=15)].copy()
    df.sort_values('_fecha',ascending=False,inplace=True)

    html=INDEX.read_text(encoding='utf-8')
    current,start,end=extract_current_data(html)
    if start is None:
        raise SystemExit('ERROR: no se encontró "const DATA=" en index.html')
    oldcoords=coord_lookup(current)

    places={}
    for _,r in df.iterrows():
        rut=clean(r['Rut Destinatario']); nombre=clean(r['Nombre Destinatario'])
        comuna=clean(r['Ciudad de Destino']); direccion=clean(r['Dirección de Destino'])
        key='|'.join((rut,nombre,comuna,direccion))
        p=places.setdefault(key,{'id':key,'rut':rut,'nombre':nombre,'comuna':comuna,'direccion':direccion,'items':[]})
        qty=pd.to_numeric(r['Toneladas Visadas'],errors='coerce')
        p['items'].append({
            'emision':iso_date(r['Fecha Emision']),
            'tipo':clean(r['Producto/Recurso']),
            'especie':clean(r['Especie Recurso o Producto']),
            'descripcion':clean(r['Nombre Recurso o Producto']) if 'Nombre Recurso o Producto' in df.columns else clean(r['Especie Recurso o Producto']),
            'ton':0 if pd.isna(qty) else float(qty),
            'objeto':clean(r['Objeto del Destino']) if 'Objeto del Destino' in df.columns else '',
            'estado':clean(r['Estado_Visacion']) if 'Estado_Visacion' in df.columns else ''
        })

    result=[]
    for i,p in enumerate(places.values()):
        exact=p['id']; saved=oldcoords.get(exact) or oldcoords.get(p['rut'])
        text=(p['nombre']+' '+p['direccion']).lower()
        market=any(x in text for x in ('mercado','plaza sargento aldea','plaza sgto','plaza 5 de abril','plaza cinco de abril'))
        if saved:
            p['lat'],p['lng'],p['precision']=saved
        elif market:
            p['lat']=-36.60755+((i%7)-3)*0.00016
            p['lng']=-72.10355+(((i//7)%7)-3)*0.00016
            p['precision']='Cuadrante Mercado de Chillán'
        else:
            lat,lng=CENTERS.get(p['comuna'],(-36.6066,-72.1034))
            dy,dx=stable_offset(exact)
            p['lat'],p['lng']=lat+dy,lng+dx
            p['precision']='Centro comunal aproximado'
        result.append(p)

    data_json=json.dumps(result,ensure_ascii=False,separators=(',',':'))
    html=html[:start]+data_json+html[end:]
    stamp=maxdate.strftime('%Y-%m-%dT%H:%M:%S')
    html,n=re.subn(r"const\s+MAXDATE\s*=\s*new Date\('[^']*'\)",f"const MAXDATE=new Date('{stamp}')",html,count=1)
    if n!=1:
        raise SystemExit('ERROR: no se encontró const MAXDATE=new Date(...) en index.html')
    # Actualiza textos visibles de fecha sin alterar la lógica móvil desde la hora real.
    d=maxdate.strftime('%d-%m-%Y')
    html=re.sub(r'Datos actualizados hasta el \d{2}-\d{2}-\d{4}\.',f'Datos actualizados hasta el {d}.',html,count=1)
    html=re.sub(r'Datos actualizados hasta: \d{4}-\d{2}-\d{2}\.',f'Datos actualizados hasta: {maxdate:%Y-%m-%d}.',html,count=1)
    INDEX.write_text(html,encoding='utf-8')
    print(f'OK: {len(df)} filas y {len(result)} ubicaciones publicadas desde {xlsx.name}')
    print(f'OK: fecha de corte {stamp}')
    print('OK: se conservaron las coordenadas existentes cuando fue posible')

if __name__=='__main__':
    main()
