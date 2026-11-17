#!/usr/bin/env python3
import hashlib, json, re
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent
DATOS = ROOT / 'datos'
INDEX = ROOT / 'index.html'
REQUIRED = {
    'Rut_Remitente_Visa','Nr_Solicitud','Fecha Emision','N° Visación',
    'Producto/Recurso','Especie Recurso o Producto','Toneladas Visadas',
    'Rut Destinatario','Nombre Destinatario','Region Destino',
    'Ciudad de Destino','Dirección de Destino'
}

def clean(v):
    return '' if pd.isna(v) else re.sub(r'\s+', ' ', str(v)).strip()

def public_rut(v):
    r = re.sub(r'[^0-9Kk]', '', clean(v)).upper()
    if not r:
        return ''
    return 'pub-' + hashlib.sha256(('PescAOL|' + r).encode('utf-8')).hexdigest()[:16]

def accreditation_id(solicitud, visacion):
    # Nr_Solicitud identifica la AOL completa. Una misma solicitud puede contener
    # varias lineas y numeros de visacion, uno por recurso o producto.
    raw = clean(solicitud) or clean(visacion)
    return '' if not raw else hashlib.sha256(('PescAOL-AOL|' + raw).encode('utf-8')).hexdigest()[:16]

def iso_date(v):
    d = pd.to_datetime(v, errors='coerce')
    return '' if pd.isna(d) else d.strftime('%Y-%m-%dT%H:%M:%S')

def main():
    files = sorted(DATOS.glob('*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise SystemExit('ERROR: no hay archivos .xlsx en la carpeta datos/')
    xlsx = files[0]
    df = pd.read_excel(xlsx, engine='openpyxl')
    missing = sorted(REQUIRED - set(df.columns))
    if missing:
        raise SystemExit('ERROR: faltan columnas obligatorias: ' + ', '.join(missing))
    if df.empty:
        raise SystemExit('ERROR: el Excel no contiene filas')

    region = df['Region Destino'].map(clean).str.upper()
    filtered = df[region.str.contains(r'XVI|NUBLE|ÑUBLE', regex=True, na=False)].copy()
    if filtered.empty:
        raise SystemExit('ERROR: no se encontraron destinos de la Region de Nuble')

    filtered['_fecha'] = pd.to_datetime(filtered['Fecha Emision'], errors='coerce')
    filtered = filtered[filtered['_fecha'].notna()].copy()
    if filtered.empty:
        raise SystemExit('ERROR: ninguna Fecha Emision es valida')
    filtered.sort_values('_fecha', ascending=False, inplace=True)

    rows = []
    for _, r in filtered.iterrows():
        qty = pd.to_numeric(r['Toneladas Visadas'], errors='coerce')
        rows.append({
            'fecha': iso_date(r['Fecha Emision']),
            'remitenteRut': public_rut(r['Rut_Remitente_Visa']),
            'destRut': public_rut(r['Rut Destinatario']),
            'nombre': clean(r['Nombre Destinatario']),
            'comuna': clean(r['Ciudad de Destino']),
            'direccion': clean(r['Dirección de Destino']),
            'especie': clean(r['Especie Recurso o Producto']),
            'producto': clean(r['Producto/Recurso']),
            'cantidad': 0 if pd.isna(qty) else float(qty),
            # Clave irreversible de la solicitud completa. No publica el codigo AOL.
            'acreditacionId': accreditation_id(r['Nr_Solicitud'], r['N° Visación']),
        })

    maxdate = filtered['_fecha'].max().strftime('%Y-%m-%dT%H:%M:%S')
    html = INDEX.read_text(encoding='utf-8')
    data_json = json.dumps(rows, ensure_ascii=False, separators=(',', ':'))
    pattern = r"const DATA=\[.*?\],COORDS="
    replacement = 'const DATA=' + data_json + ',COORDS='
    html2, count = re.subn(pattern, replacement, html, count=1, flags=re.S)
    if count != 1:
        raise SystemExit('ERROR: no se encontro el bloque const DATA=[...] en index.html')
    html2, count = re.subn(
        r"MAXDATE=new Date\('[^']+'\)",
        "MAXDATE=new Date('" + maxdate + "')",
        html2, count=1
    )
    if count != 1:
        raise SystemExit('ERROR: no se encontro MAXDATE en index.html')

    # Defensa adicional: no permitir que reaparezca el bloque antiguo en la ficha.
    html2 = re.sub(
        r'<div class="last"><small>Documento AOL</small><div><b>.*?</b></div>',
        '<div class="last">',
        html2,
        flags=re.S
    )

    INDEX.write_text(html2, encoding='utf-8')
    print(f'OK: {len(rows)} filas publicadas desde {xlsx.name}')
    print(f'OK: fecha de corte {maxdate}')

if __name__ == '__main__':
    main()
