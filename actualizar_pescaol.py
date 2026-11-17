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

def update_detail_view(html):
    """Muestra todo lo emitido al destinatario durante los ultimos 15 dias."""
    new_show = r"""function show(x){selectedRut=x?.destRut||'';if(!x){$('detail').innerHTML='<div class="empty"><b>Selecciona un local en el mapa</b><br><span>Verás los recursos y productos recibidos durante los últimos 15 días.</span></div>';return}const gm='https://www.google.com/maps/search/?api=1&query='+encodeURIComponent([x.direccion,x.comuna,'Región de Ñuble, Chile'].filter(Boolean).join(', '));const cutoff=new Date(MAXDATE);cutoff.setDate(cutoff.getDate()-15);const items=DATA.filter(r=>r.destRut===x.destRut&&new Date(r.fecha)>=cutoff&&new Date(r.fecha)<=MAXDATE);const grouped=new Map();items.forEach(r=>{const type=(r.producto||'No informado').toUpperCase(),species=r.especie||'No informado',key=type+'|'+species;const prev=grouped.get(key)||{type,species,qty:0,latest:''};prev.qty+=Number(r.cantidad)||0;if(!prev.latest||new Date(r.fecha)>new Date(prev.latest))prev.latest=r.fecha;grouped.set(key,prev)});const rows=[...grouped.values()].sort((a,b)=>new Date(b.latest)-new Date(a.latest)||a.type.localeCompare(b.type,'es')||a.species.localeCompare(b.species,'es'));const rowHtml=rows.map(r=>`<tr><td data-label="Tipo"><span class="typePill ${r.type==='RECURSO'?'resource':'product'}">${esc(r.type)}</span></td><td data-label="Recurso o producto"><strong>${esc(r.species)}</strong></td><td data-label="Cantidad" class="num">${fn(r.qty)} t</td><td data-label="Última emisión">${fd(r.latest)}</td></tr>`).join('');const latest=items.reduce((v,r)=>!v||new Date(r.fecha)>new Date(v)?r.fecha:v,'');const total=rows.reduce((n,r)=>n+r.qty,0);$('detail').innerHTML=`<button id="backToMap" class="backToMap" type="button">← Volver al mapa</button><div class="placeHead"><div><span class="eyebrow">DESTINO TERMINAL</span><h2>${esc(x.nombre)}</h2><p class="addr">${esc(x.direccion)} · ${esc(x.comuna)}</p></div><a class="mapLink" href="${gm}" target="_blank" rel="noopener">Abrir mapa ↗</a></div><div class="summaryStrip"><div><small>Período</small><b>15 días</b></div><div><small>Tipos registrados</small><b>${rows.length}</b></div><div><small>Total</small><b>${fn(total)} t</b></div><div><small>Última emisión</small><b>${latest?fd(latest):'—'}</b></div></div><div class="tableTitle"><b>Recursos y productos</b><span>${fd(cutoff)} al ${fd(MAXDATE)}</span></div><div class="dataTableWrap"><table class="dataTable"><thead><tr><th>Tipo</th><th>Recurso o producto</th><th>Cantidad</th><th>Última emisión</th></tr></thead><tbody>${rowHtml||'<tr><td colspan="4">Sin información en el período.</td></tr>'}</tbody></table></div><p class="miniNote">Las cantidades repetidas del mismo recurso o producto se muestran sumadas.</p>`;const back=$('backToMap');if(back)back.onclick=scrollToMap;if(isMobileLayout())scrollToDetail();}"""
    html2, count = re.subn(r"function show\(x\)\{.*?\}\s*function render\(\)\{", new_show + "\nfunction render(){", html, count=1, flags=re.S)
    if count != 1:
        raise SystemExit('ERROR: no se pudo actualizar la ficha del local')
    table_css = r"""<style id="pescaol-table-ui">.placeHead{display:flex;justify-content:space-between;gap:14px;align-items:flex-start;margin:6px 0 14px}.placeHead h2{margin:4px 0}.eyebrow{font-size:.72rem;font-weight:800;letter-spacing:.08em;color:#0b6b78}.mapLink{white-space:nowrap;padding:8px 11px;border-radius:9px;background:#e8f4f6;color:#075a66;font-weight:750;text-decoration:none}.summaryStrip{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:12px 0 18px}.summaryStrip>div{background:#f2f7f8;border:1px solid #d9e7e9;border-radius:10px;padding:10px}.summaryStrip small{display:block;color:#617278;font-size:.72rem;margin-bottom:3px}.summaryStrip b{font-size:.94rem;color:#173a41}.tableTitle{display:flex;justify-content:space-between;gap:10px;align-items:end;margin:8px 0}.tableTitle span{font-size:.78rem;color:#687980}.dataTableWrap{overflow-x:auto;border:1px solid #d8e3e5;border-radius:12px}.dataTable{width:100%;border-collapse:collapse;font-size:.86rem;background:#fff}.dataTable th{text-align:left;background:#eef5f6;color:#3d555b;font-size:.72rem;text-transform:uppercase;letter-spacing:.04em;padding:9px 10px}.dataTable td{padding:10px;border-top:1px solid #e7edef;vertical-align:middle}.dataTable .num{text-align:right;font-weight:800;white-space:nowrap}.typePill{display:inline-block;padding:3px 7px;border-radius:999px;font-size:.68rem;font-weight:850;letter-spacing:.03em}.typePill.resource{background:#e4f4e9;color:#176438}.typePill.product{background:#fff0d7;color:#7a4c00}.miniNote{font-size:.75rem;color:#6b7a80;margin:8px 2px 0}@media(max-width:650px){.placeHead{display:block}.mapLink{display:inline-block;margin-top:8px}.summaryStrip{grid-template-columns:repeat(2,1fr)}.dataTable thead{display:none}.dataTable,.dataTable tbody,.dataTable tr,.dataTable td{display:block;width:100%;box-sizing:border-box}.dataTable tr{padding:8px 10px;border-top:1px solid #dfe8ea}.dataTable tr:first-child{border-top:0}.dataTable td{display:flex;justify-content:space-between;gap:12px;border:0;padding:4px 0;text-align:right}.dataTable td:before{content:attr(data-label);color:#6a797e;font-size:.72rem;text-align:left}.dataTable .num{text-align:right}.tableTitle{display:block}.tableTitle span{display:block;margin-top:2px}}</style>"""
    if 'id="pescaol-table-ui"' not in html2:
        html2 = html2.replace('</head>', table_css + '</head>', 1)
    html2 = html2.replace('todos los recursos o productos incluidos en la acreditación más reciente', 'todos los recursos y productos acreditados durante los últimos 15 días')
    html2 = html2.replace('Se muestra cada local una sola vez con todos los recursos y productos incluidos en su acreditación más reciente.', 'Se muestra cada local una sola vez con todos los recursos y productos acreditados durante los últimos 15 días.')
    html2 = re.sub(r'<div class="last"><small>Documento AOL</small><div><b>.*?</b></div>', '<div class="last">', html2, flags=re.S)
    return html2

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

    html2 = update_detail_view(html2)
    INDEX.write_text(html2, encoding='utf-8')
    print(f'OK: {len(rows)} filas publicadas desde {xlsx.name}')
    print(f'OK: fecha de corte {maxdate}')

if __name__ == '__main__':
    main()
