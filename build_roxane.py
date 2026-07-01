#!/usr/bin/env python3
"""
Genera location/granollers/data.json a partir del Excel de Roxane (Granollers).

El centro de Granollers NO usa Dory sino Roxane, con un modelo de datos distinto al
de Séneca (no hay fórmulas/líneas, sino órdenes de producción, "pours", peticiones de
muestra por mes y pesaje manual vs robot). Solo stdlib: lee el .xlsx como zip+XML, sin
dependencias externas. La hoja de datos crudos "Exportación MANUAL" (~450MB, 625k filas)
se procesa por STREAMING (sin cargarla en memoria), no se descarta.

Uso:
    python build_roxane.py
Lee:  ./location/granollers/Data_Roxane_Gra.xlsx
        · hoja "Stats Samples"      → serie mensual (órdenes prod., pours, muestras, días)
        · hoja "Stats MP"           → top materias primas 2025 (producción y sample requests)
        · hoja "Exportación MANUAL" → pesadas individuales; de la col. "Módulo" se deriva
          manual (MWS_*, técnicos) vs robot (ROXY_*). Genera: manual/robot por mes,
          ranking de estaciones de pesaje y peso dosificado por vía.
Escribe: ./location/granollers/data.json
"""
import os, re, io, json, zipfile
from datetime import datetime, timedelta
from xml.etree.ElementTree import iterparse

BASE = os.path.dirname(__file__)
XLSX = os.path.join(BASE, "location", "granollers", "Data_Roxane_Gra.xlsx")
OUT  = os.path.join(BASE, "location", "granollers", "data.json")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

MESES_ROX = {'Jan':1,'Feb':2,'Mar':3,'Avr':4,'Apr':4,'May':5,'Jun':6,'July':7,'Jul':7,
             'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
MESES_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
            'Septiembre','Octubre','Noviembre','Diciembre']

# Estaciones de pesaje manual (MWS) → técnico real. Las ROXY son robots (sin persona).
PERSONAL_MWS = {
    'MWS_1': 'Manoli', 'MWS_2': 'Georgina', 'MWS_3': 'Luciano',
    'MWS_4': 'Francisco', 'MWS_6': 'Gemma',
}
def etiqueta_estacion(modulo):
    """'MWS_1' → 'Manoli (MWS_1)'; 'ROXY_A' se queda igual (es un robot)."""
    p = PERSONAL_MWS.get(modulo)
    return f'{p} ({modulo})' if p else modulo


def col_to_num(ref):
    """'B3' -> 2 (índice de columna 1-based)."""
    letters = re.match(r'[A-Z]+', ref).group()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def load_shared_strings(z):
    strings = []
    if 'xl/sharedStrings.xml' not in z.namelist():
        return strings
    for _, el in iterparse(io.BytesIO(z.read('xl/sharedStrings.xml'))):
        if el.tag == NS + 'si':
            strings.append(''.join(t.text or '' for t in el.iter(NS + 't')))
            el.clear()
    return strings


def read_sheet(z, filename, shared):
    """Devuelve una lista de filas; cada fila es {num_columna: valor}."""
    rows = []
    for _, el in iterparse(io.BytesIO(z.read(filename))):
        if el.tag == NS + 'row':
            cells = {}
            for c in el.findall(NS + 'c'):
                v = c.find(NS + 'v')
                if v is None:
                    continue
                val = shared[int(v.text)] if c.get('t') == 's' else v.text
                cells[col_to_num(c.get('r'))] = val
            rows.append(cells)
            el.clear()
    return rows


def sheet_map(z):
    """nombre de hoja -> ruta worksheets/sheetN.xml, respetando el orden del libro."""
    wb = z.read('xl/workbook.xml').decode('utf-8', 'ignore')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'ignore')
    rid_to_target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    out = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        tgt = rid_to_target.get(rid, '')
        if tgt:
            out[name] = 'xl/' + tgt.lstrip('/').replace('xl/', '')
    return out


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def es_codigo_mp(v):
    return bool(v and re.match(r'^\d{6,8}\s*-', str(v)))


# ---- Stats Samples: serie mensual ----
def parse_stats_samples(rows):
    """Columnas: E=año, F=mes, G=días, H=prodOrders, I=pours, J=sampleOrders, K=samples/día."""
    serie = []
    cur_year = None
    for r in rows:
        y = (r.get(5) or '').strip() if isinstance(r.get(5), str) else r.get(5)
        if y and re.match(r'^\d{4}$', str(y).strip()):
            cur_year = int(str(y).strip())
        mo = (r.get(6) or '').strip() if r.get(6) else ''
        if mo in MESES_ROX and cur_year:
            m = MESES_ROX[mo]
            prod = num(r.get(8))
            if prod is None:          # mes futuro sin datos: lo omitimos
                continue
            serie.append({
                'ym': f'{cur_year}-{m:02d}',
                'anio': str(cur_year),
                'mes': m,
                'label': MESES_ES[m - 1],
                'dias': num(r.get(7)),
                'prodOrders': prod,
                'pours': num(r.get(9)),
                'sampleOrders': num(r.get(10)),
                'samplesDia': num(r.get(11)),
            })
    return serie


# ---- Stats MP: top materias primas 2025 (dos tablas) ----
def parse_stats_mp(rows):
    def tabla(col_cod, col_pes, col_peso):
        out = []
        for r in rows:
            cod = r.get(col_cod)
            if not es_codigo_mp(cod):
                continue
            partes = str(cod).split(' - ', 1)
            out.append({
                'codigo': partes[0].strip(),
                'nombre': (partes[1].strip() if len(partes) > 1 else ''),
                'pesadas': num(r.get(col_pes)) or 0,
                'peso': round(num(r.get(col_peso)) or 0, 1),
            })
        out.sort(key=lambda x: -x['pesadas'])
        return out
    return {
        'produccion': tabla(1, 2, 3)[:25],       # tabla izquierda
        'sampleRequests': tabla(7, 8, 9)[:25],   # tabla derecha
    }


# ---- Exportación MANUAL: pesadas individuales (streaming, ~625k filas) ----
def parse_pesadas(z, sheet_path, shared):
    """Recorre la hoja de pesadas SIN cargarla en memoria. De la columna "Módulo"
    deriva manual (MWS_*) vs robot (ROXY_*). Devuelve agregados:
      · porMes:   {ym: {manual, robot}}  (nº de pesadas)
      · estaciones: {modulo: nº pesadas}
      · pesoVia:  {manual, robot}  (gramos dosificados, col. "Cantidad dosificada")
    Columnas (1-based): 1=fecha serial, 13=cantidad dosificada (gr), 18=Módulo.
    """
    from collections import defaultdict, Counter
    porMes = defaultdict(lambda: [0, 0])   # ym  -> [manual, robot]
    porDia = defaultdict(lambda: [0, 0])   # ymd -> [manual, robot]  (detalle diario)
    estaciones = Counter()
    pesoVia = [0.0, 0.0]                    # manual, robot
    # Detalle por estación MANUAL (MWS): pesadas, peso(gr), y OPs de muestra distintas
    pesoEst = defaultdict(float)                              # modulo -> gramos
    opsMesEst = defaultdict(lambda: defaultdict(set))         # modulo -> ym  -> set(OP)
    opsDiaEst = defaultdict(lambda: defaultdict(set))         # modulo -> ymd -> set(OP)
    # Desglose mensual por estación (para que las barras respeten el filtro de periodo)
    pesMesEst  = defaultdict(lambda: defaultdict(int))        # modulo -> ym -> nº pesadas (todas)
    pesoMesEst = defaultdict(lambda: defaultdict(float))      # modulo -> ym -> gramos (manual)
    # Variantes "solo MUESTRA" (TypeOF=Echantillon). La producción se deriva en el front
    # como (total − muestra). Así "Ambos" reutiliza los totales y no se duplica nada.
    porMesS    = defaultdict(lambda: [0, 0])                  # ym  -> [manual, robot] (muestra)
    porDiaS    = defaultdict(lambda: [0, 0])                  # ymd -> [manual, robot] (muestra)
    pesMesEstS  = defaultdict(lambda: defaultdict(int))       # modulo -> ym -> pesadas muestra
    pesoMesEstS = defaultdict(lambda: defaultdict(float))     # modulo -> ym -> gramos muestra (manual)
    EPOCH = datetime(1899, 12, 30)
    first = True
    with z.open(sheet_path) as fh:
        for _, el in iterparse(fh):
            if el.tag != NS + 'row':
                continue
            if first:                       # saltar cabecera
                first = False; el.clear(); continue
            cells = {}
            for c in el.findall(NS + 'c'):
                v = c.find(NS + 'v')
                if v is not None:
                    cells[col_to_num(c.get('r'))] = shared[int(v.text)] if c.get('t') == 's' else v.text
            el.clear()
            mod = cells.get(18)
            if not mod:
                continue
            es_robot = str(mod).startswith('ROXY')
            es_manual = mod in PERSONAL_MWS
            es_muestra = (cells.get(5) == 'Echantillon')   # tipo: muestra vs producción
            idx = 1 if es_robot else 0
            estaciones[mod] += 1
            # fecha desde el serial de Excel → agregación mensual y diaria
            ym = ymd = None
            try:
                d = EPOCH + timedelta(days=float(cells.get(1)))
                ym, ymd = d.strftime('%Y-%m'), d.strftime('%Y-%m-%d')
                porMes[ym][idx] += 1
                porDia[ymd][idx] += 1
                pesMesEst[mod][ym] += 1          # pesadas por estación y mes (todas)
                if es_muestra:
                    porMesS[ym][idx] += 1
                    porDiaS[ymd][idx] += 1
                    pesMesEstS[mod][ym] += 1
            except (TypeError, ValueError):
                pass
            dos = num(cells.get(13))
            if dos:
                pesoVia[idx] += dos
            # Detalle por estación manual
            if es_manual:
                if dos:
                    pesoEst[mod] += dos
                    if ym: pesoMesEst[mod][ym] += dos   # peso por estación y mes
                    if ym and es_muestra: pesoMesEstS[mod][ym] += dos
                # OPs de muestra distintas (TypeOF=Echantillon, col2=OP cliente)
                if es_muestra:
                    op = cells.get(2)
                    if op:
                        if ym:  opsMesEst[mod][ym].add(op)
                        if ymd: opsDiaEst[mod][ymd].add(op)

    # Lista ordenada de estaciones manuales (para ejes estables en el front)
    mws = [k for k in PERSONAL_MWS if estaciones.get(k)]
    mws.sort(key=lambda k: -estaciones[k])
    estManual = {
        'modulos': mws,
        'etiquetas': {k: etiqueta_estacion(k) for k in mws},
        'pesadas': {k: estaciones[k] for k in mws},
        'peso': {k: round(pesoEst[k], 1) for k in mws},
        # Desglose mensual por estación (las barras lo suman según el periodo activo)
        'pesadasMes': {k: dict(sorted(pesMesEst[k].items())) for k in mws},
        'pesoMes': {k: {ym: round(v, 1) for ym, v in sorted(pesoMesEst[k].items())} for k in mws},
        # Variante "solo muestra" (producción = total − muestra en el front)
        'pesadasMesSample': {k: dict(sorted(pesMesEstS[k].items())) for k in mws},
        'pesoMesSample': {k: {ym: round(v, 1) for ym, v in sorted(pesoMesEstS[k].items())} for k in mws},
        # OPs muestra distintas por mes y por día, por estación
        'ordenesMes': {k: {ym: len(s) for ym, s in sorted(opsMesEst[k].items())} for k in mws},
        'ordenesDia': {k: {ymd: len(s) for ymd, s in sorted(opsDiaEst[k].items())} for k in mws},
    }
    # Desglose mensual de pesadas por estación para el ranking (incluye robots ROXY)
    estacionesMes = {mod: dict(sorted(pesMesEst[mod].items())) for mod in estaciones}
    estacionesMesSample = {mod: dict(sorted(pesMesEstS[mod].items())) for mod in estaciones if pesMesEstS[mod]}
    return {
        'porMes': {ym: {'manual': v[0], 'robot': v[1]} for ym, v in porMes.items()},
        'porMesSample': {ym: {'manual': v[0], 'robot': v[1]} for ym, v in porMesS.items()},
        'serieDiaria': [{'ymd': d, 'manual': v[0], 'robot': v[1],
                         'manualSample': porDiaS[d][0], 'robotSample': porDiaS[d][1]}
                        for d, v in sorted(porDia.items())],
        'estaciones': [{'modulo': k, 'etiqueta': etiqueta_estacion(k), 'pesadas': n,
                        'persona': PERSONAL_MWS.get(k),
                        'tipo': ('robot' if k.startswith('ROXY') else 'manual')}
                       for k, n in estaciones.most_common()],
        'estacionesMes': estacionesMes,
        'estacionesMesSample': estacionesMesSample,
        'pesoVia': {'manual': round(pesoVia[0], 1), 'robot': round(pesoVia[1], 1)},
        'estManual': estManual,
    }


def build_scope(serie):
    """KPIs y series agregadas para un subconjunto de meses."""
    def s(k):
        return round(sum((m.get(k) or 0) for m in serie), 1)
    manual, robot = int(s('manual')), int(s('robot'))
    total_pes = manual + robot
    return {
        'kpis': {
            'prodOrders': int(s('prodOrders')),
            'pours': int(s('pours')),
            'sampleOrders': int(s('sampleOrders')),
            'dias': int(s('dias')),
            'samplesDia': round(s('sampleOrders') / s('dias'), 2) if s('dias') else 0,
            'manual': manual,
            'robot': robot,
            'pctRobot': round(robot / total_pes * 100, 1) if total_pes else 0,
        },
    }


def main():
    if not os.path.exists(XLSX):
        raise SystemExit(f"No se encuentra el Excel: {XLSX}")
    z = zipfile.ZipFile(XLSX)
    shared = load_shared_strings(z)
    smap = sheet_map(z)

    samples_rows = read_sheet(z, smap['Stats Samples'], shared)
    mp_rows = read_sheet(z, smap['Stats MP'], shared)

    serie = parse_stats_samples(samples_rows)
    mp = parse_stats_mp(mp_rows)

    # Pesadas individuales (manual vs robot) por streaming de la hoja grande
    print("  Procesando 'Exportación MANUAL' por streaming (~625k filas)…")
    pes = parse_pesadas(z, smap['Exportación MANUAL'], shared)
    # Enriquecer la serie mensual con manual/robot del mismo ym (+ variante solo muestra)
    for m in serie:
        mr = pes['porMes'].get(m['ym'], {'manual': 0, 'robot': 0})
        m['manual'] = mr['manual']
        m['robot'] = mr['robot']
        ms = pes['porMesSample'].get(m['ym'], {'manual': 0, 'robot': 0})
        m['manualSample'] = ms['manual']
        m['robotSample'] = ms['robot']

    anios = sorted(set(m['anio'] for m in serie))
    meses = [{'ym': m['ym'], 'anio': m['anio'], 'label': m['label']} for m in serie]

    scopes = {'ALL': build_scope(serie)}
    for a in anios:
        scopes[a] = build_scope([m for m in serie if m['anio'] == a])
    for m in serie:
        scopes[m['ym']] = build_scope([m])

    out = {
        'generado': datetime.now().isoformat(timespec='seconds'),
        'centro': 'Granollers',
        'fuente': 'Roxane',
        'modelo': 'samples',                # marca el tipo de centro para el front
        'anios': anios,
        'meses': meses,
        'scopes': scopes,
        'serieMensual': serie,
        'materiasPrimas': mp,
        'estaciones': pes['estaciones'],     # ranking de puestos de pesaje (MWS/ROXY)
        'estacionesMes': pes['estacionesMes'],  # pesadas por estación y mes (filtrable por periodo)
        'estacionesMesSample': pes['estacionesMesSample'],  # variante solo muestra (prod = total − muestra)
        'pesoVia': pes['pesoVia'],            # gramos dosificados manual vs robot
        'serieDiaria': pes['serieDiaria'],   # pesadas manual/robot por día (detalle)
        'estManual': pes['estManual'],       # detalle por estación manual (pesadas/peso/órdenes muestra mes+día)
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(',', ':'))

    k = scopes['ALL']['kpis']
    print(f"Escrito {OUT} ({os.path.getsize(OUT)} bytes)")
    print(f"  Años: {anios} · {len(serie)} meses con datos")
    print(f"  Total: {k['prodOrders']:,} órdenes prod · {k['pours']:,} pours · "
          f"{k['sampleOrders']:,} muestras")
    print(f"  Pesadas: {k['manual']:,} manual + {k['robot']:,} robot "
          f"({k['pctRobot']}% robot) · {len(pes['estaciones'])} estaciones")
    print(f"  Peso por vía (kg): manual={pes['pesoVia']['manual']/1000:,.0f} · "
          f"robot={pes['pesoVia']['robot']/1000:,.0f}")
    print(f"  Materias primas: {len(mp['produccion'])} (prod) · "
          f"{len(mp['sampleRequests'])} (samples)")


if __name__ == '__main__':
    main()
