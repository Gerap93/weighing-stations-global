#!/usr/bin/env python3
"""
Genera data/data.json a partir de los CSV en location/<centro>/.
Replica exactamente la lógica del modelo de Power BI (Power Query + medidas DAX).

Uso:
    python build_data.py
Lee:  ./location/*/*.csv   (una subcarpeta por centro; separador ';', encoding latin-1)
Escribe: ./data/data.json

El HTML lee ese data.json por fetch. Para actualizar el dashboard,
basta con reemplazar/añadir CSV en location/<centro>/ y volver a ejecutar este script.

NOTA: de momento se consolidan todos los centros juntos (mismo resultado que antes).
La separación/comparación por centro y la lectura de .xlsx están pendientes de diseño.
"""
import csv, glob, json, os, hashlib
from collections import defaultdict
from datetime import datetime

SRC_DIR  = os.path.join(os.path.dirname(__file__), "location")
OUT_PATH = os.path.join(os.path.dirname(__file__), "data", "data.json")

# Archivos que el modelo original excluía explícitamente
EXCLUIR = {
    "ANAFORM.csv", "BSX.csv", "Datastock.csv", "fORMLABES.csv", "outper.csv",
    "RData2025.csv", "RData2026.csv",
    "Recuento de Código fórmula por Mes y Formula resp..csv",
    "Recuento de Código ingrediente por Mes y Ingredient resp..csv",
    "TData2025.csv", "TData2026.csv",
}

MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto',
         'septiembre','octubre','noviembre','diciembre']


def num(s):
    if not s or not s.strip():
        return 0.0
    s = s.strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(s):
    s = (s or '').strip()
    if not s:
        return None
    # Tomamos solo la parte de fecha (algunos exports añaden la hora: "15/1/24 13:15").
    fecha = s.split(' ')[0]
    # Excel reescribe los CSV con su formato local al editarlos, así que toleramos
    # tanto año de 4 dígitos (dd/mm/yyyy) como de 2 (d/m/yy). strptime ya admite
    # día/mes sin cero a la izquierda.
    for fmt in ('%d/%m/%Y', '%d/%m/%y'):
        try:
            return datetime.strptime(fecha, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def clasificar(cod, nom, resp_ing, form_resp=''):
    """Replica las columnas calculadas del Power Query.

    Regla especial ALANGLOIS (Dubai): sus fórmulas marcan el tipo con el sufijo BVF/TVF
    en el nombre en vez de las palabras BASE/TRIAL. Cuando el 'Formula resp.' es ALANGLOIS
    y el nombre contiene BVF → Base, TVF → Trial, y PREVALECE sobre BASE/TRIAL (muchos
    nombres llevan 'E-BASE' pero es el sufijo BVF/TVF el que dice si esa línea es la Base
    o el Trial de esa familia; ej. 'CACAO BEAN E-BASE TVF - 2' es un Trial, no una Base).
    """
    codU = cod.upper()
    nomU = nom.upper()
    es_bxs = 1 if codU.startswith('BXS') else 0
    alanglois = ((form_resp or '').strip().upper() == 'ALANGLOIS'
                 and ('BVF' in nomU or 'TVF' in nomU))
    if es_bxs:
        es_base = es_trial = es_completa = 0
        tipo_det = 'Disoluciones'
    elif alanglois:
        es_base = 1 if 'BVF' in nomU else 0
        es_trial = 1 if 'TVF' in nomU else 0
        es_completa = 0
        tipo_det = 'Base + Trial' if (es_base and es_trial) else ('Base' if es_base else 'Trial')
    else:
        es_base = 1 if 'BASE' in nomU else 0
        es_trial = 1 if 'TRIAL' in nomU else 0
        es_completa = 1 if ('BASE' not in nomU and 'TRIAL' not in nomU) else 0
        if 'BASE' in nomU and 'TRIAL' in nomU:
            tipo_det = 'Base + Trial'
        elif 'BASE' in nomU:
            tipo_det = 'Base'
        elif 'TRIAL' in nomU:
            tipo_det = 'Trial'
        else:
            tipo_det = 'Fórmula completa'

    tipo_linea = 'Robot' if resp_ing.strip() == '' else 'Lab'
    return es_bxs, es_base, es_trial, es_completa, tipo_det, tipo_linea


def cargar_filas():
    """Devuelve (rows, meta) donde meta lleva info de integridad por archivo."""
    rows = []
    archivos = []          # info por archivo leído
    hashes = {}            # hash de contenido -> primer nombre que lo trajo
    files = sorted(f for f in glob.glob(os.path.join(SRC_DIR, "*", "*.csv")) if (lambda n: len(n) > 6 and n[:6].isdigit() and n[6] == "-")(os.path.basename(f)))
    if not files:
        raise SystemExit(f"No hay CSV en {SRC_DIR}/<centro>/")
    for path in files:
        name = os.path.basename(path)
        if name in EXCLUIR:
            continue
        raw = open(path, 'rb').read()
        h = hashlib.md5(raw).hexdigest()
        dup_de = hashes.get(h)          # ¿contenido idéntico a otro archivo ya leído?
        if dup_de is None:
            hashes[h] = name

        with open(path, 'r', encoding='latin-1', newline='') as fh:
            rdr = csv.DictReader(fh, delimiter=';')
            req = {'Código fórmula', 'Nombre fórmula', 'Peso ingr. [g]',
                   'Código ingrediente', 'Ingredient resp.', 'Fin de prod.'}
            if not req.issubset(set(rdr.fieldnames or [])):
                print(f"  · aviso: {name} no tiene las columnas esperadas, se omite")
                archivos.append({'nombre': name, 'filas': 0, 'omitido': True,
                                 'motivo': 'columnas incorrectas'})
                continue

            n_archivo = 0
            for r in rdr:
                cod = (r.get('Código fórmula') or '').strip()
                nom = (r.get('Nombre fórmula') or '').strip()
                resp_ing = (r.get('Ingredient resp.') or '').strip()
                form_resp = (r.get('Formula resp.') or '').strip()
                es_bxs, es_base, es_trial, es_completa, tipo_det, tipo_linea = \
                    clasificar(cod, nom, resp_ing, form_resp)
                fin = parse_date(r.get('Fin de prod.'))
                rows.append({
                    'archivo': name,
                    'formulaKey': cod,
                    'nombreFormula': nom,
                    'codIngrediente': (r.get('Código ingrediente') or '').strip(),
                    'nombreIngrediente': (r.get('Nombre ingrediente') or '').strip(),
                    'pesoIngr': num(r.get('Peso ingr. [g]')),
                    'lote': (r.get('Número de lote') or '').strip(),
                    'formulaResp': form_resp,
                    'ingredientResp': resp_ing,
                    'finProd': fin,
                    'finProdRaw': (r.get('Fin de prod.') or '').strip(),
                    'tipoFormulaDetalle': tipo_det,
                    'tipoLinea': tipo_linea,
                    'esBase': es_base, 'esTrial': es_trial,
                    'esCompleta': es_completa, 'bxs': es_bxs,
                })
                n_archivo += 1
            archivos.append({'nombre': name, 'filas': n_archivo, 'omitido': False,
                             'duplicadoDe': dup_de})

    # Deduplicación por clave de negocio (fórmula+ingrediente+lote+fecha+peso).
    # Los CSV semanales se solapan al re-exportarse (un export puede repetir días
    # ya incluidos en otro). Conservamos la 1ª aparición —archivo más antiguo por
    # orden de nombre— y descartamos las repetidas, para no inflar líneas ni peso.
    vistos = set()
    unicas = []
    for r in rows:
        clave = (r['formulaKey'], r['codIngrediente'], r['lote'],
                 r['finProdRaw'], r['pesoIngr'])
        if clave in vistos:
            continue
        vistos.add(clave)
        unicas.append(r)
    dup_eliminadas = len(rows) - len(unicas)
    sin_fecha = sum(1 for r in unicas if not r['finProd'])
    return unicas, {'archivos': archivos,
                    'dupFilasEliminadas': dup_eliminadas,
                    'filasSinFecha': sin_fecha}




def ym(iso):
    return iso[:7] if iso else None


def build_scope(subset):
    def dc(pred=lambda r: True):
        return len(set(r['formulaKey'] for r in subset if pred(r)))

    perf = defaultdict(lambda: {'fk': set(), 'lineas': 0, 'comp': set(),
                                'trial': set(), 'base': set(), 'diso': set()})
    lab = defaultdict(lambda: {'lineas': 0, 'peso': 0.0, 'fk': set()})
    tdet = defaultdict(set)
    ingr = defaultdict(lambda: {'lineas': 0, 'peso': 0.0, 'nombre': ''})
    # Trazabilidad: agregación por (código ingrediente, lote)
    lotes = defaultdict(lambda: {'lineas': 0, 'peso': 0.0, 'fk': set(),
                                 'nombre': '', 'cod': ''})

    for r in subset:
        fk = r['formulaKey']
        if r['formulaResp']:
            p = perf[r['formulaResp']]
            p['fk'].add(fk); p['lineas'] += 1
            if r['esCompleta']: p['comp'].add(fk)
            if r['esTrial']:    p['trial'].add(fk)
            if r['esBase']:     p['base'].add(fk)
            if r['bxs']:        p['diso'].add(fk)
        if r['ingredientResp']:
            l = lab[r['ingredientResp']]
            l['lineas'] += 1; l['peso'] += r['pesoIngr']; l['fk'].add(fk)
        tdet[r['tipoFormulaDetalle']].add(fk)
        iu = ingr[r['codIngrediente']]
        iu['lineas'] += 1; iu['peso'] += r['pesoIngr']; iu['nombre'] = r['nombreIngrediente']
        if r['lote']:
            lt = lotes[(r['codIngrediente'], r['lote'])]
            lt['lineas'] += 1; lt['peso'] += r['pesoIngr']; lt['fk'].add(fk)
            lt['nombre'] = r['nombreIngrediente']; lt['cod'] = r['codIngrediente']

    return {
        'kpis': {
            'lineas': len(subset),
            'ingredientes': len(set(r['codIngrediente'] for r in subset)),
            'formulas': dc(),
            'completas': dc(lambda r: r['esCompleta'] == 1),
            'bases': dc(lambda r: r['esBase'] == 1),
            'trials': dc(lambda r: r['esTrial'] == 1),
            'disoluciones': dc(lambda r: r['bxs'] == 1),
            'lab': sum(1 for r in subset if r['tipoLinea'] == 'Lab'),
            'robot': sum(1 for r in subset if r['tipoLinea'] == 'Robot'),
            'peso': round(sum(r['pesoIngr'] for r in subset), 1),
        },
        'tiposDetalle': sorted(
            [{'tipo': k, 'formulas': len(v)} for k, v in tdet.items()],
            key=lambda x: -x['formulas']),
        'perfumers': sorted(
            [{'nombre': k, 'formulas': len(v['fk']), 'lineas': v['lineas'],
              'completas': len(v['comp']), 'trials': len(v['trial']),
              'bases': len(v['base']), 'disoluciones': len(v['diso'])}
             for k, v in perf.items()],
            key=lambda x: -x['formulas']),
        'lab': sorted(
            [{'nombre': k, 'lineas': v['lineas'], 'peso': round(v['peso'], 1),
              'formulas': len(v['fk'])} for k, v in lab.items()],
            key=lambda x: -x['lineas']),
        'topIngredientes': sorted(
            [{'codigo': k, 'nombre': v['nombre'], 'lineas': v['lineas'],
              'peso': round(v['peso'], 1)} for k, v in ingr.items()],
            key=lambda x: -x['lineas'])[:25],
        'lotes': sorted(
            [{'cod': v['cod'], 'nombre': v['nombre'], 'lote': lote,
              'lineas': v['lineas'], 'peso': round(v['peso'], 1),
              'formulas': len(v['fk'])} for (cod, lote), v in lotes.items()],
            key=lambda x: -x['lineas'])[:80],
    }


def build_proyectos(rows):
    """Una entrada por fórmula única (FormulaKey) para la página de proyectos.
    Toma el primer nombre/fecha no vacíos y agrega líneas y peso. También registra
    quién PESÓ cada fórmula (Ingredient resp. = técnico de lab): el 98% son de un
    solo técnico, así que se guarda el principal (el que más líneas pesó) + nº de
    técnicos distintos, para señalar los casos con varios."""
    from collections import Counter
    proy = {}
    pesadores = {}                      # fk -> Counter(técnico -> nº líneas)
    for r in rows:
        fk = r['formulaKey']
        if not fk:
            continue
        p = proy.get(fk)
        if p is None:
            p = proy[fk] = {
                'codigo': fk, 'nombre': r['nombreFormula'],
                'resp': r['formulaResp'], 'tipo': r['tipoFormulaDetalle'],
                'fecha': r['finProd'], 'ym': ym(r['finProd']),
                'lineas': 0, 'peso': 0.0,
            }
            pesadores[fk] = Counter()
        p['lineas'] += 1
        p['peso'] += r['pesoIngr']
        if r['ingredientResp']:
            pesadores[fk][r['ingredientResp']] += 1
        if not p['nombre'] and r['nombreFormula']:
            p['nombre'] = r['nombreFormula']
        if not p['resp'] and r['formulaResp']:
            p['resp'] = r['formulaResp']
        if not p['fecha'] and r['finProd']:
            p['fecha'] = r['finProd']; p['ym'] = ym(r['finProd'])
    for fk, p in proy.items():
        p['peso'] = round(p['peso'], 1)
        c = pesadores[fk]
        # técnico principal (más líneas pesadas) + cuántos técnicos distintos
        p['tecnico'] = c.most_common(1)[0][0] if c else ''
        p['nTecnicos'] = len(c)
    return sorted(proy.values(), key=lambda x: -x['lineas'])


def control_integridad(rows, meta):
    """Resumen de integridad que se incrusta en data.json y se imprime en consola.
    Las filas duplicadas por solapamiento de CSV ya se eliminaron en cargar_filas;
    aquí solo se informa de cuántas eran (saneamiento automático, no es un fallo).
    Las filas sin fecha válida sí se consideran una alerta real. `rows` ya viene
    deduplicado."""
    archivos = meta['archivos']
    leidos = [a for a in archivos if not a.get('omitido')]

    # 1) Archivos con contenido idéntico a otro (subida repetida del mismo CSV)
    archivos_dup = [a['nombre'] for a in leidos if a.get('duplicadoDe')]
    # 2) Filas duplicadas por clave de negocio (ya eliminadas en cargar_filas)
    dup_filas = meta.get('dupFilasEliminadas', 0)
    # 3) Filas sin fecha válida en 'Fin de prod.'
    sin_fecha = meta.get('filasSinFecha', 0)

    # Alertas REALES: ponen el indicador del header en naranja.
    alertas = []
    if archivos_dup:
        alertas.append(f"{len(archivos_dup)} archivo(s) con contenido idéntico a otro ya cargado: "
                       + ", ".join(archivos_dup))
    if sin_fecha:
        alertas.append(f"{sin_fecha} fila(s) sin fecha válida en 'Fin de prod.'")

    # Saneamientos informativos: se resolvieron solos, no son un fallo.
    info = []
    if dup_filas:
        info.append(f"{dup_filas} fila(s) duplicada(s) por solapamiento de CSV, eliminadas automáticamente")

    return {
        'archivos': len(leidos),
        'filas': len(rows),                 # filas reales tras deduplicar
        'archivosDuplicados': archivos_dup,
        'filasDuplicadas': dup_filas,
        'filasSinFecha': sin_fecha,
        'alertas': alertas,
        'info': info,
        'ok': len(alertas) == 0,
        'detalleArchivos': [{'nombre': a['nombre'], 'filas': a['filas']}
                            for a in leidos],
    }


def main():
    rows, meta = cargar_filas()
    integridad = control_integridad(rows, meta)
    print(f"Archivos leídos: {integridad['archivos']} · Filas procesadas: {len(rows)}")

    # Avisos de integridad
    for i in integridad.get('info', []):
        print(f"  · {i}")
    if integridad['ok']:
        print("  ✓ Integridad OK: sin alertas")
    else:
        for a in integridad['alertas']:
            print(f"  ⚠ AVISO: {a}")

    meses_set = sorted(set(ym(r['finProd']) for r in rows if r['finProd']))
    out = {
        'generado': datetime.now().isoformat(timespec='seconds'),
        'integridad': integridad,
        'meses': [{'ym': y, 'anio': y[:4], 'label': MESES[int(y[5:7]) - 1].capitalize()}
                  for y in meses_set],
        'scopes': {'ALL': build_scope(rows)},
        'anios': sorted(set((ym(r['finProd']) or '')[:4] for r in rows if r['finProd'])), 'proyectos': build_proyectos(rows),
        'serieMensual': [],
    }
    for y in meses_set:
        sub = [r for r in rows if ym(r['finProd']) == y]
        out['scopes'][y] = build_scope(sub)
        out['serieMensual'].append({
            'label': MESES[int(y[5:7]) - 1].capitalize(),
            'lineas': len(sub),
            'formulas': len(set(r['formulaKey'] for r in sub)),
            'lab': sum(1 for r in sub if r['tipoLinea'] == 'Lab'),
            'robot': sum(1 for r in sub if r['tipoLinea'] == 'Robot'),
            'peso': round(sum(r['pesoIngr'] for r in sub), 1), 'ym': y, 'anio': y[:4],
        })

    out['scopes'].update({_a: build_scope([r for r in rows if (ym(r['finProd']) or '').startswith(_a)]) for _a in out['anios']})
    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(',', ':'))

    k = out['scopes']['ALL']['kpis']
    print(f"Escrito {OUT_PATH} ({os.path.getsize(OUT_PATH)} bytes)")
    print(f"  Meses: {[m['label'] for m in out['meses']]}")
    print(f"  Fórmulas: {k['formulas']} · Líneas: {k['lineas']} · "
          f"Completas: {k['completas']} · Trials: {k['trials']} · Bases: {k['bases']}")


if __name__ == '__main__':
    main()
