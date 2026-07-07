#!/usr/bin/env python3
"""
Genera location/dubai/data.json a partir del CSV de Dubai.

Dubai usa el MISMO modelo que Séneca (Dory): fórmulas, líneas, lab/robot, ingredientes
por fórmula. Por eso reutiliza casi toda la lógica de build_data.py (clasificación,
scopes, proyectos, integridad). Solo cambia la LECTURA del CSV, porque el export de
Dubai difiere del de Séneca:
  · columnas en INGLÉS (Formula code, Ingr. weight [g], Ingredient resp., Prod. end…)
  · encoding UTF-8 (Séneca es latin-1)
  · números en formato US (punto decimal: "8.35", "103.333") → float() directo
  · fecha dd/mm/yy con hora ("10/9/24 14:59") → parse_date de Séneca ya lo soporta
  · un único CSV grande (no semanales), nombre DUBxxxxxx-xxxxxx.csv

Uso:
    python build_dubai.py
Lee:  ./location/dubai/*.csv
Escribe: ./location/dubai/data.json
"""
import csv, glob, json, os
from datetime import datetime

# Reutilizar la lógica de negocio de Séneca (mismo modelo "formulas")
import build_data as B

BASE = os.path.dirname(__file__)
SRC_DIR = os.path.join(BASE, "location", "dubai")
OUT = os.path.join(BASE, "location", "dubai", "data.json")

# Mapeo columna interna ← cabecera del CSV de Dubai (inglés)
COLS = {
    'cod':       'Formula code',
    'nom':       'Formula name',
    'pesoIngr':  'Ingr. weight [g]',
    'codIngr':   'Ingredient code',
    'nomIngr':   'Ingredient name',
    'lote':      'Lot number',
    'formResp':  'Formula resp.',
    'ingrResp':  'Ingredient resp.',
    'finProd':   'Prod. end',
}


def num_us(s):
    """Número en formato US (punto decimal, sin separador de miles)."""
    s = (s or '').strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def cargar_filas():
    """Lee el/los CSV de Dubai y devuelve (rows, meta), misma forma que build_data."""
    rows = []
    archivos = []
    files = sorted(glob.glob(os.path.join(SRC_DIR, "*.csv")))
    if not files:
        raise SystemExit(f"No hay CSV en {SRC_DIR}")
    req = set(COLS.values())
    for path in files:
        name = os.path.basename(path)
        with open(path, 'r', encoding='utf-8-sig', newline='') as fh:
            rdr = csv.DictReader(fh, delimiter=';')
            if not req.issubset(set(rdr.fieldnames or [])):
                faltan = req - set(rdr.fieldnames or [])
                print(f"  · aviso: {name} no tiene las columnas esperadas (faltan {faltan}), se omite")
                archivos.append({'nombre': name, 'filas': 0, 'omitido': True,
                                 'motivo': 'columnas incorrectas'})
                continue
            n_archivo = 0
            for r in rdr:
                cod = (r.get(COLS['cod']) or '').strip()
                nom = (r.get(COLS['nom']) or '').strip()
                resp_ing = (r.get(COLS['ingrResp']) or '').strip()
                form_resp = (r.get(COLS['formResp']) or '').strip()
                es_bxs, es_base, es_trial, es_completa, tipo_det, tipo_linea = \
                    B.clasificar(cod, nom, resp_ing, form_resp)
                fin = B.parse_date(r.get(COLS['finProd']))
                rows.append({
                    'archivo': name,
                    'formulaKey': cod,
                    'nombreFormula': nom,
                    'codIngrediente': (r.get(COLS['codIngr']) or '').strip(),
                    'nombreIngrediente': (r.get(COLS['nomIngr']) or '').strip(),
                    'pesoIngr': num_us(r.get(COLS['pesoIngr'])),
                    'lote': (r.get(COLS['lote']) or '').strip(),
                    'formulaResp': form_resp,
                    'ingredientResp': resp_ing,
                    'finProd': fin,
                    'finProdRaw': (r.get(COLS['finProd']) or '').strip(),
                    'tipoFormulaDetalle': tipo_det,
                    'tipoLinea': tipo_linea,
                    'esBase': es_base, 'esTrial': es_trial,
                    'esCompleta': es_completa, 'bxs': es_bxs,
                })
                n_archivo += 1
            archivos.append({'nombre': name, 'filas': n_archivo, 'omitido': False,
                             'duplicadoDe': None})

    # Misma deduplicación por clave de negocio que Séneca
    vistos = set()
    unicas = []
    for r in rows:
        clave = (r['formulaKey'], r['codIngrediente'], r['lote'],
                 r['finProdRaw'], r['pesoIngr'])
        if clave in vistos:
            continue
        vistos.add(clave)
        unicas.append(r)
    return unicas, {'archivos': archivos,
                    'dupFilasEliminadas': len(rows) - len(unicas),
                    'filasSinFecha': sum(1 for r in unicas if not r['finProd'])}


def main():
    rows, meta = cargar_filas()
    integridad = B.control_integridad(rows, meta)
    print(f"Archivos leídos: {integridad['archivos']} · Filas procesadas: {len(rows)}")
    for i in integridad.get('info', []):
        print(f"  · {i}")
    if integridad['ok']:
        print("  ✓ Integridad OK: sin alertas")
    else:
        for a in integridad['alertas']:
            print(f"  ⚠ AVISO: {a}")

    meses_set = sorted(set(B.ym(r['finProd']) for r in rows if r['finProd']))
    out = {
        'generado': datetime.now().isoformat(timespec='seconds'),
        'centro': 'Dubai',
        'fuente': 'Dory',
        'modelo': 'formulas',
        'integridad': integridad,
        'meses': [{'ym': y, 'anio': y[:4], 'label': B.MESES[int(y[5:7]) - 1].capitalize()}
                  for y in meses_set],
        'scopes': {'ALL': B.build_scope(rows)},
        'anios': sorted(set((B.ym(r['finProd']) or '')[:4] for r in rows if r['finProd'])),
        'proyectos': B.build_proyectos(rows),
        'serieMensual': [],
    }
    for y in meses_set:
        sub = [r for r in rows if B.ym(r['finProd']) == y]
        out['scopes'][y] = B.build_scope(sub)
        out['serieMensual'].append({
            'label': B.MESES[int(y[5:7]) - 1].capitalize(),
            'lineas': len(sub),
            'formulas': len(set(r['formulaKey'] for r in sub)),
            'lab': sum(1 for r in sub if r['tipoLinea'] == 'Lab'),
            'robot': sum(1 for r in sub if r['tipoLinea'] == 'Robot'),
            'peso': round(sum(r['pesoIngr'] for r in sub), 1), 'ym': y, 'anio': y[:4],
        })
    out['scopes'].update({_a: B.build_scope([r for r in rows if (B.ym(r['finProd']) or '').startswith(_a)])
                          for _a in out['anios']})

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(',', ':'))

    k = out['scopes']['ALL']['kpis']
    print(f"Escrito {OUT} ({os.path.getsize(OUT)} bytes)")
    print(f"  Años: {out['anios']} · {len(out['meses'])} meses")
    print(f"  Fórmulas: {k['formulas']} · Líneas: {k['lineas']} · "
          f"Completas: {k['completas']} · Trials: {k['trials']} · Bases: {k['bases']}")


if __name__ == '__main__':
    main()
