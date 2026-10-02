"""Carga mensual de Movilidad Social (Crédito de Desarrollo Humano, registro 12)
en KPIs · Protección Social e Inclusión Económica.

Entrada: el Excel «resumen_CDH_<mes>_<año>.xlsx» (hojas Tipo credito, Genero, Rango edad, Etnia, Pobreza;
columnas Año, Mes, tipocredito, <desagregación>, registros, monto_total, monto_promedio).

Actualiza, en datos/kpis/:
  movilidad_social/cobertura_cdh_movilidad.csv        -> agrega (o reemplaza) las filas del mes
  movilidad_social/caracterizacion_cdh_movilidad.csv  -> agrega (o reemplaza) las filas del mes
  proteccion_social/series_historicas_servicios.csv   -> el mes queda como dato real (registro 12)
  kpi_proteccion_social.csv                           -> cifra destacada y textos de la tarjeta (registro 12)
Los meses anteriores se conservan; el portal muestra siempre el último mes cargado.

Uso (desde la raíz del repositorio):
    python scripts/actualizar_cdh.py --archivo "C:/cdh/resumen_CDH_octubre_2026.xlsx"

Si alguna desagregación no suma el total de su tipo de crédito, no escribe nada.
Requiere: pip install openpyxl
"""
import argparse, csv, io, os, re, sys

import openpyxl

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
D = lambda p: os.path.join(RAIZ, 'datos/kpis', p)
COB, CAR = D('movilidad_social/cobertura_cdh_movilidad.csv'), D('movilidad_social/caracterizacion_cdh_movilidad.csv')
SERIE, TARJ = D('proteccion_social/series_historicas_servicios.csv'), D('kpi_proteccion_social.csv')
MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
         'septiembre', 'octubre', 'noviembre', 'diciembre']
GRUPOS = {'Genero': ('genero', 'genero'), 'Rango edad': ('rango_edad', 'rangoedad'),
          'Etnia': ('etnia', 'etnia'), 'Pobreza': ('pobreza', 'pobreza')}


def leer_csv(p):
    raw = open(p, 'rb').read()
    nl = '\r\n' if b'\r\n' in raw else '\n'
    filas = list(csv.DictReader(io.StringIO(raw.decode('utf-8-sig'))))
    return filas, list(filas[0].keys()), nl


def escribir_csv(p, filas, cab, nl):
    b = io.StringIO()
    w = csv.DictWriter(b, fieldnames=cab, lineterminator=nl)
    w.writeheader()
    w.writerows(filas)
    open(p, 'wb').write(b.getvalue().encode('utf-8'))


def num(v):
    if v in (None, ''):
        return ''
    t = f'{float(v):.2f}'.rstrip('0').rstrip('.')
    return t


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--archivo', required=True)
    a = ap.parse_args()
    wb = openpyxl.load_workbook(a.archivo, data_only=True)
    hojas = {}
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(values_only=True))
        cab = [str(c).strip() for c in filas[0]]
        hojas[ws.title.strip()] = [dict(zip(cab, f)) for f in filas[1:] if f and f[0]]
    for h in ['Tipo credito', *GRUPOS]:
        if h not in hojas:
            sys.exit(f'Falta la hoja «{h}»')
    tipo = hojas['Tipo credito']
    periodos = {(int(r['Año']), str(r['Mes']).strip().lower()) for r in tipo}
    if len(periodos) != 1:
        sys.exit(f'El Excel debe traer un solo mes; trae {sorted(periodos)}')
    anio, nom_mes = periodos.pop()
    if nom_mes not in MESES:
        sys.exit(f'Mes no reconocido: {nom_mes}')
    anio, mes = str(anio), f'{MESES.index(nom_mes) + 1:02d}'
    fuente = os.path.basename(a.archivo)
    nota = f'Dato real ({fuente}).'

    # validar sumas por tipo de crédito
    tot = {str(r['tipocredito']).strip(): int(r['registros']) for r in tipo}
    errores = []
    for hoja, (g, col) in GRUPOS.items():
        for t, n in tot.items():
            s = sum(int(r['registros']) for r in hojas[hoja] if str(r['tipocredito']).strip() == t)
            if s != n:
                errores.append(f'{hoja} · {t}: suma {s} ≠ total {n}')
    if errores:
        sys.exit('No se escribió nada. Las desagregaciones no cuadran:\n  ' + '\n  '.join(errores))

    # cobertura
    C, cab_c, nl_c = leer_csv(COB)
    indicador = C[0]['indicador']
    C = [r for r in C if (r['anio'], r['mes']) != (anio, mes)]
    for r in tipo:
        C.append({'registro_id': '12', 'indicador': indicador, 'anio': anio, 'mes': mes,
                  'tipo_credito': str(r['tipocredito']).strip(), 'registros': str(int(r['registros'])),
                  'monto_total_usd': num(r['monto_total']), 'monto_promedio_usd': num(r['monto_promedio']),
                  'es_dato_real': 'TRUE', 'nota': nota})
    C.sort(key=lambda r: (r['anio'], r['mes'], r['tipo_credito']))

    # caracterización
    K, cab_k, nl_k = leer_csv(CAR)
    K = [r for r in K if (r['anio'], r['mes']) != (anio, mes)]
    for hoja, (g, col) in GRUPOS.items():
        for r in sorted(hojas[hoja], key=lambda r: str(r['tipocredito'])):
            K.append({'registro_id': '12', 'anio': anio, 'mes': mes, 'tipo_credito': str(r['tipocredito']).strip(),
                      'grupo': g, 'categoria': str(r[col]).strip(), 'registros': str(int(r['registros'])),
                      'monto_total_usd': num(r['monto_total']), 'es_dato_real': 'TRUE', 'nota': nota})
    K.sort(key=lambda r: (r['anio'], r['mes']))  # estable: conserva el orden de grupos dentro del mes

    # serie histórica (registro 12)
    total = sum(tot.values())
    S, cab_s, nl_s = leer_csv(SERIE)
    fila = next((r for r in S if r['registro_id'] == '12' and r['anio'] == anio and r['mes'] == mes), None)
    nota_s = f'Dato real, {nom_mes} {anio} ({fuente}).'
    if fila:
        fila.update(usuarios=str(total), es_dato_real='TRUE', nota=nota_s)
    else:
        idx = max(i for i, r in enumerate(S) if r['registro_id'] == '12')
        S.insert(idx + 1, dict(S[idx], anio=anio, mes=mes, usuarios=str(total), es_dato_real='TRUE', nota=nota_s))

    # tarjeta del panorama (registro 12)
    T, cab_t, nl_t = leer_csv(TARJ)
    fmt = f'{total:,}'.replace(',', '.')
    patron = re.compile(r'\b(' + '|'.join(MESES) + r') \d{4}', re.I)
    for r in T:
        if r['registro_id'] != '12':
            continue
        r['headline'] = fmt
        r['headline_sufijo'] = patron.sub(f'{nom_mes} {anio}', r['headline_sufijo'])
        r['nota'] = patron.sub(f'{nom_mes} {anio}', r['nota'])
        if r['tipo_registro'] == 'desagregacion_actual':
            r['desagregacion_categoria'] = f'{nom_mes.capitalize()} {anio}'
            r['desagregacion_valor'] = fmt

    escribir_csv(COB, C, cab_c, nl_c)
    escribir_csv(CAR, K, cab_k, nl_k)
    escribir_csv(SERIE, S, cab_s, nl_s)
    escribir_csv(TARJ, T, cab_t, nl_t)
    print(f'CDH {nom_mes} {anio}: {total} créditos ({", ".join(f"{t}: {n}" for t, n in tot.items())})')
    print('Listo. Revisar MDTDH/kpis.html#/cobertura-movilidad-social')


if __name__ == '__main__':
    main()
