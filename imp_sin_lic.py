#!/usr/bin/env python3
"""
imp_sin_lic.py — Imponible sin licencia (IMP SIN LIC)

Para cada FICHA del mes en proceso con DIAS LICENCIA > 0, busca hacia atrás
en los meses anteriores el primer mes en que la misma FICHA tuvo
DIAS LICENCIA == 0. En ese mes suma los haberes afectos (según
Equivalencias Tw.xlsx) y toma el menor entre esa suma y el
topeImp_pesos_afp del mes encontrado (parametrosMesuales.xlsx).

Si no encuentra ningún mes sin licencia, el imponible puede ingresarse a
mano; si no se ingresa, queda "Imp no encontrado" y el SIS se calcula con
el Sueldo Base del listado de empleados.

Salidas:
  * Copia del archivo de entrada con la columna IMP SIN LIC insertada
    inmediatamente a la derecha de DIAS LICENCIA (cero por defecto).
  * Informe ImpSinLic<MES AAAA>.xlsx con RUT, FICHA, CONTRATO,
    DIAS LICENCIA, ULT IMP SIN LIC e IMP IMP SIS = (ULT IMP SIN LIC / 30) * DIAS LICENCIA
    (DIAS LICENCIA limitado a 30).

Uso por línea de comandos (sin ingreso manual):
    python imp_sin_lic.py "JULIO 2026.xlsx"
  Busca los meses anteriores, Equivalencias, parámetros y empleados en la
  misma carpeta del archivo.
"""

import io
import os
import re
import sys
import glob

import openpyxl
from openpyxl.styles import Font, PatternFill

COL_DIAS_LIC   = 'DIAS LICENCIA'
COL_NUEVA      = 'IMP SIN LIC'
COL_SIS        = 'IMP SL SIS'
NO_ENCONTRADO  = 'Imp no encontrado'
FILA_ENCABEZADO = 8          # en el archivo mensual de TW

# Haberes afectos que TeamWork resta de la base imponible
HABERES_NEGATIVOS = {
    'DESCTO. HORAS ATRASO',
    'DESCTO. HORAS ATRASO PT',
    'HORAS NO TRABAJADAS $',
    'DESC. PAGO EN EXC. IMPONIBLE',
}

MESES_ES = {
    'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
    'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11,
    'DICIEMBRE': 12,
}
NOMBRE_MES = {v: k for k, v in MESES_ES.items()}

INFORME_HEADERS = ['RUT', 'FICHA', 'CONTRATO', 'DIAS LICENCIA', 'ULT IMP SIN LIC',
                   'IMP IMP SIS']


# ─────────────────────────────────────────────────────────────────
#  Auxiliares
# ─────────────────────────────────────────────────────────────────

def n(v):
    if v is None:
        return 0.0
    try:
        return float(str(v).replace(',', '.').strip())
    except Exception:
        return 0.0


def normalizar_rut(raw):
    """'099.999.999-9' -> '99999999-9'"""
    if not raw:
        return ''
    s = str(raw).strip().upper().replace('.', '').replace(' ', '')
    if '-' in s:
        cuerpo, dv = s.rsplit('-', 1)
        return f"{cuerpo.lstrip('0')}-{dv}"
    return s.lstrip('0')


def parsear_periodo(texto):
    """'Mes a procesar: JULIO 2026' -> '2026-07'"""
    up = str(texto or '').upper()
    for nombre, num in MESES_ES.items():
        if nombre in up:
            m = re.search(r'(\d{4})', up)
            if m:
                return f"{m.group(1)}-{num:02d}"
    return None


def nombre_periodo(periodo):
    """'2026-07' -> 'JULIO 2026'"""
    y, m = periodo.split('-')
    return f"{NOMBRE_MES[int(m)]} {y}"


def _wb(fuente, read_only=True):
    if isinstance(fuente, (bytes, bytearray)):
        fuente = io.BytesIO(fuente)
    return openpyxl.load_workbook(fuente, read_only=read_only, data_only=True)


def _hoja_tw(wb):
    """Hoja con 'Mes a procesar' en la fila 2 o FICHA en el encabezado."""
    for ws in wb.worksheets:
        cab = list(ws.iter_rows(min_row=1, max_row=FILA_ENCABEZADO, values_only=True))
        fila2 = str(cab[1][0]) if len(cab) > 1 and cab[1] else ''
        if parsear_periodo(fila2) or (len(cab) >= FILA_ENCABEZADO and 'FICHA' in
                                      [str(c).strip() for c in cab[-1] if c]):
            return ws
    return wb.active


# ─────────────────────────────────────────────────────────────────
#  Carga de referencias
# ─────────────────────────────────────────────────────────────────

def cargar_haberes_afectos(equiv_fuente):
    """Nombres de columnas TW clasificadas como 'Haber afecto'."""
    wb = _wb(equiv_fuente)
    ws = wb.worksheets[0]
    cols = set()
    for row in ws.iter_rows(values_only=True):
        if not row or not row[0] or row[0] == 'Nombre columna':
            continue
        id_c = str(row[1] or '').strip()
        tipo = str(row[2] or '').strip()
        if tipo == 'Haber afecto' and id_c != 'No aplica concepto':
            cols.add(str(row[0]).strip())
    wb.close()
    return cols


def cargar_topes(params_fuente):
    """{'2026-07': topeImp_pesos_afp, ...}"""
    wb = _wb(params_fuente)
    rows = wb.active.iter_rows(values_only=True)
    hdr = [str(h).strip() if h else '' for h in next(rows)]
    i_tope = hdr.index('topeImp_pesos_afp')
    topes = {}
    for r in rows:
        if r and r[0]:
            topes[str(r[0]).strip()] = n(r[i_tope])
    wb.close()
    return topes


def cargar_contratos(emp_fuente):
    """{FICHA ('Nombre del contrato'): {'contrato': 'Contrato', 'sueldo_base': 'Sueldo Base'}}"""
    wb = _wb(emp_fuente)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    next(it, None)                                   # fila 1: título
    hdr = [str(h).strip() if h else '' for h in (next(it, None) or [])]
    i_fic = next(hdr.index(c) for c in ('Nombre del contrato', 'Nombre contr.',
                                        'Nombre contrato') if c in hdr)
    i_con = next(hdr.index(c) for c in ('Contrato', 'N° Contrato', 'Nº Contrato',
                                        'N° contrato') if c in hdr)
    i_sb = next((hdr.index(c) for c in ('Sueldo Base', 'Base contrato') if c in hdr), None)
    contratos = {}
    for r in it:
        if not r or r[i_fic] is None:
            continue
        c = r[i_con]
        if isinstance(c, float) and c.is_integer():
            c = int(c)
        sb = n(r[i_sb]) if i_sb is not None and i_sb < len(r) else 0.0
        contratos[str(r[i_fic]).strip()] = {'contrato': c, 'sueldo_base': sb}
    wb.close()
    return contratos


def leer_mes(fuente, haberes_afectos):
    """Lee un archivo mensual de TW.
    Retorna {'periodo': 'AAAA-MM', 'fichas': {FICHA: {...}}}."""
    wb = _wb(fuente)
    ws = _hoja_tw(wb)
    it = ws.iter_rows(values_only=True)
    cab = [next(it, ()) for _ in range(FILA_ENCABEZADO)]
    periodo = parsear_periodo(cab[1][0] if cab[1] else '')
    hdr = [str(h).strip() if h is not None else '' for h in cab[-1]]
    i_fic, i_rut, i_lic = hdr.index('FICHA'), hdr.index('RUT'), hdr.index(COL_DIAS_LIC)
    afectos = [(i, c) for i, c in enumerate(hdr) if c in haberes_afectos]

    fichas = {}
    for r in it:
        if not r or i_fic >= len(r) or r[i_fic] in (None, ''):
            continue
        suma = 0.0
        for i, c in afectos:
            v = n(r[i]) if i < len(r) else 0.0
            suma += -v if c in HABERES_NEGATIVOS else v
        fichas[str(r[i_fic]).strip()] = {
            'rut': normalizar_rut(r[i_rut]),
            'dias_lic': n(r[i_lic]),
            'suma_afectos': suma,
        }
    wb.close()
    return {'periodo': periodo, 'fichas': fichas}


# ─────────────────────────────────────────────────────────────────
#  Cálculo
# ─────────────────────────────────────────────────────────────────

def calcular(mes_actual, meses_previos, topes, contratos):
    """mes_actual / meses_previos: resultados de leer_mes().
    Retorna (resultados, advertencias). Un resultado por FICHA con licencia:
      {rut, ficha, contrato, dias_lic, imp (float|None), mes_origen, suma_afectos, tope}
    """
    advertencias = []
    previos = sorted([m for m in meses_previos if m['periodo'] and
                      m['periodo'] < mes_actual['periodo']],
                     key=lambda m: m['periodo'], reverse=True)

    resultados = []
    for ficha, d in mes_actual['fichas'].items():
        if d['dias_lic'] <= 0:
            continue
        emp = contratos.get(ficha) or {}
        res = {'rut': d['rut'], 'ficha': ficha, 'contrato': emp.get('contrato', ''),
               'sueldo_base': emp.get('sueldo_base', 0.0),
               'dias_lic': d['dias_lic'], 'imp': None, 'mes_origen': None,
               'suma_afectos': None, 'tope': None}
        if ficha not in contratos:
            advertencias.append(f"FICHA {ficha}: no encontrada en el listado de empleados")
        for m in previos:
            x = m['fichas'].get(ficha)
            if x is None or x['dias_lic'] != 0:
                continue
            tope = topes.get(m['periodo'], 0)
            if not tope:
                advertencias.append(f"FICHA {ficha}: sin topeImp_pesos_afp para "
                                    f"{m['periodo']}, se usa la suma sin tope")
            suma = x['suma_afectos']
            res.update(imp=float(round(min(suma, tope) if tope else suma)),
                       mes_origen=m['periodo'], suma_afectos=suma, tope=tope)
            break
        resultados.append(res)
    return resultados, advertencias


MAX_DIAS_SIS = 30


def sis_final(r, imp):
    """IMP SL SIS / IMP IMP SIS de un resultado.
    Con imponible: (imponible / 30) * min(DIAS LICENCIA, 30).
    Con 'Imp no encontrado': (Sueldo Base del listado de empleados / 30)
    * min(DIAS LICENCIA, 30). Sin sueldo base: 'Imp no encontrado'."""
    v = imp_imp_sis(imp, r['dias_lic'])
    if v is None and n(r.get('sueldo_base')) > 0:
        v = imp_imp_sis(n(r['sueldo_base']), r['dias_lic'])
    return NO_ENCONTRADO if v is None else v


def imp_imp_sis(imp, dias_lic):
    """(imp / 30) * DIAS LICENCIA, con DIAS LICENCIA limitado a 30."""
    if not isinstance(imp, (int, float)):
        return None
    return round(imp / 30 * min(dias_lic, MAX_DIAS_SIS))


# ─────────────────────────────────────────────────────────────────
#  Archivos de salida
# ─────────────────────────────────────────────────────────────────

def generar_entrada_con_columna(fuente, resultados, imp_manual=None):
    """Copia del archivo mensual con dos columnas nuevas inmediatamente a la
    derecha de DIAS LICENCIA:
      IMP SIN LIC = imponible del último mes sin licencia (0 si no hay licencia)
      IMP SL SIS  = (IMP SIN LIC / 30) * min(DIAS LICENCIA, 30)
    Las fichas con licencia sin imponible quedan con 'Imp no encontrado' en
    IMP SIN LIC, e IMP SL SIS se calcula con el Sueldo Base del listado de
    empleados: (Sueldo Base / 30) * min(DIAS LICENCIA, 30)."""
    imp_manual = imp_manual or {}
    por_ficha = {r['ficha']: r for r in resultados}

    wb = _wb(fuente, read_only=False)
    ws = _hoja_tw(wb)
    hdr = [str(c.value).strip() if c.value is not None else ''
           for c in ws[FILA_ENCABEZADO]]
    col_lic = hdr.index(COL_DIAS_LIC) + 1          # 1-based
    col_fic = hdr.index('FICHA') + 1
    col_imp, col_sis = col_lic + 1, col_lic + 2

    ws.insert_cols(col_imp, amount=2)
    # insertar columnas no corre las celdas combinadas que están a la derecha
    for rango in list(ws.merged_cells.ranges):
        if rango.min_col >= col_imp:
            rango.shift(col_shift=2)

    ref = ws.cell(FILA_ENCABEZADO, col_lic)
    letra = openpyxl.utils.get_column_letter
    ancho = ws.column_dimensions[letra(col_lic)].width
    for col, titulo in ((col_imp, COL_NUEVA), (col_sis, COL_SIS)):
        enc = ws.cell(FILA_ENCABEZADO, col, titulo)
        if ref.has_style:
            enc._style = ref._style
        ws.column_dimensions[letra(col)].width = max(ancho or 0, 14)

    for r in range(FILA_ENCABEZADO + 1, ws.max_row + 1):
        ficha = ws.cell(r, col_fic).value
        if ficha in (None, ''):
            continue
        res = por_ficha.get(str(ficha).strip())
        if res is None:
            imp, sis = 0, 0
        else:
            imp = imp_final(res, imp_manual)
            sis = sis_final(res, imp)
        for col, v in ((col_imp, imp), (col_sis, sis)):
            c = ws.cell(r, col, v)
            if isinstance(v, (int, float)):
                c.number_format = '#,##0'

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generar_informe(resultados, imp_manual=None):
    """Informe ImpSinLic. imp_manual: {FICHA: monto} ingresado por el usuario."""
    imp_manual = imp_manual or {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'ImpSinLic'
    ws.append(INFORME_HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1E5591')
    for r in resultados:
        imp = imp_final(r, imp_manual)
        sis = sis_final(r, imp)
        dias = int(r['dias_lic']) if float(r['dias_lic']).is_integer() else r['dias_lic']
        ws.append([r['rut'], r['ficha'], r['contrato'], dias, imp, sis])
    for fila in ws.iter_rows(min_row=2, min_col=5, max_col=6):
        for c in fila:
            if isinstance(c.value, (int, float)):
                c.number_format = '#,##0'
    for col, w in zip('ABCDEF', (14, 14, 11, 14, 18, 16)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def imp_final(r, imp_manual=None):
    """Imponible calculado, o el manual, o 'Imp no encontrado'."""
    if r['imp'] is not None:
        return int(r['imp'])
    m = (imp_manual or {}).get(r['ficha'])
    if m not in (None, '') and n(m) > 0:
        return int(round(n(m)))
    return NO_ENCONTRADO


# ─────────────────────────────────────────────────────────────────
#  Reconocimiento automático de los archivos cargados
# ─────────────────────────────────────────────────────────────────

TIPOS_POR_NOMBRE = [            # (tipo, fragmentos del nombre en minúsculas)
    ('salida', ('imp sin lic', 'impsinlic', 'salida_rex')),   # resultados previos: se ignoran
    ('equiv',  ('equivalencia',)),
    ('params', ('parametro',)),
    ('cot',    ('cot_afp', 'cot afp', 'cotafp')),
    ('asig',   ('asig inst', 'asig_inst', 'asiginst')),
    ('emp',    ('emplead',)),
]


def periodo_de_archivo(fuente):
    """Período ('AAAA-MM') de un archivo mensual de TW, o None si no lo es."""
    try:
        wb = _wb(fuente)
        for ws in wb.worksheets:
            cab = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
            if len(cab) > 1 and cab[1]:
                p = parsear_periodo(cab[1][0]) if 'MES A PROCESAR' in \
                    str(cab[1][0] or '').upper() else None
                if p:
                    wb.close()
                    return p
        wb.close()
    except Exception:
        pass
    return None


def tipo_por_nombre(nombre):
    base = os.path.basename(nombre).lower()
    for tipo, frags in TIPOS_POR_NOMBRE:
        if any(f in base for f in frags):
            return tipo
    return None


def nombre_informe(periodo):
    return f"ImpSinLic {nombre_periodo(periodo)}.xlsx"


def nombre_entrada(periodo):
    return f"{nombre_periodo(periodo)} IMP SIN LIC.xlsx"


# ─────────────────────────────────────────────────────────────────
#  Línea de comandos
# ─────────────────────────────────────────────────────────────────

def _buscar(carpeta, *patrones):
    for p in patrones:
        r = sorted(glob.glob(os.path.join(carpeta, p)))
        r = [x for x in r if not os.path.basename(x).startswith('~$')]
        if r:
            return r[0]
    raise FileNotFoundError(f"No se encontró {patrones[0]} en {carpeta}")


def main(ruta_mes):
    carpeta = os.path.dirname(os.path.abspath(ruta_mes))
    print("Cargando referencias…")
    afectos   = cargar_haberes_afectos(_buscar(carpeta, 'Equivalencias Tw.xlsx'))
    topes     = cargar_topes(_buscar(carpeta, 'parametrosMe*suales.xlsx'))
    contratos = cargar_contratos(_buscar(carpeta, 'EmpleadosTW.xlsx', 'empleadostw.xlsx'))

    print(f"Leyendo {os.path.basename(ruta_mes)}…")
    actual = leer_mes(ruta_mes, afectos)

    previos = []
    y, m = map(int, actual['periodo'].split('-'))
    while True:                                   # mes anterior, el anterior…
        m -= 1
        if m == 0:
            y, m = y - 1, 12
        ruta = os.path.join(carpeta, f"{NOMBRE_MES[m]} {y}.xlsx")
        if not os.path.exists(ruta):
            break
        print(f"Leyendo {os.path.basename(ruta)}…")
        previos.append(leer_mes(ruta, afectos))

    resultados, adv = calcular(actual, previos, topes, contratos)
    for a in adv:
        print("  ⚠", a)

    salida_ent = os.path.join(carpeta, nombre_entrada(actual['periodo']))
    salida_inf = os.path.join(carpeta, nombre_informe(actual['periodo']))
    with open(salida_inf, 'wb') as f:
        f.write(generar_informe(resultados))
    print("Generando copia con IMP SIN LIC / IMP SL SIS…")
    with open(salida_ent, 'wb') as f:
        f.write(generar_entrada_con_columna(ruta_mes, resultados))

    nf = sum(1 for r in resultados if r['imp'] is None)
    print(f"Con licencia: {len(resultados)} | encontrados: {len(resultados) - nf} "
          f"| no encontrados: {nf}")
    print(f"✓ {salida_inf}\n✓ {salida_ent}")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1])
