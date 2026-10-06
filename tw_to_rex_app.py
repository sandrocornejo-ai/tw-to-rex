#!/usr/bin/env python3
"""
tw_to_rex_app.py — Aplicación Streamlit
Transforma liquidaciones de TeamWork al formato de importación Rex+.

Uso:
    streamlit run tw_to_rex_app.py
"""

import io
import math
import re
import threading
import time
import openpyxl
import streamlit as st

import excel_rapido as er
import imp_sin_lic as isl

# ─────────────────────────────────────────────────────────────────
#  CONSTANTES / REGLAS DE NEGOCIO
# ─────────────────────────────────────────────────────────────────

HABERES_NEGATIVOS = {
    'DESCTO. HORAS ATRASO',
    'DESCTO. HORAS ATRASO PT',
    'HORAS NO TRABAJADAS $',
    'DESC. PAGO EN EXC. IMPONIBLE',
}

EXCLUIR_DE_EXENTOS = set()   # SOBREGIRO (compensaSobre) sí entra en rentas no gravadas

SIEMPRE_GENERAR = {'impuesto'}   # cesEmpleado solo se genera si su monto es distinto de cero

MESES_ES = {
    'ENERO': '01', 'FEBRERO': '02', 'MARZO': '03', 'ABRIL': '04',
    'MAYO': '05', 'JUNIO': '06', 'JULIO': '07', 'AGOSTO': '08',
    'SEPTIEMBRE': '09', 'OCTUBRE': '10', 'NOVIEMBRE': '11', 'DICIEMBRE': '12',
}

CONCEPTOS_AFP_INST = {
    'afp', 'aporteAFPemp', 'reliquidaAporteAFP', 'cesEmpleado',
    'reliquidaCesSol', 'cesAporteSol', 'reliquidaCesCi', 'cesAporteCi',
    'reliquidaCesEmpl', 'reliquidaAfp', 'afpAhor', 'reliquidaTrabPesa',
    'reliquidaTrabEmpl', 'trabajoPesaEmpl', 'trabajoPesa',
}

CONCEPTOS_CCAF = {
    'reliquidaCcaf', 'cajaComp', 'cajaSegu', 'cajaOtro', 'cajaVida',
    'cajaLeas', 'cajaDent', 'cajaCred', 'ccafReliquida', 'cajaAhor',
}

CONCEPTOS_MUTUAL = {'mutual', 'reliquidaMutual'}

CONCEPTOS_SEGURIDADSOCIAL = {'aporteFAPPBAC', 'aporteFAPPCEV', 'reliquidaAporteBAC'}

ISAPRE_MAPPING = {
    'BANMEDICA': 'banmedica',
    'COLMENA GOLDEN CROSS': 'colmena',
    'CONSALUD': 'consalud',
    'CRUZ BLANCA': 'cruzblanca',
    'ESENCIAL': 'esencial',
    'ISALUD': 'isalud',
    'NUEVA MASVIDA': 'nuevamasvida',
    'VIDA TRES': 'vidatres',
}

CONCEPTOS_BASE_AFP = {
    'afp', 'afpAhor', 'reliquidaAfp', 'aporteAFPemp', 'reliquidaAporteAFP',
    'aporteFAPPBAC', 'aporteFAPPCEV', 'mutual', 'reliquidaMutual',
    'sis', 'trabajoPesa', 'reliquidaTrabPesa', 'trabajoPesaEmpl', 'reliquidaTrabEmpl',
    'isapre', 'reliquidaIsapre',
}
CONCEPTOS_BASE_CES = {
    'cesEmpleado', 'cesAporteSol', 'cesAporteCi',
    'reliquidaCesSol', 'reliquidaCesCi', 'reliquidaCesEmpl', 'solidarioremu',
}
CONCEPTOS_BASE_IMP = {
    'impuesto', 'impuestoAgricola', 'reliquidaImpuesto', 'imptoindMi',
}
CONCEPTOS_BASE_TOT = {'totalesEmpl'}

# ── Afecto según licencia (licenciaDias) ─────────────────────────────
# Sin licencia: grupo AFP = min(suma haberes afectos, topeImp_pesos_afp)
#               grupo CES = min(suma haberes afectos, topeCes_pesos)
# Con licencia: ver calcular_afectos_licencia()
AFECTO_GRUPO_AFP_SIN_LIC = {
    'afp', 'isapre', 'cajaComp', 'mutual', 'sis',
    'aporteAFPemp', 'aporteFAPPCEV', 'aporteFAPPBAC',
}
AFECTO_GRUPO_AFP_CON_LIC = {          # suma o (topeAFP - IMP SL SIS)
    'afp', 'isapre', 'cajaComp', 'mutual', 'aporteFAPPBAC',
    'aporteAFPemp',
}
AFECTO_AFP_MAS_SIS = {'sis', 'aporteFAPPCEV'}     # afecto afp + IMP SL SIS
AFECTO_GRUPO_CES = {'cesAporteCi', 'cesAporteSol'}
AFECTO_CES_EMPLEADO = {'cesEmpleado'}            # siempre min(suma afectos, topeCes_pesos)

# Con licencia, el afecto de estos conceptos se deduce del monto de TeamWork:
#   sis, aporteFAPPCEV -> monto / (Cotización de jubilación / 100)
#   cesAporteSol       -> monto / 0,03
AFECTO_DESDE_MONTO_COTIZ = {'sis', 'aporteFAPPCEV'}
AFECTO_DESDE_MONTO_TASA = {'cesAporteSol': 0.03}


def afecto_desde_monto(id_concepto, monto, cotiz):
    """Afecto con licencia deducido del monto; None si no aplica."""
    if id_concepto in AFECTO_DESDE_MONTO_COTIZ:
        tasa = n(cotiz) / 100
    elif id_concepto in AFECTO_DESDE_MONTO_TASA:
        tasa = AFECTO_DESDE_MONTO_TASA[id_concepto]
    else:
        return None
    return round(monto / tasa) if tasa > 0 else None

# ── Parcial7 / Parcial8 (solo con licencia; 0 en cualquier otro caso) ─────
# Parcial7 = IMP SIN LIC (columna agregada en la Etapa 1)
# Parcial8: isapre       -> min(suma haberes afectos, topeImp_pesos_afp)
#           cesAporteSol -> min(suma haberes afectos, topeCes_pesos)
PARCIAL7_CON_LIC = {'mutual', 'isapre', 'cesAporteSol', 'aporteAFPemp', 'aporteFAPPCEV', 'sis'}
PARCIAL8_TOPE_AFP = {'isapre'}
PARCIAL8_TOPE_CES = {'cesAporteSol'}


def calcular_parciales(id_concepto, dias_lic, imp_sin_lic, suma_afectos, tope_afp, tope_ces):
    """Retorna (Parcial7, Parcial8) del concepto."""
    if dias_lic <= 0:
        return 0, 0
    suma = max(suma_afectos, 0)
    tope = lambda v, t: min(v, t) if t > 0 else v
    p7 = imp_sin_lic if id_concepto in PARCIAL7_CON_LIC else 0
    if id_concepto in PARCIAL8_TOPE_AFP:
        p8 = tope(suma, tope_afp)
    elif id_concepto in PARCIAL8_TOPE_CES:
        p8 = tope(suma, tope_ces)
    else:
        p8 = 0
    return p7, p8


DESC_LEGAL_MANUALES = {
    'AFP', 'FONASA', 'ISAPRE', 'IMPUESTO UNICO',
    'IMPUESTO UNICO DOBLE CONTRATO', 'SEG SES TRAB',
    'APV', 'APV REG (A)', 'OTROS APV', 'LIQUIDO', 'SUELDO DEL MES',
}

OUTPUT_HEADERS = [
    'Fecha de proceso', 'Id empleado', 'Número de contrato', 'Id del concepto',
    'Monto del concepto', 'Afecto', 'Id de institución', 'Cotización de jubilación',
    'Días de licencias', 'Días trabajados', 'Fecha de aplicación', 'Empresa',
    'Total de rebajas por LLSS', 'Rentas no gravadas', 'Rebaja por zona extrema',
    'Jornada', 'Días de vacaciones', 'Monto Init', 'Fase', 'Parcial7', 'Parcial8',
]


# ─────────────────────────────────────────────────────────────────
#  FUNCIONES AUXILIARES
# ─────────────────────────────────────────────────────────────────

def n(v):
    if v is None:
        return 0.0
    try:
        return float(str(v).replace(',', '.').strip())
    except Exception:
        return 0.0


def formatear_rut(raw):
    if not raw:
        return ''
    s = str(raw).strip().replace('.', '')
    partes = s.split('-')
    if len(partes) == 2:
        return f"{partes[0].lstrip('0')}-{partes[1]}"
    return s


def parsear_periodo(texto):
    upper = texto.upper() if texto else ''
    for nombre, num in MESES_ES.items():
        if nombre in upper:
            m = re.search(r'(\d{4})', texto)
            if m:
                return f"{m.group(1)}-{num}"
    return None


def desde_agosto_2026(periodo):
    try:
        y, mo = map(int, periodo.split('-'))
        return (y > 2026) or (y == 2026 and mo >= 8)
    except Exception:
        return False


def calcular_afectos_licencia(dias_lic, suma_afectos, tope_afp, tope_ces, imp_sl_sis):
    """Afecto de los conceptos previsionales según haya o no licencia.

    Siempre (con o sin licencia):
      cesEmpleado                   -> min(suma afectos, topeCes_pesos)
    licenciaDias = 0:
      afp, isapre, cajaComp, mutual, sis, aporteAFPemp,
      aporteFAPPCEV, aporteFAPPBAC  -> min(suma afectos, topeImp_pesos_afp)
      cesAporteCi, cesAporteSol     -> min(suma afectos, topeCes_pesos)
    licenciaDias > 0:
      afp, isapre, cajaComp, mutual, aporteFAPPBAC, aporteAFPemp
          -> suma afectos si es <= topeImp_pesos_afp; si no, topeImp_pesos_afp - IMP SL SIS
      sis, aporteFAPPCEV            -> afecto afp + IMP SL SIS
      cesAporteCi, cesAporteSol     -> min(suma afectos + IMP SL SIS, topeCes_pesos)
    """
    suma = max(suma_afectos, 0)
    tope = lambda v, t: min(v, t) if t > 0 else v
    af = {}
    for c in AFECTO_CES_EMPLEADO:
        af[c] = tope(suma, tope_ces)
    if dias_lic <= 0:
        for c in AFECTO_GRUPO_AFP_SIN_LIC:
            af[c] = tope(suma, tope_afp)
        for c in AFECTO_GRUPO_CES:
            af[c] = tope(suma, tope_ces)
        return af
    if tope_afp <= 0 or suma <= tope_afp:
        af_afp = suma
    else:
        af_afp = max(tope_afp - imp_sl_sis, 0)
    for c in AFECTO_GRUPO_AFP_CON_LIC:
        af[c] = af_afp
    for c in AFECTO_AFP_MAS_SIS:
        af[c] = af_afp + imp_sl_sis
    for c in AFECTO_GRUPO_CES:
        af[c] = tope(suma + imp_sl_sis, tope_ces)
    return af


def get_afecto(id_concepto, base_afp, base_ces, base_imp, suma_afectos):
    if id_concepto in CONCEPTOS_BASE_AFP:
        return round(base_afp)
    if id_concepto in CONCEPTOS_BASE_CES:
        return round(base_ces)
    if id_concepto in CONCEPTOS_BASE_IMP:
        return round(base_imp)
    if id_concepto in CONCEPTOS_BASE_TOT:
        return round(suma_afectos)
    return 0


# ─────────────────────────────────────────────────────────────────
#  FUNCIONES DE CARGA DE ARCHIVOS (desde bytes en memoria)
# ─────────────────────────────────────────────────────────────────

def cargar_equivalencias(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    equiv = {}
    for row in ws.iter_rows(values_only=True):
        if not row[0] or row[0] == 'Nombre columna':
            continue
        col_name    = str(row[0]).strip()
        id_concepto = str(row[1]).strip() if row[1] else ''
        tipo        = str(row[2]).strip() if row[2] else ''
        obs         = str(row[3]).strip() if row[3] else ''
        equiv[col_name] = {'id_concepto': id_concepto, 'tipo': tipo, 'obs': obs}
    wb.close()
    return equiv


# Nombres alternativos de columnas en el maestro de empleados
# (el "Listado de Empleados" de Rex+ usa nombres distintos)
ALIAS_EMPLEADOS = {
    'Nombre del contrato': ['Nombre contr.', 'Nombre contrato'],
    'N° Contrato':         ['Contrato', 'Nº Contrato', 'N° contrato'],
    'Base contrato':       ['Sueldo Base'],
    'horasSema':           ['Horas Semanales'],
    'Id empresa':          ['Empresa'],
    'Id Afp':              ['AFP'],
    'Id Salud':            ['Isapre'],
    'Id Mutual':           ['Mutual'],
    '% Mutual':            ['Tasa Mutual', 'Tasa mutual'],
    'Id CCAF':             ['CCAF'],
}


def cargar_empleados(file_bytes, avance=None):
    """Lee el maestro de empleados fila a fila (streaming) y guarda solo las
    columnas que usa la app. Evita cargar el Excel completo en memoria, que
    con listados grandes (~180 mil filas x 189 columnas) colgaba la app."""
    _, total, it = er.leer_hoja(file_bytes)
    next(it, None)                       # fila 1: título
    headers = [str(h).strip() if h else '' for h in (next(it, None) or [])]
    pos = {}
    for i, h in enumerate(headers):      # si hay duplicados, gana el último
        if h:
            pos[h] = i

    # columna canónica -> índice de la primera variante que exista
    idx = {}
    for canon, alts in ALIAS_EMPLEADOS.items():
        for nombre in [canon] + alts:
            if nombre in pos:
                idx[canon] = pos[nombre]
                break

    empleados = {}
    for k, row in enumerate(it):
        if avance and total and k % 2000 == 0:
            avance(min(k / total, 1.0))
        if not row or not row[0]:
            continue
        d = {c: (row[i] if i < len(row) else None) for c, i in idx.items()}
        nc = d.get('N° Contrato')
        if isinstance(nc, float) and nc.is_integer():
            d['N° Contrato'] = int(nc)
        key = str(d.get('Nombre del contrato') or '').strip()
        if key:
            empleados[key] = d
    return empleados


def cargar_params(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdrs = [str(h).strip() if h else '' for h in rows[0]]
    params = {}
    for row in rows[1:]:
        if not row[0]:
            continue
        key = str(row[0]).strip()
        params[key] = {hdrs[i]: row[i] for i in range(len(hdrs))}
    return params


def cargar_cot_afp(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    cot_afp = {}
    for row in ws.iter_rows(values_only=True):
        if not row[0] or row[0] == 'id_afp_hist':
            continue
        cot_afp[str(row[0]).strip()] = n(row[4])
    wb.close()
    return cot_afp


def cargar_asig(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    asig = {}
    for row in rows[1:]:
        if not row[0]:
            continue
        asig[str(row[0]).strip()] = {
            'hasta_jul': str(row[1]).strip() if row[1] else '',
            'desde_ago': str(row[2]).strip() if row[2] else '',
        }
    return asig


def cargar_tw(file_bytes, avance=None):
    """Lee la hoja de liquidaciones de TW. No depende de la hoja activa:
    busca la hoja cuya fila 2 contiene el período ("Mes a procesar: ...")
    o cuyo encabezado incluye FICHA; si no encuentra ninguna, usa la activa."""
    def es_tw(cab):
        fila2 = str(cab[1][0]) if len(cab) > 1 and cab[1] and cab[1][0] is not None else ''
        tiene_ficha = any(r and 'FICHA' in [str(c).strip() for c in r if c] for r in cab)
        return bool(parsear_periodo(fila2) or tiene_ficha)

    _, total, it = er.leer_hoja(file_bytes, es_tw)
    rows = []
    for k, r in enumerate(it):
        if avance and total and k % 200 == 0:
            avance(min(k / total, 1.0))
        rows.append(r)
    return rows


# ─────────────────────────────────────────────────────────────────
#  LÓGICA PRINCIPAL DE TRANSFORMACIÓN
# ─────────────────────────────────────────────────────────────────

def procesar(tw_bytes, equiv_bytes, emp_bytes, params_bytes, cot_bytes, asig_bytes=None,
             progress_callback=None, log_callback=None, paso_callback=None):
    """
    Ejecuta la transformación completa y retorna (output_rows, advertencias, stats).
    paso_callback(clave): avisa el inicio de cada paso
        ('equiv', 'emp', 'refs', 'tw', 'proc')
    progress_callback(fraccion): avance del paso en curso (0.0 a 1.0)
    log_callback(texto): emite mensajes de estado
    """

    def log(msg):
        if log_callback:
            log_callback(msg)

    def paso(clave):
        if paso_callback:
            paso_callback(clave)

    # ── Cargar referencias ──────────────────────────────────────────
    paso('equiv')
    log("Cargando Equivalencias Tw.xlsx…")
    equiv = cargar_equivalencias(equiv_bytes)
    log(f"  ✓ {len(equiv)} equivalencias")

    paso('emp')
    log("Cargando empleadostw.xlsx…")
    empleados = cargar_empleados(emp_bytes, progress_callback)
    log(f"  ✓ {len(empleados)} empleados")

    paso('refs')
    log("Cargando parametrosMesuales.xlsx…")
    params = cargar_params(params_bytes)
    log(f"  ✓ {len(params)} períodos")

    log("Cargando cot_afp_hist.xlsx…")
    cot_afp = cargar_cot_afp(cot_bytes)
    log(f"  ✓ {len(cot_afp)} cotizaciones AFP")

    if asig_bytes:
        log("Cargando Asig Inst LD.xlsx…")
        asig = cargar_asig(asig_bytes)
        log(f"  ✓ {len(asig)} reglas de institución")
    else:
        asig = {}
        log("  ℹ Asig Inst LD.xlsx no cargado (opcional)")

    # ── Funciones que dependen de las tablas cargadas ───────────────
    def get_institucion(concepto, emp, row, desde_ago):
        if concepto in CONCEPTOS_AFP_INST:
            nombre_afp = str(safe_val(row, IDX_NOMBRE_AFP) or '').strip().lower()
            return nombre_afp if nombre_afp else 'afp'
        if concepto in CONCEPTOS_CCAF:
            return emp.get('Id CCAF') or ''
        if concepto in CONCEPTOS_MUTUAL:
            return emp.get('Id Mutual') or ''
        if concepto == 'isapre':
            inst = str(safe_val(row, IDX_INST_SALUD) or '').strip().upper()
            if not inst:
                return 'fonasa'
            return ISAPRE_MAPPING.get(inst, 'Falta Id Salud')
        if concepto == 'impuesto':
            return 'Impuesto'
        if concepto == 'sis':
            if desde_ago:
                return 'seguridadsocial'
            return str(safe_val(row, IDX_NOMBRE_AFP) or '').strip().lower()
        if concepto in CONCEPTOS_SEGURIDADSOCIAL:
            return 'seguridadsocial'
        # fallback: usar tabla Asig Inst LD
        regla = asig.get(concepto, {})
        col = regla.get('desde_ago' if desde_ago else 'hasta_jul', '')
        if col == 'Id afp':
            return emp.get('Id Afp') or ''
        if col == 'Id Salud':
            return emp.get('Id Salud') or ''
        if col == 'Id Mutual':
            return emp.get('Id Mutual') or ''
        if col == 'Id CCAF':
            return emp.get('Id CCAF') or ''
        if col == 'seguridadsocial':
            return 'seguridadsocial'
        if col == 'Impuesto':
            return 'Impuesto'
        return None

    def get_cotizacion(concepto, emp, afecto, periodo, desde_ago,
                       sis_tasa, aporte_afp, seg_vida, aporte_bac):
        id_afp = emp.get('Id Afp', '') or ''
        if concepto == 'afp':
            key = f"{periodo}{id_afp}"
            cot = cot_afp.get(key, 0)
            if not cot and '-' in periodo:
                y, mo = int(periodo[:4]), int(periodo[5:])
                mo -= 1
                if mo == 0:
                    mo, y = 12, y - 1
                key_prev = f"{y:04d}-{mo:02d}{id_afp}"
                cot = cot_afp.get(key_prev, 0)
            return round(cot * 100, 4) if cot else None
        if concepto == 'isapre':
            return afecto
        if concepto == 'mutual':
            return n(emp.get('% Mutual', 0))
        if concepto == 'sis':
            return round(sis_tasa, 2)
        if concepto == 'aporteAFPemp':
            return aporte_afp
        if concepto == 'aporteFAPPCEV':
            return seg_vida
        if concepto == 'aporteFAPPBAC':
            return aporte_bac
        if concepto == 'totalesEmpl':
            return round(afecto)   # = base_afp (topada), igual al afecto de afp/isapre
        if concepto == 'cesEmpleado':
            return 0.6
        return None

    # ── Cargar tw.xlsx ──────────────────────────────────────────────
    paso('tw')
    log("Cargando tw.xlsx…")
    tw_rows_all = cargar_tw(tw_bytes, progress_callback)

    periodo = parsear_periodo(str(tw_rows_all[1][0]))
    if not periodo:
        raise ValueError("No se pudo detectar el período en la fila 2 de tw.xlsx")
    log(f"  ✓ Período detectado: {periodo}")

    tw_hdrs = [str(h).strip() if h is not None else '' for h in tw_rows_all[7]]
    tw_data  = tw_rows_all[8:]
    log(f"  ✓ {len(tw_data)} empleados en tw.xlsx")

    # ── Parámetros del período ──────────────────────────────────────
    pm = params.get(periodo, {})
    if not pm:
        log(f"  ⚠ Sin parámetros para {periodo}, usando ceros")

    tope_afp   = n(pm.get('topeImp_pesos_afp', 0))
    tope_ces   = n(pm.get('topeCes_pesos', 0))
    tope_salud = n(pm.get('topeSalud_pesos', 0))
    sis_tasa   = n(pm.get('sis', 0))
    aporte_afp = n(pm.get('Aporte AFP', 0))
    seg_vida   = n(pm.get('Seg Social Exp vida', 0))
    aporte_bac = n(pm.get('aporteFAPPBAC', 0))
    aporte_ccaf = n(pm.get('aporte_Ccaf', 0))
    desde_ago  = desde_agosto_2026(periodo)

    log(f"  Tope AFP: {tope_afp:,.0f}  |  Tope CES: {tope_ces:,.0f}  |  Tope Salud: {tope_salud:,.0f}")

    # ── Índices de columnas ─────────────────────────────────────────
    def col_idx(name):
        try:
            return tw_hdrs.index(name)
        except ValueError:
            return None

    IDX_RUT   = col_idx('RUT')
    IDX_FICHA = col_idx('FICHA')
    IDX_DIAS_LIC  = col_idx('DIAS LICENCIA')
    IDX_IMP_SL_SIS = col_idx('IMP SL SIS')        # agregada en la Etapa 1
    IDX_IMP_SIN_LIC = col_idx('IMP SIN LIC')      # agregada en la Etapa 1
    IDX_DIAS_TRAB = col_idx('DIAS TRABAJADOS')
    IDX_FONASA    = col_idx('FONASA')
    IDX_ISAPRE    = col_idx('ISAPRE')
    IDX_AFP       = col_idx('AFP')
    IDX_LIQUIDO   = col_idx('LIQUIDO')
    IDX_SUELDO     = col_idx('SUELDO DEL MES')
    IDX_INST_SALUD  = col_idx('INST SALUD')
    IDX_NOMBRE_AFP  = col_idx('NOMBRE AFP')
    IDX_VAC      = col_idx('SUELDO POR VACACIONES')
    IDX_IMP1      = col_idx('IMPUESTO UNICO')
    IDX_IMP2      = col_idx('IMPUESTO UNICO DOBLE CONTRATO')
    IDX_SEG_SES   = col_idx('SEG SES TRAB')

    APV_COLS = [i for c, i in [
        ('APV',        col_idx('APV')),
        ('APV REG (A)', col_idx('APV REG (A)')),
        ('OTROS APV',  col_idx('OTROS APV')),
    ] if i is not None]

    haber_afecto_cols    = []
    haber_exento_cols    = []
    descuento_legal_cols = []
    descuento_cols       = []
    aporte_emp_cols      = []

    for col_name, eq in equiv.items():
        ci = col_idx(col_name)
        if ci is None or eq['id_concepto'] == 'No aplica concepto':
            continue
        tipo = eq['tipo']
        if tipo == 'Haber afecto':
            haber_afecto_cols.append((col_name, ci, eq))
        elif tipo == 'Haber exento':
            haber_exento_cols.append((col_name, ci, eq))
        elif tipo == 'Descuento Legal' and col_name not in DESC_LEGAL_MANUALES:
            descuento_legal_cols.append((col_name, ci, eq))
        elif tipo == 'Descuento' and col_name not in DESC_LEGAL_MANUALES:
            descuento_cols.append((col_name, ci, eq))
        elif tipo == 'Aporte Empleador' and col_name not in DESC_LEGAL_MANUALES:
            aporte_emp_cols.append((col_name, ci, eq))

    # ── Procesar empleados ──────────────────────────────────────────
    output_rows  = [OUTPUT_HEADERS]
    n_procesados = 0
    n_filas_gen  = [0]
    advertencias = []
    total        = len(tw_data)

    def safe_val(row, idx):
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    paso('proc')
    log("Procesando empleados…")

    for fila_n, row in enumerate(tw_data, start=9):

        ficha_raw = safe_val(row, IDX_FICHA)
        if ficha_raw is None or str(ficha_raw).strip() == '':
            continue

        ficha = str(ficha_raw).strip()
        rut   = formatear_rut(safe_val(row, IDX_RUT))

        dias_lic  = n(safe_val(row, IDX_DIAS_LIC))
        dias_trab = n(safe_val(row, IDX_DIAS_TRAB))

        emp = empleados.get(ficha, {})
        if not emp:
            advertencias.append(f"Fila {fila_n}: FICHA '{ficha}' no encontrada en empleadostw")

        id_empresa    = emp.get('Id empresa', '') or ''
        horas_sem     = n(emp.get('horasSema', 42))
        base_contrato = n(emp.get('Base contrato', 0))
        n_contrato    = emp.get('N° Contrato', '') if emp else ''
        jornada       = 'P' if horas_sem < 31 else 'C'

        # Sumas de haberes
        suma_afectos = 0.0
        for col_name, ci, eq in haber_afecto_cols:
            v = n(safe_val(row, ci))
            suma_afectos += (-v if col_name in HABERES_NEGATIVOS else v)

        suma_exentos = 0.0
        for col_name, ci, eq in haber_exento_cols:
            v = n(safe_val(row, ci))
            if col_name not in EXCLUIR_DE_EXENTOS:
                suma_exentos += v

        base_afp = max(min(suma_afectos, tope_afp) if tope_afp > 0 else suma_afectos, 0)
        base_ces = max(min(suma_afectos, tope_ces) if tope_ces > 0 else suma_afectos, 0)
        suma_afectos_pos = max(suma_afectos, 0)

        # Afectos previsionales según licencia (IMP SL SIS viene de la Etapa 1)
        imp_sl_sis = n(safe_val(row, IDX_IMP_SL_SIS)) if dias_lic > 0 else 0.0
        # IMP SIN LIC = "Imp no encontrado" (texto): se reemplaza por
        # (IMP SL SIS / min(DIAS LICENCIA, 30)) * 30
        imp_sin_lic = 0.0
        if dias_lic > 0:
            v_isl = safe_val(row, IDX_IMP_SIN_LIC)
            if isinstance(v_isl, str) and v_isl.strip() == isl.NO_ENCONTRADO:
                imp_sin_lic = imp_sl_sis / min(dias_lic, 30) * 30
            else:
                imp_sin_lic = n(v_isl)
        afectos_lic = calcular_afectos_licencia(dias_lic, suma_afectos, tope_afp,
                                                tope_ces, imp_sl_sis)

        v_afp_col    = n(safe_val(row, IDX_AFP))
        v_seg_ses    = n(safe_val(row, IDX_SEG_SES))
        v_apv_sum    = sum(n(safe_val(row, i)) for i in APV_COLS)
        v_fonasa     = n(safe_val(row, IDX_FONASA))
        v_isapre_col = n(safe_val(row, IDX_ISAPRE))
        salud_total  = v_fonasa + v_isapre_col
        rebaja_salud = min(salud_total, tope_salud) if tope_salud > 0 else salud_total

        total_rebajas_llss = v_afp_col + v_apv_sum + v_seg_ses + rebaja_salud
        base_imp = max(suma_afectos - total_rebajas_llss, 0)
        monto_isapre = v_fonasa if v_fonasa != 0 else v_isapre_col

        def fila(id_concepto, monto, afecto_val=0, id_inst=None, cotiz=None,
                 total_reb=0, rentas_ng=0, monto_init=0):
            parcial7, parcial8 = calcular_parciales(id_concepto, dias_lic, imp_sin_lic,
                                                    suma_afectos, tope_afp, tope_ces)
            output_rows.append([
                periodo,
                rut,
                n_contrato,
                id_concepto,
                round(monto) if isinstance(monto, float) else monto,
                round(afecto_val),
                id_inst,
                cotiz,
                int(dias_lic),
                int(dias_trab),
                periodo,
                id_empresa,
                round(total_reb),
                round(rentas_ng),
                0,
                jornada,
                0,
                round(monto_init),
                1,
                round(parcial7),
                round(parcial8),
            ])
            n_filas_gen[0] += 1

        # sueldoBase = SUELDO DEL MES + SUELDO POR VACACIONES
        v_sueldo      = n(safe_val(row, IDX_SUELDO))
        v_vac         = n(safe_val(row, IDX_VAC))
        v_sueldo_base = v_sueldo + v_vac
        if v_sueldo_base != 0:
            fila('sueldoBase', v_sueldo_base, monto_init=base_contrato)

        # sueldoBaseVac — solo si SUELDO POR VACACIONES > 0
        if v_vac > 0:
            fila('sueldoBaseVac', v_vac)

        # sueldoBasesinVac — valor de SUELDO DEL MES
        if v_sueldo != 0:
            fila('sueldoBasesinVac', v_sueldo)

        # Haberes afectos agrupados
        hab_afecto_agrup = {}
        for col_name, ci, eq in haber_afecto_cols:
            if col_name in ('SUELDO DEL MES', 'SUELDO POR VACACIONES'):
                continue
            id_c = eq['id_concepto']
            v = n(safe_val(row, ci))
            monto_v = -v if col_name in HABERES_NEGATIVOS else v
            hab_afecto_agrup[id_c] = hab_afecto_agrup.get(id_c, 0) + monto_v
        for id_c, monto_v in hab_afecto_agrup.items():
            if monto_v != 0:
                fila(id_c, monto_v)

        # Haberes exentos agrupados
        hab_exento_agrup = {}
        for col_name, ci, eq in haber_exento_cols:
            id_c = eq['id_concepto']
            v = n(safe_val(row, ci))
            hab_exento_agrup[id_c] = hab_exento_agrup.get(id_c, 0) + v
        for id_c, monto_v in hab_exento_agrup.items():
            if monto_v != 0:
                fila(id_c, monto_v)

        # AFP
        if v_afp_col != 0:
            inst  = get_institucion('afp', emp, row, desde_ago)
            cotiz = get_cotizacion('afp', emp, base_afp, periodo, desde_ago,
                                   sis_tasa, aporte_afp, seg_vida, aporte_bac)
            fila('afp', v_afp_col, afectos_lic['afp'], inst, cotiz)

        # isapre (siempre) — cotización = mismo valor que monto
        inst_isapre = get_institucion('isapre', emp, row, desde_ago)
        fila('isapre', monto_isapre, afectos_lic['isapre'], inst_isapre, monto_isapre)

        # cesEmpleado (solo si el monto es distinto de cero)
        if round(v_seg_ses) != 0:
            inst_ces = get_institucion('cesEmpleado', emp, row, desde_ago)
            fila('cesEmpleado', v_seg_ses, afectos_lic['cesEmpleado'], inst_ces, 0.6)

        # impuesto (siempre)
        v_imp1 = n(safe_val(row, IDX_IMP1))
        v_imp2 = n(safe_val(row, IDX_IMP2))
        fila('impuesto', v_imp1 + v_imp2, base_imp, 'Impuesto', None,
             total_reb=total_rebajas_llss, rentas_ng=suma_exentos)

        # APV
        if v_apv_sum != 0:
            fila('apvi', v_apv_sum)

        # Descuentos legales
        desc_legal_agrup = {}
        for col_name, ci, eq in descuento_legal_cols:
            id_c = eq['id_concepto']
            v = n(safe_val(row, ci))
            desc_legal_agrup[id_c] = desc_legal_agrup.get(id_c, 0) + v
        for id_c, monto_v in desc_legal_agrup.items():
            if monto_v != 0:
                afecto_v = afectos_lic.get(id_c)
                if afecto_v is None:
                    afecto_v = get_afecto(id_c, base_afp, base_ces, base_imp, suma_afectos_pos)
                inst  = get_institucion(id_c, emp, row, desde_ago)
                cotiz = get_cotizacion(id_c, emp, afecto_v, periodo, desde_ago,
                                       sis_tasa, aporte_afp, seg_vida, aporte_bac)
                if dias_lic > 0:
                    af_monto = afecto_desde_monto(id_c, monto_v, cotiz)
                    if af_monto is not None:
                        afecto_v = af_monto
                fila(id_c, monto_v, afecto_v, inst, cotiz)

        # Descuentos normales
        desc_agrup = {}
        for col_name, ci, eq in descuento_cols:
            id_c = eq['id_concepto']
            v = n(safe_val(row, ci))
            desc_agrup[id_c] = desc_agrup.get(id_c, 0) + v
        for id_c, monto_v in desc_agrup.items():
            if monto_v != 0:
                inst = get_institucion(id_c, emp, row, desde_ago)
                fila(id_c, monto_v, id_inst=inst)

        # Aportes empleador
        aporte_agrup = {}
        for col_name, ci, eq in aporte_emp_cols:
            id_c = eq['id_concepto']
            v = n(safe_val(row, ci))
            aporte_agrup[id_c] = aporte_agrup.get(id_c, 0) + v
        for id_c, monto_v in aporte_agrup.items():
            if monto_v != 0:
                afecto_v = afectos_lic.get(id_c)
                if afecto_v is None:
                    afecto_v = get_afecto(id_c, base_afp, base_ces, base_imp, suma_afectos_pos)
                inst  = get_institucion(id_c, emp, row, desde_ago)
                cotiz = get_cotizacion(id_c, emp, afecto_v, periodo, desde_ago,
                                       sis_tasa, aporte_afp, seg_vida, aporte_bac)
                if dias_lic > 0:
                    af_monto = afecto_desde_monto(id_c, monto_v, cotiz)
                    if af_monto is not None:
                        afecto_v = af_monto
                fila(id_c, monto_v, afecto_v, inst, cotiz)

        # totalesEmpl (LIQUIDO)
        v_liq = n(safe_val(row, IDX_LIQUIDO))
        cotiz_tot = get_cotizacion('totalesEmpl', emp, base_afp, periodo,
                                   desde_ago, sis_tasa, aporte_afp, seg_vida, aporte_bac)
        fila('totalesEmpl', v_liq, suma_afectos_pos, None, cotiz_tot)

        # licenciaDias
        if dias_lic > 0:
            fila('licenciaDias', int(dias_lic))

        # cajaComp — si INST SALUD del archivo de entrada está en blanco o vacío
        inst_salud_tw = str(safe_val(row, IDX_INST_SALUD) or '').strip()
        if inst_salud_tw == '' and aporte_ccaf > 0:
            monto_ccaf = suma_afectos * (aporte_ccaf / 100)
            if monto_ccaf != 0:
                id_ccaf = emp.get('Id CCAF') or ''
                fila('cajaComp', monto_ccaf, afectos_lic['cajaComp'], id_inst=id_ccaf)

        n_procesados += 1
        if progress_callback and total > 0 and n_procesados % 50 == 0:
            progress_callback(n_procesados / total)

    stats = {
        'periodo': periodo,
        'empleados': n_procesados,
        'filas': n_filas_gen[0],
        'advertencias': len(advertencias),
    }
    return output_rows, advertencias, stats


def generar_excel(output_rows, avance=None):
    """Genera el archivo Excel en memoria y retorna bytes.
    avance(fraccion): opcional. Armar las filas va de 0 a 0,4; guardar el
    archivo no se puede medir (la barra sigue avanzando sola)."""
    hdr = output_rows[0]
    c_con = hdr.index('Id del concepto')
    c_cot = hdr.index('Cotización de jubilación')
    rapido = er.escribir_filas(
        output_rows, 'Liquidaciones',
        {c_cot: lambda fila: '0.00' if fila[c_con] == 'sis' else None},
        avance)
    if rapido is not None:
        return rapido

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Liquidaciones'
    total = len(output_rows)
    for k, row in enumerate(output_rows):
        if avance and k % 2000 == 0:
            avance(0.35 * k / total)
        ws.append(row)
    # Cotización de jubilación del sis con 2 decimales
    hdr = output_rows[0]
    c_con = hdr.index('Id del concepto') + 1
    c_cot = hdr.index('Cotización de jubilación') + 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, c_con).value == 'sis':
            ws.cell(r, c_cot).number_format = '0.00'
    if avance:
        avance(0.4)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────
#  INTERFAZ STREAMLIT — proceso único en dos etapas
# ─────────────────────────────────────────────────────────────────

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Segundos estimados por paso (solo reparten la barra de avance)
_RAPIDO = not er.motores_rapidos()
T = ({'emp': 2, 'mes': 1, 'entrada': 4, 'excel': 10} if _RAPIDO else
     {'emp': 17, 'mes': 9, 'entrada': 33, 'excel': 20})


def _version_codigo():
    """Fecha de modificación de los .py de la app: si cambia el código, los
    resultados guardados en la sesión se descartan (evita descargar un
    archivo generado con la versión anterior)."""
    import os
    carpeta = os.path.dirname(os.path.abspath(__file__))
    return tuple(os.path.getmtime(os.path.join(carpeta, f))
                 for f in ('tw_to_rex_app.py', 'imp_sin_lic.py', 'excel_rapido.py')
                 if os.path.exists(os.path.join(carpeta, f)))

try:
    from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx
except ImportError:                                   # versiones antiguas de Streamlit
    add_script_run_ctx = get_script_run_ctx = None


def _mmss(seg):
    seg = int(seg)
    return f"{seg // 60}:{seg % 60:02d}"


class Avance:
    """Barra de avance de un proceso con varios pasos.

    pasos: lista de (clave, texto, segundos_estimados). El peso de cada paso en
    la barra es proporcional a su tiempo estimado.
    - paso(clave): pasa al paso indicado y deja el anterior marcado como hecho.
    - avance(fraccion): avance real (0..1) del paso en curso.
    Un hilo refresca la barra cada medio segundo con el tiempo transcurrido y,
    si el paso no informa avance (abrir o guardar un Excel), la hace avanzar
    sola de a poco, sin pasar el 95 % del paso. Así nunca parece detenida.
    """

    TOPE = 0.95

    def __init__(self, pasos, contenedor=None):
        self.pasos = pasos
        total = sum(p[2] for p in pasos) or 1
        self.ancho = [p[2] / total for p in pasos]
        self.base = [sum(self.ancho[:i]) for i in range(len(pasos))]
        self.idx = {p[0]: i for i, p in enumerate(pasos)}
        cont = contenedor or st
        self.lineas = cont.empty()
        self.barra = cont.progress(0.0, text="Iniciando…")
        self.hechos = []
        self.i = -1
        self.t0 = self.t_paso = self.t_sub = time.time()
        self.sub = 0.0
        self.mostrado = 0.0
        self.lock = threading.Lock()
        self.fin = threading.Event()
        self.hilo = None
        if add_script_run_ctx and get_script_run_ctx() is not None:
            self.hilo = threading.Thread(target=self._latido, daemon=True)
            add_script_run_ctx(self.hilo)
            self.hilo.start()

    # ── API ──
    def paso(self, clave):
        with self.lock:
            ahora = time.time()
            if self.i >= 0:
                self.hechos.append(f"✅ {self.pasos[self.i][1]} · {_mmss(ahora - self.t_paso)}")
            self.i = self.idx[clave]
            self.t_paso = self.t_sub = ahora
            self.sub = 0.0
            self._pintar(lineas=True)

    def avance(self, frac):
        with self.lock:
            frac = max(0.0, min(float(frac), 1.0))
            if frac > self.sub:
                self.sub, self.t_sub = frac, time.time()
                self._pintar()

    def cerrar(self, texto="Listo", limpiar=True):
        self.fin.set()
        if self.hilo:
            self.hilo.join(timeout=2)
        with self.lock:
            if self.i >= 0:
                self.hechos.append(f"✅ {self.pasos[self.i][1]} · {_mmss(time.time() - self.t_paso)}")
                self.i = -1
            if limpiar:
                self.barra.empty()
                self.lineas.empty()
            else:
                self.barra.progress(1.0, text=f"{texto} · {_mmss(time.time() - self.t0)}")
                self._pintar_lineas()
        return _mmss(time.time() - self.t0)

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.cerrar()
        return False

    # ── interno ──
    def _latido(self):
        while not self.fin.wait(0.5):
            with self.lock:
                if self.i >= 0:
                    self._pintar()

    def _fraccion_paso(self):
        """Avance real del paso; si lleva rato sin novedades, avanza solo hacia el 95 %."""
        esperado = max(self.pasos[self.i][2], 1)
        quieto = time.time() - self.t_sub
        extra = (self.TOPE - self.sub) * (1 - math.exp(-quieto / esperado))
        return self.sub + max(extra, 0.0) if self.sub < self.TOPE else self.sub

    def _pintar(self, lineas=False):
        if self.i < 0:
            return
        frac = self.base[self.i] + self.ancho[self.i] * self._fraccion_paso()
        self.mostrado = max(self.mostrado, min(frac, 0.999))     # nunca retrocede
        texto = (f"Paso {self.i + 1} de {len(self.pasos)} · {self.pasos[self.i][1]}… "
                 f"{self.mostrado * 100:.0f} % · ⏱ {_mmss(time.time() - self.t0)}")
        try:
            self.barra.progress(self.mostrado, text=texto)
            if lineas:
                self._pintar_lineas()
        except Exception:                                  # la página se recargó
            self.fin.set()

    def _pintar_lineas(self):
        actual = [f"⏳ {self.pasos[self.i][1]}…"] if self.i >= 0 else []
        self.lineas.markdown("  \n".join(self.hechos + actual))


ETIQUETAS = {
    'emp':    "👥 EmpleadosTW.xlsx  *(maestro de empleados)*",
    'equiv':  "🔄 Equivalencias Tw.xlsx  *(mapeo de conceptos)*",
    'params': "⚙️ parametrosMesuales.xlsx  *(topes y tasas)*",
    'cot':    "📊 cot_afp_hist.xlsx  *(cotizaciones AFP)*",
    'asig':   "🏦 Asig Inst LD.xlsx  *(opcional)*",
}
OBLIGATORIOS = ('emp', 'equiv', 'params', 'cot')


@st.cache_data(show_spinner=False)
def _periodo(nombre, datos):
    return isl.periodo_de_archivo(datos)


@st.cache_data(show_spinner=False)
def _haberes(datos):
    return sorted(isl.cargar_haberes_afectos(datos))


@st.cache_data(show_spinner=False)
def _topes(datos):
    return isl.cargar_topes(datos)


# Las funciones con caché no pueden tocar la barra de avance: Streamlit
# "repite" lo que dibujan y falla si el elemento se creó fuera de ellas.
# Por eso no reciben el callback de avance (la barra avanza sola mientras leen).
@st.cache_data(show_spinner=False)
def _contratos(datos):
    return isl.cargar_contratos(datos)


@st.cache_data(show_spinner=False)
def _mes(datos, haberes):
    return isl.leer_mes(datos, set(haberes))


def clasificar_meses(archivos):
    """Detecta el período de cada archivo mensual. Retorna (meses, ignorados)."""
    meses, ignorados = {}, []
    barra = st.progress(0.0, text="Revisando archivos…") if len(archivos) > 1 else None
    for k, f in enumerate(archivos):
        if barra:
            barra.progress(k / len(archivos),
                           text=f"Revisando archivo {k + 1} de {len(archivos)}: {f.name}")
        if f.name.startswith('~$'):
            continue
        if isl.tipo_por_nombre(f.name) == 'salida':
            ignorados.append(f"{f.name} (resultado de un proceso anterior)")
            continue
        p = _periodo(f.name, f.getvalue())
        if p and p in meses:
            ignorados.append(f"{f.name} (repite {isl.nombre_periodo(p)})")
        elif p:
            meses[p] = f
        else:
            ignorados.append(f"{f.name} (no es un archivo mensual de TeamWork)")
    if barra:
        barra.empty()
    return meses, ignorados


st.set_page_config(page_title="TeamWork → Rex+", page_icon="💼", layout="wide")
st.title("💼 TeamWork → Rex+")
st.caption("Transforma liquidaciones de TeamWork al formato de importación Rex+")
if er.motores_rapidos():
    st.warning("⚡ La app funcionará más lenta: falta instalar **" +
               ", ".join(er.motores_rapidos()) + "**. En la consola ejecuta:  "
               "`pip install python-calamine xlsxwriter`  y vuelve a iniciar la app.")

# ── Carga de archivos ───────────────────────────────────────────────
c1, c2 = st.columns(2)
with c1:
    st.subheader("📅 Meses de TeamWork")
    archivos_mes = st.file_uploader(
        "Sube juntos el mes a procesar y los meses anteriores",
        type=["xlsx"], accept_multiple_files=True, key="meses")
    meses, ignorados = clasificar_meses(archivos_mes or [])
    mes_proc = None
    if meses:
        orden = sorted(meses, reverse=True)
        mes_proc = st.selectbox("Mes a procesar", orden, index=0,
                                format_func=isl.nombre_periodo)
        ant = [p for p in orden if p < mes_proc]
        st.caption("Meses anteriores para buscar el imponible: " +
                   (", ".join(isl.nombre_periodo(p) for p in ant) or "ninguno"))
    for x in ignorados:
        st.caption(f"⚠️ Se ignora: {x}")

with c2:
    st.subheader("📚 Archivos de referencia")
    refs = {}
    for tipo, etiqueta in ETIQUETAS.items():
        f = st.file_uploader(etiqueta, type=["xlsx"], key=f"ref_{tipo}")
        if f:
            refs[tipo] = f

firma = (_version_codigo(),
         tuple(sorted((f.name, f.size) for f in (archivos_mes or []))),
         tuple(sorted((t, f.name, f.size) for t, f in refs.items())), mes_proc)
if st.session_state.get('firma') != firma:          # cambiaron los archivos
    for k in ('e1', 'e2'):
        st.session_state.pop(k, None)
    st.session_state['firma'] = firma

NOMBRES = {'emp': "EmpleadosTW", 'equiv': "Equivalencias", 'params': "parametrosMesuales",
           'cot': "cot_afp_hist"}
faltan = [NOMBRES[t] for t in OBLIGATORIOS if t not in refs]
if not meses:
    faltan.insert(0, "meses de TeamWork")
listo = not faltan
if faltan:
    st.info("Para comenzar faltan: **" + ", ".join(faltan) + "**")

st.divider()

# ── ETAPA 1 ─────────────────────────────────────────────────────────
st.subheader("1️⃣ Etapa 1 — Archivo de entrada")
st.caption("Calcula el imponible sin licencia y agrega las columnas **IMP SIN LIC** e "
           "**IMP SL SIS** junto a DIAS LICENCIA.")

if st.button("▶️ Generar archivo de entrada", type="primary", disabled=not listo):
    anteriores = sorted((p for p in meses if p < mes_proc), reverse=True)
    pasos = ([('refs', "Leyendo Equivalencias y parámetros", 1),
              ('emp', "Leyendo listado de empleados", T['emp']),
              (mes_proc, f"Leyendo {isl.nombre_periodo(mes_proc)}", T['mes'])]
             + [(p, f"Leyendo {isl.nombre_periodo(p)}", T['mes']) for p in anteriores]
             + [('calc', "Calculando imponibles sin licencia", 1),
                ('entrada', "Insertando columnas en el archivo de entrada", T['entrada']),
                ('cuad', "Validando cuadratura del líquido", T['mes']),
                ('desc', "Preparando descargas", 3)])
    stt = st.status("Etapa 1 en curso…", expanded=True)
    av = Avance(pasos, stt)
    try:
        av.paso('refs')
        haberes   = _haberes(refs['equiv'].getvalue())
        topes     = _topes(refs['params'].getvalue())
        av.paso('emp')
        contratos = _contratos(refs['emp'].getvalue())
        av.paso(mes_proc)
        actual = _mes(meses[mes_proc].getvalue(), haberes)
        previos = []
        for p in anteriores:
            av.paso(p)
            previos.append(_mes(meses[p].getvalue(), haberes))
        sin_listado = [(f, d['rut']) for f, d in actual['fichas'].items()
                       if f not in contratos]
        av.paso('calc')
        resultados, adv = isl.calcular(actual, previos, topes, contratos)
        av.paso('entrada')
        mes_bytes = meses[mes_proc].getvalue()
        entrada = isl.generar_entrada_con_columna(mes_bytes, resultados, avance=av.avance)
        av.paso('cuad')
        cuadratura = isl.cuadratura_liquido(entrada, refs['equiv'].getvalue(), av.avance)
        av.paso('desc')
        informe = isl.generar_informe(resultados)
        sin_lst_xlsx = None
        if sin_listado:
            detalle_sin = isl.sin_listado(meses[mes_proc].getvalue(), contratos)
            sin_lst_xlsx = isl.generar_informe_sin_listado(detalle_sin, mes_proc)
        cuad_xlsx = isl.generar_cuadratura(cuadratura, mes_proc) if cuadratura else None
        dur = av.cerrar("Etapa 1 lista", limpiar=False)
        stt.update(label=f"✅ Etapa 1 lista en {dur}", state="complete", expanded=False)
        st.session_state['e1'] = {
            'periodo': mes_proc, 'mes_bytes': mes_bytes, 'resultados': resultados,
            'advertencias': [a for a in adv if 'no encontrada en el listado' not in a],
            'sin_listado': sin_listado, 'n_fichas': len(actual['fichas']),
            'entrada': entrada, 'manuales': 0, 'cuadratura': cuadratura,
            'informe': informe, 'cuad_xlsx': cuad_xlsx, 'sin_lst_xlsx': sin_lst_xlsx,
        }
        st.session_state.pop('e2', None)
    except Exception as e:
        av.cerrar(limpiar=False)
        stt.update(label="❌ Etapa 1 con error", state="error", expanded=True)
        st.error(f"❌ Error en la Etapa 1: {e}")
        import traceback
        with st.expander("Detalle del error"):
            st.code(traceback.format_exc())

e1 = st.session_state.get('e1')
if e1:
    res = e1['resultados']
    no_enc = [r for r in res if r['imp'] is None]
    sin_listado = e1.get('sin_listado', [])
    if sin_listado:
        st.error(
            f"⚠️ **{len(sin_listado):,} de {e1['n_fichas']:,} fichas de ".replace(',', '.') +
            f"{isl.nombre_periodo(e1['periodo'])} no están en el listado de empleados.** "
            "Revisa que EmpleadosTW esté actualizado: estas fichas quedan sin contrato ni "
            "sueldo base, y la Etapa 2 las informará como advertencias.")
        with st.expander(f"Ver fichas que no están en el listado ({len(sin_listado)})"):
            st.dataframe([{'FICHA': f, 'RUT': r} for f, r in sin_listado],
                         hide_index=True, width='stretch')
        if e1.get('sin_lst_xlsx'):
            st.download_button(f"⬇️ {isl.nombre_sin_listado(e1['periodo'])}", e1['sin_lst_xlsx'],
                               file_name=isl.nombre_sin_listado(e1['periodo']), mime=XLSX_MIME)
    else:
        st.caption(f"✅ Las {e1['n_fichas']:,} fichas del mes están en el listado de empleados."
                   .replace(',', '.'))

    # Cuadratura: LIQUIDO = (haberes afectos + exentos) - (desc. legales + descuentos)
    cuad = e1.get('cuadratura', [])
    no_cuadran = [f for f in cuad if abs(f['dif']) > isl.TOLERANCIA_CUADRATURA]
    fmt = lambda x: f"{x:,}".replace(',', '.')
    if no_cuadran:
        st.error(f"⚠️ **Cuadratura del líquido: {fmt(len(no_cuadran))} de {fmt(len(cuad))} "
                 "fichas no cuadran** (líquido ≠ haberes − descuentos).")
        with st.expander(f"Ver fichas que no cuadran ({len(no_cuadran)})"):
            st.dataframe([{'RUT': f['rut'], 'FICHA': f['ficha'], 'TOTAL HABERES': f['ha'] + f['he'],
                           'TOTAL DESCUENTOS': f['dl'] + f['de'], 'LIQUIDO CALCULADO': f['calc'],
                           'LIQUIDO TW': f['liq'], 'DIFERENCIA': f['dif']} for f in no_cuadran],
                         hide_index=True, width='stretch')
    elif cuad:
        st.caption(f"✅ Cuadratura del líquido: las {fmt(len(cuad))} fichas cuadran "
                   "(líquido = haberes afectos + exentos − desc. legales − descuentos).")
    m1, m2, m3 = st.columns(3)
    m1.metric("Fichas con licencia", len(res))
    m2.metric("Imponible encontrado", len(res) - len(no_enc))
    m3.metric("Sin imponible", len(no_enc))
    for a in e1['advertencias'][:50]:
        st.warning(a)

    if no_enc:
        with st.expander(f"✍️ Ingresar imponibles faltantes ({len(no_enc)})", expanded=True):
            st.caption("No hay un mes sin licencia para estas fichas. Ingresa el imponible "
                       f"si lo tienes; si lo dejas vacío quedará **{isl.NO_ENCONTRADO}** "
                       "e IMP SL SIS se calculará con el Sueldo Base.")
            editado = st.data_editor(
                [{'RUT': r['rut'], 'FICHA': r['ficha'], 'CONTRATO': r['contrato'],
                  'DIAS LICENCIA': r['dias_lic'], 'SUELDO BASE': round(r.get('sueldo_base') or 0),
                  'IMPONIBLE': None} for r in no_enc],
                column_config={'IMPONIBLE': st.column_config.NumberColumn(
                    'IMPONIBLE', min_value=0, step=1, format="%d")},
                disabled=['RUT', 'FICHA', 'CONTRATO', 'DIAS LICENCIA', 'SUELDO BASE'],
                hide_index=True, width='stretch',
                key=f"editor_{e1['periodo']}")
            manual = {f['FICHA']: f['IMPONIBLE'] for f in editado
                      if f.get('IMPONIBLE') not in (None, '')}
            if st.button(f"💾 Aplicar imponibles ingresados ({len(manual)})",
                         disabled=not manual):
                with Avance([('entrada', "Actualizando archivo de entrada", T['entrada']),
                             ('desc', "Actualizando informe", 1)]) as av:
                    av.paso('entrada')
                    e1['entrada'] = isl.generar_entrada_con_columna(
                        e1['mes_bytes'], res, manual, av.avance)
                    av.paso('desc')
                    e1['informe'] = isl.generar_informe(res, manual)
                    e1['manuales'] = len(manual)
                    e1['manual'] = manual
                st.session_state.pop('e2', None)
                st.rerun()

    st.success(f"✅ Archivo de entrada listo"
               + (f" ({e1['manuales']} imponibles ingresados a mano)" if e1['manuales'] else ""))
    d1, d2, d3 = st.columns(3)
    d1.download_button(f"⬇️ {isl.nombre_entrada(e1['periodo'])}", e1['entrada'],
                       file_name=isl.nombre_entrada(e1['periodo']), mime=XLSX_MIME)
    if 'informe' not in e1:                       # sesión iniciada con una versión anterior
        e1['informe'] = isl.generar_informe(res, e1.get('manual'))
    d2.download_button(f"⬇️ {isl.nombre_informe(e1['periodo'])}", e1['informe'],
                       file_name=isl.nombre_informe(e1['periodo']), mime=XLSX_MIME)
    if cuad and 'cuad_xlsx' not in e1:
        e1['cuad_xlsx'] = isl.generar_cuadratura(cuad, e1['periodo'])
    if e1.get('cuad_xlsx'):
        d3.download_button(f"⬇️ {isl.nombre_cuadratura(e1['periodo'])}", e1['cuad_xlsx'],
                           file_name=isl.nombre_cuadratura(e1['periodo']), mime=XLSX_MIME)

st.divider()

# ── ETAPA 2 ─────────────────────────────────────────────────────────
st.subheader("2️⃣ Etapa 2 — Archivo de salida Rex+")
st.caption("Transforma el archivo de entrada de la Etapa 1 al formato de importación Rex+.")

if st.button("▶️ Generar archivo de salida", type="primary", disabled=not e1):
    stt = st.status("Etapa 2 en curso…", expanded=True)
    av = Avance([('equiv', "Leyendo Equivalencias", 1),
                 ('emp', "Leyendo listado de empleados", T['emp']),
                 ('refs', "Leyendo parámetros y cotizaciones AFP", 1),
                 ('tw', "Leyendo archivo de entrada", T['mes']),
                 ('proc', "Procesando empleados", 2),
                 ('excel', "Generando el Excel de salida", T['excel'])], stt)
    log_box = stt.empty()

    def log_callback(msg):
        log_box.caption(msg.strip())

    try:
        output_rows, advertencias, stats = procesar(
            tw_bytes     = e1['entrada'],
            equiv_bytes  = refs['equiv'].getvalue(),
            emp_bytes    = refs['emp'].getvalue(),
            params_bytes = refs['params'].getvalue(),
            cot_bytes    = refs['cot'].getvalue(),
            asig_bytes   = refs['asig'].getvalue() if 'asig' in refs else None,
            progress_callback = av.avance,
            log_callback      = log_callback,
            paso_callback     = av.paso,
        )
        av.paso('excel')
        excel = generar_excel(output_rows, av.avance)
        dur = av.cerrar("Etapa 2 lista", limpiar=False)
        stt.update(label=f"✅ Etapa 2 lista en {dur}", state="complete", expanded=False)
        st.session_state['e2'] = {
            'excel': excel, 'stats': stats, 'advertencias': advertencias,
        }
    except Exception as e:
        av.cerrar(limpiar=False)
        stt.update(label="❌ Etapa 2 con error", state="error", expanded=True)
        st.error(f"❌ Error en la Etapa 2: {e}")
        import traceback
        with st.expander("Detalle del error"):
            st.code(traceback.format_exc())

e2 = st.session_state.get('e2')
if e2:
    stats, advertencias = e2['stats'], e2['advertencias']
    st.success("✅ Proceso completado")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Período", stats['periodo'])
    m2.metric("Empleados procesados", f"{stats['empleados']:,}")
    m3.metric("Filas generadas", f"{stats['filas']:,}")
    m4.metric("Advertencias", stats['advertencias'])
    st.download_button("⬇️ Descargar salida_rex.xlsx", e2['excel'],
                       file_name=f"salida_rex_{stats['periodo']}.xlsx",
                       mime=XLSX_MIME, type="primary")
    if advertencias:
        with st.expander(f"⚠️ Ver advertencias ({len(advertencias)})"):
            for a in advertencias[:200]:
                st.text(a)
            if len(advertencias) > 200:
                st.caption(f"… y {len(advertencias) - 200} advertencias más.")
