# TW → Rex+ Migración de Liquidaciones

App Streamlit que transforma el archivo de liquidaciones de **TeamWork** al formato de importación de **Rex+**.

## Requisitos

```bash
pip install streamlit openpyxl
```

## Uso

```bash
streamlit run tw_to_rex_app.py
```

Luego abre el navegador en `http://localhost:8501` y sigue las instrucciones de la interfaz.

## Flujo (una sola pantalla)

1. **Archivos:** los meses de TeamWork se suben juntos (el mes a procesar y los anteriores);
   la app lee el período de la fila 2 ("Mes a procesar") y propone el más reciente (se puede cambiar).
   Los archivos de referencia se suben de a uno: `EmpleadosTW.xlsx`, `Equivalencias Tw.xlsx`,
   `parametrosMesuales.xlsx`, `cot_afp_hist.xlsx` y `Asig Inst LD.xlsx` (opcional).
2. **Etapa 1 — Archivo de entrada:** inserta junto a `DIAS LICENCIA` las columnas
   - `IMP SIN LIC`: si hay licencia, imponible del último mes anterior sin licencia
     = menor entre la suma de haberes afectos y `topeImp_pesos_afp` de ese mes (0 si no hay licencia).
   - `IMP SL SIS` = (IMP SIN LIC / 30) × DIAS LICENCIA (máximo 30 días).
   Si no hay mes sin licencia se puede ingresar el imponible a mano; si no, IMP SIN LIC queda
   "Imp no encontrado" e IMP SL SIS = (Sueldo Base del listado de empleados / 30) × DIAS LICENCIA (máx. 30).
   Descargas: `<MES AAAA> IMP SIN LIC.xlsx` e informe `ImpSinLic <MES AAAA>.xlsx` (RUT, FICHA, CONTRATO, SUELDO CONTRATO, DIAS LICENCIA, ULT IMP SIN LIC, IMP IMP SIS).
3. **Etapa 2 — Archivo de salida:** transforma el archivo de la Etapa 1 al formato Rex+.

## Archivos necesarios

| Archivo | Descripción |
|---|---|
| `<MES AAAA>.xlsx` | Exportación mensual de TeamWork (mes a procesar y meses anteriores) |
| `Equivalencias Tw.xlsx` | Mapeo de conceptos TW → Rex+ |
| `parametrosMesuales.xlsx` | UF, tope AFP, cotizaciones del mes |
| `cot_afp_hist.xlsx` | Cotizaciones AFP históricas |
| `EmpleadosTW.xlsx` | Listado de empleados (FICHA → contrato) |
| `Asig Inst LD.xlsx` | Instituciones por concepto (opcional) |

## Archivos de salida

- `salida_rex_YYYY-MM.xlsx` — archivo listo para importar en Rex+ (Migraciones)

## Notas

- Las columnas `DESCTO. HORAS ATRASO`, `DESCTO. HORAS ATRASO PT`, `HORAS NO TRABAJADAS $` y `DESC. PAGO EN EXC. IMPONIBLE` se restan de la base imponible aunque estén clasificadas como haber afecto.
- El tope AFP (83,3 UF) se aplica automáticamente.
