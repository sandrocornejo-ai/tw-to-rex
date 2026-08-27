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

## Archivos necesarios

| Archivo | Descripción |
|---|---|
| `tw.xlsx` | Exportación de liquidaciones desde TeamWork |
| `Equivalencias Tw.xlsx` | Mapeo de conceptos TW → Rex+ |
| `parametrosMensuales.xlsx` | UF, tope AFP, cotizaciones del mes |

## Archivos de salida

- `salida_rex_YYYY-MM.xlsx` — archivo listo para importar en Rex+ (Migraciones)

## Notas

- Las columnas `DESCTO. HORAS ATRASO`, `DESCTO. HORAS ATRASO PT`, `HORAS NO TRABAJADAS $` y `DESC. PAGO EN EXC. IMPONIBLE` se restan de la base imponible aunque estén clasificadas como haber afecto.
- El tope AFP (83,3 UF) se aplica automáticamente.
