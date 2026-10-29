# Salidas para Power BI

Al ejecutar `python capstone_dgt_completo.py` se generan aquí:

- `datos_powerbi.csv` — matriculaciones y porcentajes por provincia, año, tipo y etiqueta.
- `clustering_powerbi.csv` — perfiles y grupos provinciales.
- `proyecciones_powerbi.csv` — proyecciones exploratorias 2026–2030.

Los CSV utilizan `;` como separador de campos y `,` como separador decimal. En pandas: `pd.read_csv("archivo.csv", sep=";", decimal=",")`.

La selección metodológica final para interpretar y presentar las proyecciones se centra en 2026–2027. Las predicciones de matriculaciones y las de porcentajes se obtienen mediante modelos separados y no deben interpretarse como una misma descomposición aritmética.
