# -*- coding: utf-8 -*-
"""Capstone DGT | Evolución de matriculaciones y etiquetas ambientales (2021–2025).

Proyecto integrado: turismos, motocicletas, comparación, clustering K-Means,
regresión lineal y exportaciones para Power BI.

Datos de entrada: ficheros Excel anuales de la DGT, organizados en data/2021,
... data/2025, con los nombres originales que figuran en el README.

Ejecución desde la carpeta del proyecto: python capstone_dgt_completo.py
No se incluyen los archivos de origen: deben descargarse por separado.
"""
from pathlib import Path
import os

# Por defecto, los datos se buscan en la carpeta data junto a este script.
# Se puede definir DGT_DATA_DIR como variable de entorno para otra ubicación.
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DGT_DATA_DIR", BASE_DIR / "data"))
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ARCHIVOS_DGT = {
    2021: "Matriculaciones-Tablas-estadisticas2021.xlsx",
    2022: "Matriculaciones-Tablas-estadisticas-2022.xlsx",
    2023: "Matriculaciones-Tablas-estadisticas-2023.xlsx",
    2024: "Matriculaciones-Tablas-estadisticas-2024.xlsx",
    2025: "Matriculaciones-Tablas-estadisticas-2025.xlsx",
}

faltantes = [DATA_DIR / str(ano) / archivo
            for ano, archivo in ARCHIVOS_DGT.items()
            if not (DATA_DIR / str(ano) / archivo).is_file()]
if faltantes:
    raise FileNotFoundError(
        "Faltan archivos Excel de la DGT. Colócalos en data/<año>/ "
        "o configura DGT_DATA_DIR. Archivos no encontrados:\n" +
        "\n".join(str(ruta) for ruta in faltantes)
    )



# =============================================================================
# TURISMOS — EXTRACCIÓN, LIMPIEZA Y ANÁLISIS EXPLORATORIO
# =============================================================================

import pandas as pd

"""Llamamos los archivos, indicando la hoja especifica con la que vamos a 
trabajar, en este momento solo la de informacion de matriculas medioambientales
de turismos y creando la variable que especifica turismo y el año"""

turismos_2021= pd.read_excel(DATA_DIR / "2021" / "Matriculaciones-Tablas-estadisticas2021.xlsx", 
                             sheet_name= "M_TUR_MEDIO", header = 2)

turismos_2022= pd.read_excel(DATA_DIR / "2022" / "Matriculaciones-Tablas-estadisticas-2022.xlsx", 
                             sheet_name= "M_TUR_MEDIO", header = 2)

turismos_2023= pd.read_excel(DATA_DIR / "2023" / "Matriculaciones-Tablas-estadisticas-2023.xlsx", 
                             sheet_name= "M_TUR_MEDIO", header = 2)

turismos_2024= pd.read_excel(DATA_DIR / "2024" / "Matriculaciones-Tablas-estadisticas-2024.xlsx", 
                             sheet_name= "M_TUR_MEDIO", header = 2)

turismos_2025= pd.read_excel(DATA_DIR / "2025" / "Matriculaciones-Tablas-estadisticas-2025.xlsx", 
                             sheet_name= "M_TUR_MEDIO", header = 2)

"""en el código partimos de la fila donde esta el encabezado de la tabla,
 es por eso que en el codigo de llamado de los archivos colocamos header =2"""
 
"""Procedemos a revisar las variables creadas.
Con la función .shape: vemos la cantidad de filas y columnas que poseen
Con la función .head(): nos muestra la informacion de las 5 primeras columnas
Con la función .info(): vemos si hay nulo y el tipo dato que posee"""


turismos_2021.shape
turismos_2021.head()
turismos_2021.info()

turismos_2022.shape
turismos_2022.head()
turismos_2022.info()

turismos_2023.shape
turismos_2023.head()
turismos_2023.info()

turismos_2024.shape
turismos_2024.head()
turismos_2024.info()


turismos_2025.shape
turismos_2025.head()
turismos_2025.info()


""" Procedemos a agregar una columna en cada uno de los dataframe con el año correpondiente"""
turismos_2021["Año"] = 2021 
turismos_2022["Año"] = 2022 
turismos_2023["Año"] = 2023 
turismos_2024["Año"] = 2024 
turismos_2025["Año"] = 2025 

""" Procedemos a limpiar los títulos de las columnas, unificando criterios, 
quitando espacios entre palabras, quitando los caracteres diacríticos, 
espacios, etc"""

from unidecode import unidecode

turismos_2021.columns =[
    unidecode(col).lower().strip().replace(" ","_")
    for col in turismos_2021.columns]

turismos_2022.columns =[
    unidecode(col).lower().strip().replace(" ","_")
    for col in turismos_2022.columns]

turismos_2023.columns =[
    unidecode(col).lower().strip().replace(" ","_")
    for col in turismos_2023.columns]

turismos_2024.columns =[
    unidecode(col).lower().strip().replace(" ","_")
    for col in turismos_2024.columns]

turismos_2025.columns =[
    unidecode(col).lower().strip().replace(" ","_")
    for col in turismos_2025.columns]

""" Hacemos comprobacion del proceso anterior"""

turismos_2021.columns

turismos_2022.columns

turismos_2023.columns

turismos_2024.columns

turismos_2025.columns



"""En el dataframe de turismos_2022 hay un error en la columna de 
'sin distintivo', procedemos a arreglarlo y verificarlo"""

turismos_2022.columns =[
    unidecode(col).lower().strip().replace(" ","_").replace("\n","_")
    for col in turismos_2022.columns]
turismos_2022.columns



"""Volvemos a hacer comprobaciones de los dataframes, para verificar que en los
 procesos no se hayan generado modicafiones a estos"""
turismos_2021.shape

turismos_2022.shape

turismos_2023.shape

turismos_2024.shape

turismos_2025.shape



turismos_2021.isnull().sum()
turismos_2022.isnull().sum()
turismos_2023.isnull().sum()
turismos_2024.isnull().sum()
turismos_2025.isnull().sum()

turismos_2021.duplicated().sum()
turismos_2022.duplicated().sum()
turismos_2023.duplicated().sum()
turismos_2024.duplicated().sum()
turismos_2025.duplicated().sum()



""" Vamos a proceder a crear un nuevo dataframe, donde a través de la función 
concat() unamos todos los dataframe anteriores y hacemos las comprobaciones
correspondientes"""

turismos = pd.concat([turismos_2021, turismos_2022, turismos_2023,
                      turismos_2024, turismos_2025],ignore_index=True)

turismos = turismos[
    turismos["provincias"].str.strip().str.lower() != "total"
].reset_index(drop=True)


turismos.info()
turismos.shape
turismos.head()
turismos.tail()
turismos["ano"].value_counts()

turismos.isnull().sum()

turismos["provincias"].nunique()
turismos.duplicated(subset=["provincias","ano"]).sum()


"""Creamos una nueva columna con la sumatoria de las matriculaciones totales de
cada provicia y año"""

turismos["total_turismos"] = turismos[["cero","b","c","eco","sin_distintivo"]].sum(axis=1)

turismos.isnull().sum()

turismos["provincias"].nunique()
turismos.duplicated(subset=["provincias","ano"]).sum()

"""Creamos un dataframe que nos de la sumatoria de matriculaciones por año"""
turismos_por_ano = turismos.groupby("ano",as_index=False)["total_turismos"].sum()

"""Hacemos comprobación del nuevo dataframe (turismos_por_ano), para que no 
hayan errores"""

turismos_por_ano.shape
turismos_por_ano.info()
turismos_por_ano.head()

"""Creamos un dataframe que nos de la sumatoria de las matriculaciones de 
cada etiqueta por año"""

etiquetas_por_ano = turismos.groupby(
    "ano",as_index=False)[["cero","b","c","eco","sin_distintivo"]].sum()



"""Hacemos comprobación del nuevo dataframe (etiquetas_por_ano), para que no 
hayan errores"""
etiquetas_por_ano.shape
etiquetas_por_ano.info()
etiquetas_por_ano.head()


"""Creamos una nueva columna en el dataframe de etiquetas_por_ano, 
total_turismos, la cual nos va a arrojar la sumatoria anual de matriculaciones 
por año"""

etiquetas_por_ano["total_turismos"]= etiquetas_por_ano[[
    "cero","b","c","eco","sin_distintivo"]].sum(axis=1)



"""Confirmamos que las totalizaciones dan el mismo número en ambos dataframe"""
turismos_por_ano[["ano","total_turismos"]]
etiquetas_por_ano[["ano","total_turismos"]]

"""Procedemos a crear nuevas columnas en el dataframe, etiquetas_por_ano, 
donde comenzamos a calcular los porcentajes de matriculaciones de cada etiqueta 
respecto al año correspondiente"""


etiquetas_por_ano["porcentaje_cero"]= etiquetas_por_ano[
    "cero"] / etiquetas_por_ano["total_turismos"] * 100 

etiquetas_por_ano["porcentaje_b"]= etiquetas_por_ano[
    "b"] / etiquetas_por_ano["total_turismos"] * 100

etiquetas_por_ano["porcentaje_c"]= etiquetas_por_ano[
    "c"] / etiquetas_por_ano["total_turismos"] * 100

etiquetas_por_ano["porcentaje_eco"]= etiquetas_por_ano[
    "eco"] / etiquetas_por_ano["total_turismos"] * 100

etiquetas_por_ano["porcentaje_sin_distintivo"]= etiquetas_por_ano[
    "sin_distintivo"] / etiquetas_por_ano["total_turismos"] * 100

"""

HEMOS HECHO UN .ROUN(2), PERO AL EJECUTARLO HEMOS VISTO QUE NO NOS DA 100% EXACTO
POR TANTO LO DEJAMOS SIN EL REDONDE POR LOS MOMENTOS, PARA PREGUNTAR EN TUTORIA Y QUIZAS
CORREGIR MAS ADELANTE A TRAVÉS DE POWER BI


"""

""" En el dataframe turismos_por_ano, vamos a calcular el porcentaje de variación
de las matriculaciones con respecto al año anterior"""


turismos_por_ano["variacion_porcentual"] = turismos_por_ano[
    "total_turismos"].pct_change() * 100



"""En el dataframe etiquetas_por_ano, vamos a calcular el porcentaje de variación
de las matriculaciones con respecto al año anterior"""


etiquetas_por_ano["variacion_cero"] =etiquetas_por_ano["cero"].pct_change() * 100

etiquetas_por_ano["variacion_b"] =etiquetas_por_ano["b"].pct_change() * 100

etiquetas_por_ano["variacion_c"] =etiquetas_por_ano["c"].pct_change() * 100

etiquetas_por_ano["variacion_eco"] =etiquetas_por_ano["eco"].pct_change() * 100

etiquetas_por_ano["variacion_sin_distintivo"] =etiquetas_por_ano["sin_distintivo"].pct_change() * 100

""" Vamos a calcular la variación porcentual de cada etiqueta con respecto al 
período 2021 y los consiguientes años"""

etiquetas_por_ano["cambio_cero"] = (
    etiquetas_por_ano["porcentaje_cero"] -
    etiquetas_por_ano["porcentaje_cero"].iloc[0])

etiquetas_por_ano["cambio_b"] = (
    etiquetas_por_ano["porcentaje_b"] -
    etiquetas_por_ano["porcentaje_b"].iloc[0])

etiquetas_por_ano["cambio_c"] = (
    etiquetas_por_ano["porcentaje_c"] -
    etiquetas_por_ano["porcentaje_c"].iloc[0])

etiquetas_por_ano["cambio_eco"] = (
    etiquetas_por_ano["porcentaje_eco"] -
    etiquetas_por_ano["porcentaje_eco"].iloc[0])

etiquetas_por_ano["cambio_sin_distintivo"] = (
    etiquetas_por_ano["porcentaje_sin_distintivo"] -
    etiquetas_por_ano["porcentaje_sin_distintivo"].iloc[0])

""" Creamos una nueva columna en el dataframe etiquetas_por_ano, para ver cual 
es la más matriculada cada año y así ver como va cambiando"""


etiquetas_por_ano["etiqueta_mas_matriculada"] = etiquetas_por_ano[
    ["cero", "b", "c", "eco", "sin_distintivo"]].idxmax(axis=1)

etiquetas_por_ano.loc[:, ["ano", "etiqueta_mas_matriculada"]]


"""Ahora vamos a pasar a hacer analisis de las matriculaciones en función 
a las provincias, comenzamos por hacer una sumatoria de la matriculación de 
cada provincia en los años de estudio"""


top_provincias = turismos.groupby("provincias", as_index=False)["total_turismos"].sum()

top_provincias = top_provincias.sort_values("total_turismos",ascending=False)


top_provincias.head(10)


"""Eliminamos la fila de total del dataframe top_provincias"""

top_provincias = top_provincias[
    top_provincias["provincias"] != "Total"]

top_provincias.head(10)


"""Vamos a empezar a crear rankings de las etiquetas medioambientales, para
así ir viendo donde se matriculan más de cada una de ellas

Creamos un nuevo dataframe, para realizar un ranking por etiqueta medioambiental
en cada provincia"""


ranking_provincias = turismos.groupby("provincias",as_index=False)[[
    "cero","b","c","eco","sin_distintivo"]].sum()



"""Eliminamos la fila de Total del dataframe"""
ranking_provincias = ranking_provincias[ranking_provincias["provincias"]!= "Total"]


"""Ahora si las vamos a organizar como un ranking de mayor a menor, 
por cada tipo de etiqueta.

Usamos method=min para que si hay igualdad se repita el mismo número de ranking
y luego pase al siguiente"""

ranking_provincias["ranking_cero"] = ranking_provincias["cero"].rank(ascending=False,method="min")

ranking_provincias["ranking_b"] = ranking_provincias["b"].rank(ascending=False,method="min")

ranking_provincias["ranking_c"] = ranking_provincias["c"].rank(ascending=False,method="min")

ranking_provincias["ranking_eco"] = ranking_provincias["eco"].rank(ascending=False,method="min")

ranking_provincias["ranking_sin_distintivo"] = ranking_provincias["sin_distintivo"].rank(ascending=False,method="min")


"""Empezamos a analizar como evolucionan las etiquetas año a año

Hacemos un gráfico de líneas"""

import matplotlib.pyplot as plt

plt.figure(figsize=(10, 6))

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["porcentaje_cero"],
    marker="o",
    label="CERO")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["porcentaje_b"],
    marker="o",
    label="B")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["porcentaje_c"],
    marker="o",
    label="C")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["porcentaje_eco"],
    marker="o",
    label="ECO")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["porcentaje_sin_distintivo"],
    marker="o",
    label="Sin distintivo")

plt.title("Evolución de las etiquetas medioambientales en turismos")
plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")
plt.xticks(etiquetas_por_ano["ano"])
plt.legend()
plt.grid(True)
plt.show()



"""vamos a hacer un gráfico también con las cantidades absolutas"""

plt.figure(figsize=(10, 6))

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["cero"],
    marker="o",
    label="CERO")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["b"],
    marker="o",
    label="B")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["c"],
    marker="o",
    label="C")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["eco"],
    marker="o",
    label="ECO")

plt.plot(
    etiquetas_por_ano["ano"],
    etiquetas_por_ano["sin_distintivo"],
    marker="o",
    label="Sin distintivo")

plt.title("Evolución de las matriculaciones de turismos por etiqueta")
plt.xlabel("Año")
plt.ylabel("Número de matriculaciones")
plt.xticks(etiquetas_por_ano["ano"])
plt.legend()
plt.grid(True)
plt.show()


#Ya podemos empezar a ver através de gráficos la evolución según etiqueta


""" Veamos la evolución de las matriculaciones según las provincias de España, 
del año 2021 con el 2025"""

evolucion_provincias= turismos.pivot(index="provincias",columns="ano",values=[
    "cero","b","c","eco","sin_distintivo"])


evolucion_provincias.columns=[f"{etiqueta}_{ano}"
                              for etiqueta, ano in evolucion_provincias.columns]


"""Veamos comparación de las matriculaciones según etiquetas en números 
absolutos entre 2021 y 2025"""


evolucion_provincias["evolucion_cero"]= (evolucion_provincias["cero_2025"] - evolucion_provincias["cero_2021"])

evolucion_provincias["evolucion_b"]= (evolucion_provincias["b_2025"] - evolucion_provincias["b_2021"])

evolucion_provincias["evolucion_c"]= (evolucion_provincias["c_2025"] - evolucion_provincias["c_2021"])

evolucion_provincias["evolucion_eco"]= (evolucion_provincias["eco_2025"] - evolucion_provincias["eco_2021"])

evolucion_provincias["evolucion_sin_distintivo"]= (evolucion_provincias["sin_distintivo_2025"] - evolucion_provincias["sin_distintivo_2021"])


"""Agregamos en el data frame turismos, nuevas columnas con los porcentajes de 
matriculación de cada etiqueta"""

turismos["porcentaje_cero"]= (turismos["cero"] / turismos["total_turismos"] * 100)

turismos["porcentaje_b"]= (turismos["b"] / turismos["total_turismos"] * 100)

turismos["porcentaje_c"]= (turismos["c"] / turismos["total_turismos"] * 100)

turismos["porcentaje_eco"]= (turismos["eco"] / turismos["total_turismos"] * 100)

turismos["porcentaje_sin_distintivo"]= (turismos["sin_distintivo"] / turismos["total_turismos"] * 100)

""" Procedemos a calcular la evolución del porcentaje de cada etiqueta por provincia
con respecto a los años 2021 y 2025"""

evolucion_provincias= turismos[turismos["ano"].isin([2021,2025])]




evolucion_provincias = evolucion_provincias.pivot(index="provincias",
                                                  columns= "ano",
                                                  values=["porcentaje_cero",
                                                          "porcentaje_b",
                                                          "porcentaje_c",
                                                          "porcentaje_eco",
                                                          "porcentaje_sin_distintivo"])


evolucion_provincias.columns=[f"{etiqueta}_{ano}"
                              for etiqueta, ano in evolucion_provincias.columns]



evolucion_provincias["cambio_cero"] = (
    evolucion_provincias["porcentaje_cero_2025"]
    - evolucion_provincias["porcentaje_cero_2021"])

evolucion_provincias["cambio_b"] = (
    evolucion_provincias["porcentaje_b_2025"]
    - evolucion_provincias["porcentaje_b_2021"])

evolucion_provincias["cambio_c"] = (
    evolucion_provincias["porcentaje_c_2025"]
    - evolucion_provincias["porcentaje_c_2021"])

evolucion_provincias["cambio_eco"] = (
    evolucion_provincias["porcentaje_eco_2025"]
    - evolucion_provincias["porcentaje_eco_2021"])

evolucion_provincias["cambio_sin_distintivo"] = (
    evolucion_provincias["porcentaje_sin_distintivo_2025"]
    - evolucion_provincias["porcentaje_sin_distintivo_2021"])




""" queremos también tener los valores absolutos dentro del dataframe,
 por tanto, creamos uno nuevo, para unirlos al dataframe anterior"""

valores_absolutos = turismos[turismos["ano"].isin([2021,2025])][["provincias",
                                                                 "ano",
                                                                 "cero","b","c",
                                                                 "eco",
                                                                 "sin_distintivo",
                                                                 "total_turismos"]]

valores_absolutos = valores_absolutos.pivot(index = "provincias",
                                            columns = "ano",
                                            values = ["cero",
                                                      "b",
                                                      "c",
                                                      "eco",
                                                      "sin_distintivo",
                                                      "total_turismos"])

valores_absolutos.columns = [f"{etiqueta}_{ano}"
    for etiqueta, ano in valores_absolutos.columns]

evolucion_provincias = evolucion_provincias.join(valores_absolutos)




turismos.head()
turismos.info()
turismos[["provincias", "cero", "b", "c", "eco",
     "sin_distintivo", "ano", "total_turismos"]].head()































pd.set_option("display.max_columns", None)
pd.reset_option("display.max_columns")


# =============================================================================
# MOTOCICLETAS — EXTRACCIÓN, LIMPIEZA Y ANÁLISIS EXPLORATORIO
# =============================================================================

import pandas as pd

"""Llamamos los archivos, indicando la hoja especifica con la que vamos a 
trabajar, en este momento solo la de informacion de matriculas medioambientales
de motocicletas y creando la variable que especifica motocicletas y el año"""

motos_2021= pd.read_excel(DATA_DIR / "2021" / "Matriculaciones-Tablas-estadisticas2021.xlsx",sheet_name= "M_MOTO_MEDIO", header = 2)

motos_2022= pd.read_excel(DATA_DIR / "2022" / "Matriculaciones-Tablas-estadisticas-2022.xlsx",sheet_name= "M_MOTO_MEDIO", header = 2)

motos_2023= pd.read_excel(DATA_DIR / "2023" / "Matriculaciones-Tablas-estadisticas-2023.xlsx",sheet_name= "M_MOTO_MEDIO", header = 2)

motos_2024= pd.read_excel(DATA_DIR / "2024" / "Matriculaciones-Tablas-estadisticas-2024.xlsx",sheet_name= "M_MOTO_MEDIO", header = 2)

motos_2025= pd.read_excel(DATA_DIR / "2025" / "Matriculaciones-Tablas-estadisticas-2025.xlsx",sheet_name= "M_MOTO_MEDIO", header = 2)


"""en el código partimos de la fila donde esta el encabezado de la tabla,
 es por eso que en el codigo de llamado de los archivos colocamos header =2"""
 
"""Procedemos a revisar las variables creadas.
Con la función .shape: vemos la cantidad de filas y columnas que poseen
Con la función .head(): nos muestra la informacion de las 5 primeras columnas
Con la función .info(): vemos si hay nulo y el tipo dato que posee"""

motos_2021.shape
motos_2021.info()
motos_2021.head()

motos_2022.shape
motos_2022.info()
motos_2022.head()

motos_2023.shape
motos_2023.info()
motos_2023.head()

motos_2024.shape
motos_2024.info()
motos_2024.head()

motos_2025.shape
motos_2025.info()
motos_2025.head()


""" Procedemos a agregar una columna en cada uno de los dataframe con el año correpondiente"""

motos_2021["Año"] = 2021

motos_2022["Año"] = 2022 

motos_2023["Año"] = 2023
 
motos_2024["Año"] = 2024 

motos_2025["Año"] = 2025


""" Procedemos a limpiar los títulos de las columnas, unificando criterios, 
quitando espacios entre palabras, quitando los caracteres diacríticos, 
espacios, etc"""

from unidecode import unidecode

motos_2021.columns=[unidecode(col).lower().strip().replace(" ","_").replace("\n","_")
                    for col in motos_2021.columns]

motos_2022.columns=[unidecode(col).lower().strip().replace(" ","_").replace("\n","_")
                    for col in motos_2022.columns]

motos_2023.columns = [unidecode(col).lower().strip().replace(" ","_").replace("\n","_")
                      for col in motos_2023.columns]

motos_2024.columns=[unidecode(col).lower().strip().replace(" ","_").replace("\n","_")
                    for col in motos_2024.columns]

motos_2025.columns=[unidecode(col).lower().strip().replace(" ","_").replace("\n","_")
                    for col in motos_2025.columns]


""" Hacemos comprobacion del proceso anterior"""
motos_2021.columns
motos_2022.columns
motos_2023.columns
motos_2024.columns
motos_2025.columns


"""Volvemos a hacer comprobaciones de los dataframes, para verificar que en los
 procesos no se hayan generado modicafiones a estos"""
 
motos_2021.shape
motos_2022.shape
motos_2023.shape
motos_2024.shape
motos_2025.shape

motos_2021.isnull().sum()
motos_2022.isnull().sum()
motos_2023.isnull().sum()
motos_2024.isnull().sum()
motos_2025.isnull().sum()

motos_2021.duplicated().sum()
motos_2022.duplicated().sum()
motos_2023.duplicated().sum()
motos_2024.duplicated().sum()
motos_2025.duplicated().sum()


""" Vamos a proceder a crear un nuevo dataframe, donde a través de la función 
concat() unamos todos los dataframe anteriores y hacemos las comprobaciones
correspondientes"""

motos = pd.concat([motos_2021,motos_2022,motos_2023,motos_2024,motos_2025],ignore_index=True)

motos = motos[
    motos["provincias"].str.strip().str.lower() != "total"
].reset_index(drop=True)

motos.info()
motos.shape
motos.head()
motos.tail()
motos["ano"].value_counts()
motos.isnull().sum()

motos["provincias"].nunique()

motos.duplicated(subset=["provincias","ano"]).sum()


"""Creamos una nueva columna con la sumatoria de las matriculaciones totales de
cada provicia y año"""

motos["total_motos"]=motos[["cero","b","c","eco","sin_distintivo"]].sum(axis=1)

motos.isnull().sum()
motos["provincias"].nunique()
motos.duplicated(subset=["provincias","ano"]).sum()

"""Creamos un dataframe que nos de la sumatoria de matriculaciones por año"""

motos_por_ano = motos.groupby("ano", as_index=False)["total_motos"].sum()

"""Hacemos comprobación del nuevo dataframe (motos_por_ano), para que no 
hayan errores"""
motos_por_ano.shape
motos_por_ano.info()
motos_por_ano.head()

"""Creamos un dataframe que nos de la sumatoria de las matriculaciones de 
cada etiqueta por año"""

etiquetas_motos = motos.groupby("ano", as_index = False)[["cero","b","c","eco","sin_distintivo"]].sum()


"""Hacemos comprobación del nuevo dataframe (etiquetas_motos), para que no 
hayan errores"""

etiquetas_motos.shape
etiquetas_motos.info()
etiquetas_motos.head()

"""Creamos una nueva columna en el dataframe de etiquetas_motos, 
total_motos, la cual nos va a arrojar la sumatoria anual de matriculaciones 
por año"""

etiquetas_motos["total_motos"] = etiquetas_motos[[
    "cero","b","c","eco","sin_distintivo"]].sum(axis = 1)

"""Confirmamos que las totalizaciones dan el mismo número en ambos dataframe"""

motos_por_ano[["ano","total_motos"]]
etiquetas_motos[["ano","total_motos"]]

"""Procedemos a crear nuevas columnas en el dataframe, etiquetas_motos, 
donde comenzamos a calcular los porcentajes de matriculaciones de cada etiqueta 
respecto al año correspondiente"""

etiquetas_motos["porcentaje_cero"]= etiquetas_motos[
    "cero"] / etiquetas_motos["total_motos"] * 100

etiquetas_motos["porcentaje_b"]= etiquetas_motos[
    "b"] / etiquetas_motos["total_motos"] * 100

etiquetas_motos["porcentaje_c"]= etiquetas_motos[
    "c"] / etiquetas_motos["total_motos"] * 100

etiquetas_motos["porcentaje_eco"]= etiquetas_motos[
    "eco"] / etiquetas_motos["total_motos"] * 100

etiquetas_motos["porcentaje_sin_distintivo"]= etiquetas_motos[
    "sin_distintivo"] / etiquetas_motos["total_motos"] * 100


"""

HEMOS HECHO UN .ROUN(2), PERO AL EJECUTARLO HEMOS VISTO QUE NO NOS DA 100% EXACTO
POR TANTO LO DEJAMOS SIN EL REDONDE POR LOS MOMENTOS, PARA PREGUNTAR EN TUTORIA Y QUIZAS
CORREGIR MAS ADELANTE A TRAVÉS DE POWER BI


"""


""" En el dataframe motos_por_ano, vamos a calcular el porcentaje de variación
de las matriculaciones con respecto al año anterior"""

motos_por_ano["variacion_porcentual"] = motos_por_ano[
    "total_motos"].pct_change() * 100


"""En el dataframe etiquetas_por_ano, vamos a calcular el porcentaje de variación
de las matriculaciones con respecto al año anterior"""

etiquetas_motos["variacion_cero"] = etiquetas_motos[
    "cero"].pct_change() * 100

etiquetas_motos["variacion_b"] = etiquetas_motos[
    "b"].pct_change() * 100

etiquetas_motos["variacion_c"] = etiquetas_motos[
    "c"].pct_change() * 100

etiquetas_motos["variacion_eco"] = etiquetas_motos[
    "eco"].pct_change() * 100

etiquetas_motos["variacion_sin_distintivo"] = etiquetas_motos[
    "sin_distintivo"].pct_change() * 100


""" Vamos a calcular la variación porcentual de cada etiqueta con respecto al 
período 2021 y los consiguientes años"""

etiquetas_motos["cambio_cero"] = (
    etiquetas_motos["porcentaje_cero"] - 
    etiquetas_motos["porcentaje_cero"].iloc[0])

etiquetas_motos["cambio_b"] = (
    etiquetas_motos["porcentaje_b"] -
    etiquetas_motos["porcentaje_b"].iloc[0])

etiquetas_motos["cambio_c"] = (
    etiquetas_motos["porcentaje_c"] -
    etiquetas_motos["porcentaje_c"].iloc[0])

etiquetas_motos["cambio_eco"] = (
    etiquetas_motos["porcentaje_eco"] -
    etiquetas_motos["porcentaje_eco"].iloc[0])

etiquetas_motos["cambio_sin_distintivo"] = (
    etiquetas_motos["porcentaje_sin_distintivo"] -
    etiquetas_motos["porcentaje_sin_distintivo"].iloc[0])


""" Creamos una nueva columna en el dataframe etiquetas_por_ano, para ver cual 
es la más matriculada cada año y así ver como va cambiando"""

etiquetas_motos["etiqueta_mas_matriculada"] = etiquetas_motos[
    ["cero", "b", "c", "eco", "sin_distintivo"]
].idxmax(axis=1)

etiquetas_motos.loc[:,["ano","etiqueta_mas_matriculada"]]


"""Ahora vamos a pasar a hacer analisis de las matriculaciones en función 
a las provincias, comenzamos por hacer una sumatoria de la matriculación de 
cada provincia en los años de estudio"""

top_provincias = motos.groupby("provincias", as_index=False)["total_motos"].sum()

top_provincias = top_provincias.sort_values("total_motos",ascending=False)

top_provincias.head(10)


"""Eliminamos la fila de total del dataframe top_provincias"""

top_provincias = top_provincias[
    top_provincias["provincias"] != "Total"]

top_provincias.head(10)




"""Vamos a empezar a crear rankings de las etiquetas medioambientales, para
así ir viendo donde se matriculan más de cada una de ellas

Creamos un nuevo dataframe, para realizar un ranking por etiqueta medioambiental
en cada provincia"""


ranking_provincias = motos.groupby("provincias",as_index=False)[[
    "cero","b","c","eco","sin_distintivo"]].sum()



"""Eliminamos la fila de Total del dataframe"""
ranking_provincias = ranking_provincias[ranking_provincias["provincias"]!= "Total"]


"""Ahora si las vamos a organizar como un ranking de mayor a menor, 
por cada tipo de etiqueta.

Usamos method=min para que si hay igualdad se repita el mismo número de ranking
y luego pase al siguiente"""

ranking_provincias["ranking_cero"] = ranking_provincias["cero"].rank(ascending=False,method="min")

ranking_provincias["ranking_b"] = ranking_provincias["b"].rank(ascending=False,method="min")

ranking_provincias["ranking_c"] = ranking_provincias["c"].rank(ascending=False,method="min")

ranking_provincias["ranking_eco"] = ranking_provincias["eco"].rank(ascending=False,method="min")

ranking_provincias["ranking_sin_distintivo"] = ranking_provincias["sin_distintivo"].rank(ascending=False,method="min")




"""Empezamos a analizar como evolucionan las etiquetas año a año

Hacemos un gráfico de líneas"""


import matplotlib.pyplot as plt

plt.figure(figsize=(10, 6))

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["porcentaje_cero"],
    marker="o",
    label="CERO")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["porcentaje_b"],
    marker="o",
    label="B")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["porcentaje_c"],
    marker="o",
    label="C")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["porcentaje_eco"],
    marker="o",
    label="ECO")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["porcentaje_sin_distintivo"],
    marker="o",
    label="Sin distintivo")

plt.title("Evolución de las etiquetas medioambientales en motocicletas")
plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")
plt.xticks(etiquetas_motos["ano"])
plt.legend()
plt.grid(True)
plt.show()



"""vamos a hacer un gráfico también con las cantidades absolutas"""

plt.figure(figsize=(10, 6))

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["cero"],
    marker="o",
    label="CERO")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["b"],
    marker="o",
    label="B")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["c"],
    marker="o",
    label="C")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["eco"],
    marker="o",
    label="ECO")

plt.plot(
    etiquetas_motos["ano"],
    etiquetas_motos["sin_distintivo"],
    marker="o",
    label="Sin distintivo")

plt.title("Evolución de las matriculaciones de motocicletas por etiqueta")
plt.xlabel("Año")
plt.ylabel("Número de matriculaciones")
plt.xticks(etiquetas_motos["ano"])
plt.legend()
plt.grid(True)
plt.show()

#Ya podemos empezar a ver através de gráficos la evolución según etiqueta

""" Veamos la evolución de las matriculaciones según las provincias de España, 
del año 2021 con el 2025"""

evolucion_provincias= motos.pivot(index="provincias",columns="ano",values=[
    "cero","b","c","eco","sin_distintivo"])


evolucion_provincias.columns=[f"{etiqueta}_{ano}"
                              for etiqueta, ano in evolucion_provincias.columns]


"""Veamos comparación de las matriculaciones según etiquetas en números 
absolutos entre 2021 y 2025"""


evolucion_provincias["evolucion_cero"]= (evolucion_provincias["cero_2025"] - evolucion_provincias["cero_2021"])

evolucion_provincias["evolucion_b"]= (evolucion_provincias["b_2025"] - evolucion_provincias["b_2021"])

evolucion_provincias["evolucion_c"]= (evolucion_provincias["c_2025"] - evolucion_provincias["c_2021"])

evolucion_provincias["evolucion_eco"]= (evolucion_provincias["eco_2025"] - evolucion_provincias["eco_2021"])

evolucion_provincias["evolucion_sin_distintivo"]= (evolucion_provincias["sin_distintivo_2025"] - evolucion_provincias["sin_distintivo_2021"])


"""Agregamos en el data frame motos, nuevas columnas con los porcentajes de 
matriculación de cada etiqueta"""

motos["porcentaje_cero"]= (motos["cero"] / motos["total_motos"] * 100)

motos["porcentaje_b"]= (motos["b"] / motos["total_motos"] * 100)

motos["porcentaje_c"]= (motos["c"] / motos["total_motos"] * 100)

motos["porcentaje_eco"]= (motos["eco"] / motos["total_motos"] * 100)

motos["porcentaje_sin_distintivo"]= (motos["sin_distintivo"] / motos["total_motos"] * 100)

""" Procedemos a calcular la evolución del porcentaje de cada etiqueta por provincia
con respecto a los años 2021 y 2025"""

evolucion_provincias= motos[motos["ano"].isin([2021,2025])]




evolucion_provincias = evolucion_provincias.pivot(index="provincias",
                                                  columns= "ano",
                                                  values=["porcentaje_cero",
                                                          "porcentaje_b",
                                                          "porcentaje_c",
                                                          "porcentaje_eco",
                                                          "porcentaje_sin_distintivo"])


evolucion_provincias.columns=[f"{etiqueta}_{ano}"
                              for etiqueta, ano in evolucion_provincias.columns]



evolucion_provincias["cambio_cero"] = (
    evolucion_provincias["porcentaje_cero_2025"]
    - evolucion_provincias["porcentaje_cero_2021"])

evolucion_provincias["cambio_b"] = (
    evolucion_provincias["porcentaje_b_2025"]
    - evolucion_provincias["porcentaje_b_2021"])

evolucion_provincias["cambio_c"] = (
    evolucion_provincias["porcentaje_c_2025"]
    - evolucion_provincias["porcentaje_c_2021"])

evolucion_provincias["cambio_eco"] = (
    evolucion_provincias["porcentaje_eco_2025"]
    - evolucion_provincias["porcentaje_eco_2021"])

evolucion_provincias["cambio_sin_distintivo"] = (
    evolucion_provincias["porcentaje_sin_distintivo_2025"]
    - evolucion_provincias["porcentaje_sin_distintivo_2021"])



""" queremos también tener los valores absolutos dentro del dataframe,
 por tanto, creamos uno nuevo, para unirlos al dataframe anterior"""

valores_absolutos = motos[motos["ano"].isin([2021,2025])][["provincias",
                                                                 "ano",
                                                                 "cero","b","c",
                                                                 "eco",
                                                                 "sin_distintivo",
                                                                 "total_motos"]]

valores_absolutos = valores_absolutos.pivot(index = "provincias",
                                            columns = "ano",
                                            values = ["cero",
                                                      "b",
                                                      "c",
                                                      "eco",
                                                      "sin_distintivo",
                                                      "total_motos"])

valores_absolutos.columns = [f"{etiqueta}_{ano}"
    for etiqueta, ano in valores_absolutos.columns]

evolucion_provincias = evolucion_provincias.join(valores_absolutos)


# =============================================================================
# COMPARACIÓN, MACHINE LEARNING Y EXPORTACIÓN POWER BI
# =============================================================================

import pandas as pd

# Seleccionamos las columnas que necesitamos para comparar ambos vehículos.

columnas = [
    "ano",
    "cero",
    "b",
    "c",
    "eco",
    "sin_distintivo",
    "porcentaje_cero",
    "porcentaje_b",
    "porcentaje_c",
    "porcentaje_eco",
    "porcentaje_sin_distintivo"]

# Creamos una copia de los datos de turismos.
comparacion_turismos = etiquetas_por_ano[columnas].copy()

# Añadimos una columna para identificar el tipo de vehículo.
comparacion_turismos["tipo"] = "turismos"

# Añadimos el total de matriculaciones.
comparacion_turismos["total_matriculaciones"] = etiquetas_por_ano["total_turismos"]


# Repetimos el proceso con motocicletas.
comparacion_motos = etiquetas_motos[columnas].copy()

comparacion_motos["tipo"] = "motocicletas"

comparacion_motos["total_matriculaciones"] = etiquetas_motos["total_motos"]


# Unimos ambos DataFrames.
comparacion_etiquetas = pd.concat(
    [comparacion_turismos, comparacion_motos],
    ignore_index=True)

# Ordenamos por año y tipo de vehículo.
comparacion_etiquetas = comparacion_etiquetas.sort_values(
    ["ano", "tipo"]).reset_index(drop=True)

# Comprobamos el resultado.
comparacion_etiquetas.head(10)



"""Vamos a comprobar los porcentajes"""
comparacion_etiquetas.loc[:,["ano", "tipo", "porcentaje_cero", "porcentaje_c", "porcentaje_eco"]].round(2)



"""Comparamos la evolución de la etiqueta C entre turismos
y motocicletas durante el período 2021-2025"""

import matplotlib.pyplot as plt

# Separamos los datos por tipo de vehículo.

datos_turismos = comparacion_etiquetas[comparacion_etiquetas["tipo"] == "turismos"]

datos_motos = comparacion_etiquetas[comparacion_etiquetas["tipo"] == "motocicletas"]


# Creamos el gráfico.

plt.figure(figsize=(10, 6))

plt.plot(
    datos_turismos["ano"],
    datos_turismos["porcentaje_c"],
    marker="o",
    label="Turismos")

plt.plot(
    datos_motos["ano"],
    datos_motos["porcentaje_c"],
    marker="o",
    label="Motocicletas")

plt.title("Evolución de la etiqueta C: turismos vs motocicletas")

plt.xlabel("Año")

plt.ylabel("Porcentaje de matriculaciones (%)")

plt.xticks(comparacion_etiquetas["ano"].unique())

plt.ylim(0, 100)

plt.legend()

plt.grid(True)

plt.show()


"""Comparamos la evolución de la etiqueta ECO entre turismos
y motocicletas durante el período 2021-2025"""

plt.figure(figsize=(10, 6))

plt.plot(
    datos_turismos["ano"],
    datos_turismos["porcentaje_eco"],
    marker="o",
    label="Turismos")

plt.plot(
    datos_motos["ano"],
    datos_motos["porcentaje_eco"],
    marker="o",
    label="Motocicletas")

plt.title("Evolución de la etiqueta ECO: turismos vs motocicletas")

plt.xlabel("Año")

plt.ylabel("Porcentaje de matriculaciones (%)")

plt.xticks(comparacion_etiquetas["ano"].unique())

plt.ylim(0, 100)

plt.legend()

plt.grid(True)

plt.show()

"""Comparamos la evolución de la etiqueta Cero entre turismos
y motocicletas durante el período 2021-2025"""

plt.figure(figsize=(10, 6))

plt.plot(
    datos_turismos["ano"],
    datos_turismos["porcentaje_cero"],
    marker="o",
    label="Turismos")

plt.plot(
    datos_motos["ano"],
    datos_motos["porcentaje_cero"],
    marker="o",
    label="Motocicletas")

plt.title("Evolución de la etiqueta Cero: turismos vs motocicletas")

plt.xlabel("Año")

plt.ylabel("Porcentaje de matriculaciones (%)")

plt.xticks(comparacion_etiquetas["ano"].unique())

plt.ylim(0, 100)

plt.legend()

plt.grid(True)

plt.show()

"""Comparamos la evolución de la etiqueta B entre turismos
y motocicletas durante el período 2021-2025"""

plt.figure(figsize=(10, 6))

plt.plot(
    datos_turismos["ano"],
    datos_turismos["porcentaje_b"],
    marker="o",
    label="Turismos")

plt.plot(
    datos_motos["ano"],
    datos_motos["porcentaje_b"],
    marker="o",
    label="Motocicletas")

plt.title("Evolución de la etiqueta B: turismos vs motocicletas")

plt.xlabel("Año")

plt.ylabel("Porcentaje de matriculaciones (%)")

plt.xticks(comparacion_etiquetas["ano"].unique())

plt.ylim(0, 5)
plt.yticks([0, 1, 2, 3, 4, 5])

plt.legend()

plt.grid(True)

plt.show()

"""Comparamos la evolución de las matriculaciones sin distintivo
entre turismos y motocicletas durante el período 2021-2025"""

plt.figure(figsize=(10, 6))

plt.plot(
    datos_turismos["ano"],
    datos_turismos["porcentaje_sin_distintivo"],
    marker="o",
    label="Turismos")

plt.plot(
    datos_motos["ano"],
    datos_motos["porcentaje_sin_distintivo"],
    marker="o",
    label="Motocicletas")

plt.title("Evolución de las matriculaciones sin distintivo: turismos vs motocicletas")

plt.xlabel("Año")

plt.ylabel("Porcentaje de matriculaciones (%)")

plt.xticks(comparacion_etiquetas["ano"].unique())

plt.ylim(0, 5)
plt.yticks([0, 1, 2, 3, 4, 5])

plt.legend()

plt.grid(True)

plt.show()


"""Comparamos los porcentajes de las cinco etiquetas ambientales
entre 2021 y 2025 para turismos y motocicletas"""

comparacion_2021_2025 = comparacion_etiquetas[
    comparacion_etiquetas["ano"].isin([2021, 2025])].copy()


"""Vamos a crear una tabla resumen de los porcentaje de cada etiqueta 2021 Vs 2025"""

resumen_comparativo = comparacion_2021_2025.pivot(
    index="tipo",
    columns="ano",
    values=[
        "porcentaje_cero",
        "porcentaje_b",
        "porcentaje_c",
        "porcentaje_eco",
        "porcentaje_sin_distintivo"])



"""Vamos a unir las dos primeras filas de cada columna del dataframe, para que
se vea más simplificada la tabla"""
resumen_comparativo.columns = [f"{etiqueta}_{ano}"
    for etiqueta, ano in resumen_comparativo.columns]


"""Calculamos los cambios en puntos porcentuales"""

etiquetas = ["cero", "b", "c", "eco", "sin_distintivo"]

for etiqueta in etiquetas:

    resumen_comparativo[f"cambio_{etiqueta}"] = (
        resumen_comparativo[f"porcentaje_{etiqueta}_2025"]
        - resumen_comparativo[f"porcentaje_{etiqueta}_2021"])



"""Creamos una tabla más sencilla solamente con los cambios de las etiquetas"""

cambios_etiquetas = resumen_comparativo[
    ["cambio_cero",
     "cambio_b",
     "cambio_c",
     "cambio_eco",
     "cambio_sin_distintivo"]].copy()



"""Creamos un DataFrame para comparar la evolución provincial
de turismos y motocicletas."""

columnas_provincias = [
    "provincias",
    "ano",
    "porcentaje_cero",
    "porcentaje_b",
    "porcentaje_c",
    "porcentaje_eco",
    "porcentaje_sin_distintivo"]

provincias_turismos = turismos[columnas_provincias].copy()

provincias_turismos["tipo"] = "turismos"

provincias_motos = motos[columnas_provincias].copy()

provincias_motos["tipo"] = "motocicletas"



"""Procedemos a unir los el dataframe provincias_turismos y provincias_motos
y los ordenamos por provincias, año y tipo"""
comparacion_provincias = pd.concat([provincias_turismos, provincias_motos],ignore_index=True)


comparacion_provincias = comparacion_provincias.sort_values(
    ["provincias", "ano", "tipo"]).reset_index(drop=True)




"""Creamos un dataframe para observar las diferencias entre los años 2025 y 2021"""
comparacion_provincias_2021_2025 = comparacion_provincias[
    comparacion_provincias["ano"].isin([2021, 2025])].copy()


resumen_provincias = comparacion_provincias_2021_2025.pivot(
    index=["provincias", "tipo"],
    columns="ano",
    values=[
        "porcentaje_cero",
        "porcentaje_b",
        "porcentaje_c",
        "porcentaje_eco",
        "porcentaje_sin_distintivo"])

resumen_provincias.columns = [f"{etiqueta}_{ano}"
    for etiqueta, ano in resumen_provincias.columns]

resumen_provincias = resumen_provincias.reset_index()


"""Calculamos el cambio en puntos porcentuales de cada etiqueta
entre 2021 y 2025 para cada provincia."""

etiquetas = ["cero", "b", "c", "eco", "sin_distintivo"]

for etiqueta in etiquetas:

    resumen_provincias[f"cambio_{etiqueta}"] = (
        resumen_provincias[f"porcentaje_{etiqueta}_2025"]
        - resumen_provincias[f"porcentaje_{etiqueta}_2021"])
    
cambios_provincias = resumen_provincias[
    ["provincias",
     "tipo",
     "cambio_cero",
     "cambio_b",
     "cambio_c",
     "cambio_eco",
     "cambio_sin_distintivo"]].copy()


cambios_provincias_turismos = cambios_provincias[
    cambios_provincias["tipo"] == "turismos"].copy()

cambios_provincias_motos = cambios_provincias[
    cambios_provincias["tipo"] == "motocicletas"].copy()

"""Precedemos a crear un nuevo dataframe, con los porcetanjes para poder organizarlos y 
ya empezar a tener información, para trabajar en Power BI"""


powerbi_porcentajes = comparacion_provincias.melt(
    id_vars=["provincias", "ano", "tipo"],
    value_vars=[
        "porcentaje_cero",
        "porcentaje_b",
        "porcentaje_c",
        "porcentaje_eco",
        "porcentaje_sin_distintivo"],
    var_name="etiqueta",
    value_name="porcentaje")


"""De igual manera con los datos absolutos"""
columnas_absolutas = [
    "provincias",
    "ano",
    "cero",
    "b",
    "c",
    "eco",
    "sin_distintivo"]

absolutos_turismos = turismos[columnas_absolutas].copy()
absolutos_turismos["tipo"] = "turismos"

absolutos_motos = motos[columnas_absolutas].copy()
absolutos_motos["tipo"] = "motocicletas"

comparacion_absolutos = pd.concat(
    [absolutos_turismos, absolutos_motos],
    ignore_index=True)

powerbi_absolutos = comparacion_absolutos.melt(
    id_vars=["provincias", "ano", "tipo"],
    value_vars=[
        "cero",
        "b",
        "c",
        "eco",
        "sin_distintivo"],
    var_name="etiqueta",
    value_name="matriculaciones")

"""Homogeneizamos los nombres de la columna de etiquetas"""
powerbi_porcentajes["etiqueta"] = (
    powerbi_porcentajes["etiqueta"]
    .str.replace("porcentaje_", "", regex=False))


powerbi_final = pd.merge(
    powerbi_absolutos,
    powerbi_porcentajes,
    on=["provincias", "ano", "tipo", "etiqueta"],
    how="inner")

comprobacion_porcentajes = (
    powerbi_final
    .groupby(["provincias", "ano", "tipo"])["porcentaje"]
    .sum())


"""Cambiamos y arreglamos algunos nombres de columnas para que tengan más
sentido en Power BI"""
powerbi_final["etiqueta"] = powerbi_final["etiqueta"].replace({
    "cero": "Cero",
    "b": "B",
    "c": "C",
    "eco": "ECO",
    "sin_distintivo": "Sin distintivo"})
print(comprobacion_porcentajes.head(10))

"""Exportamos la tabla de Power BI con separador y decimal compatibles con España."""
datos_powerbi = powerbi_final.copy()
(datos_powerbi).to_csv(
    OUTPUT_DIR / "datos_powerbi.csv",
    index=False,
    encoding="utf-8-sig",
    sep=";",
    decimal=",")

# Verificamos la exportación utilizando el mismo formato con el que se escribió.
datos_comprobacion = pd.read_csv(
    OUTPUT_DIR / "datos_powerbi.csv",
    encoding="utf-8-sig",
    sep=";",
    decimal=",")
print("Datos Power BI:", datos_comprobacion.shape)
print(datos_comprobacion.dtypes)


""" empezamos con el machine learning de turismos"""

X_cluster_turismos = cambios_provincias_turismos[
    [
        "cambio_cero",
        "cambio_b",
        "cambio_c",
        "cambio_eco",
        "cambio_sin_distintivo"]].copy()


print(X_cluster_turismos.shape)
print(X_cluster_turismos.head())
print(X_cluster_turismos.describe())

""" Vamos a estandarizar los datos de las variables"""
from sklearn.preprocessing import StandardScaler

scaler = StandardScaler()

X_cluster_turismos_scaled = scaler.fit_transform(
    X_cluster_turismos)


print(X_cluster_turismos_scaled.shape)
print(X_cluster_turismos_scaled[:5])


"""vamos a aplicar el metodo de kmeans clustering para poder determinar
 como se van agrupar , esto lo realizaremos a través del metodo del codo"""
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt

inercias = []

for k in range(2, 11):
    modelo_kmeans = KMeans(
        n_clusters=k,
        random_state=42,
        n_init=10)
    
    modelo_kmeans.fit(X_cluster_turismos_scaled)
    
    inercias.append(modelo_kmeans.inertia_)
    
    

plt.plot(
    range(2, 11),
    inercias,
    marker="o")

plt.xlabel("Número de clusters (k)")
plt.ylabel("Inercia")
plt.title("Método del codo - Turismos")
plt.xticks(range(2, 11))
plt.grid(True)

plt.show()

"""A través del metodo del codo no apreciamos un cambio significativo asi que 
para serciorarnos vamos a aplicar una segunda comprobacion a través de 
Silhouette Score"""
from sklearn.metrics import silhouette_score

silhouettes = []

for k in range(2, 11):
    modelo_kmeans = KMeans(
        n_clusters=k,
        random_state=42,
        n_init=10)

    clusters = modelo_kmeans.fit_predict(
        X_cluster_turismos_scaled)

    score = silhouette_score(
        X_cluster_turismos_scaled,
        clusters)

    silhouettes.append(score)
    
    
for k, score in zip(range(2, 11), silhouettes):
    print(f"k = {k}: {score:.3f}")
    
    
    
"""Con el silhouette el k mas apropiado sería 10, pero no sería adecuado
 dividir el grupo de 52 provincias en 10, y hemos observado que el resultado
 del silhouette y del metodo del codo, en el k = 5 es muy similar, así que
 aunque no es matemáticamente el indicado, analíticamente es el adecuado"""
 
 
kmeans_turismos = KMeans(
    n_clusters=5,
    random_state=42,
    n_init=10)

clusters_turismos = kmeans_turismos.fit_predict(
    X_cluster_turismos_scaled)


cambios_provincias_turismos["cluster"] = clusters_turismos

print(cambios_provincias_turismos["cluster"].value_counts().sort_index())

for cluster in sorted(cambios_provincias_turismos["cluster"].unique()):
    print(f"\nCLUSTER {cluster}")
    
    print(cambios_provincias_turismos.loc[cambios_provincias_turismos["cluster"] == cluster,"provincias"].tolist())
    
    

"""buscamos la media de cada etiqueta en cada provincia para poder hacer un 
gráfico y comprobar que el número de cluster es el adecuado""" 

perfil_clusters_turismos = (
    cambios_provincias_turismos
    .groupby("cluster")[
        ["cambio_cero",
         "cambio_b",
         "cambio_c",
         "cambio_eco",
         "cambio_sin_distintivo"]].mean().round(2))

print(perfil_clusters_turismos)




perfil_grafico = perfil_clusters_turismos.rename(
    columns={
        "cambio_cero": "Cero",
        "cambio_b": "B",
        "cambio_c": "C",
        "cambio_eco": "ECO",
        "cambio_sin_distintivo": "Sin distintivo"})

perfil_grafico.plot(
    kind="bar",
    figsize=(11, 6))

plt.axhline(0, color="black", linewidth=0.8)

plt.xlabel("Cluster")
plt.ylabel("Cambio en puntos porcentuales")
plt.title("Perfil de los clusters de turismos (2021-2025)")
plt.xticks(rotation=0)

plt.tight_layout()
plt.show()


""" empezamos con el machine learning de motocicletas"""
X_cluster_motos = cambios_provincias_motos[
    [
        "cambio_cero",
        "cambio_b",
        "cambio_c",
        "cambio_eco",
        "cambio_sin_distintivo"]].copy()

print(X_cluster_motos.shape)
print(X_cluster_motos.head())
print(X_cluster_motos.describe())


""" Vamos a estandarizar los datos de las variables"""
X_cluster_motos_scaled = scaler.fit_transform(
    X_cluster_motos)


print(X_cluster_motos_scaled.shape)
print(X_cluster_motos_scaled[:5])




"""vamos a aplicar el metodo de kmeans clustering para poder determinar
 como se van agrupar , esto lo realizaremos a través del metodo del codo"""
inercias_motos = []

for k in range(2, 11):
    modelo_kmeans = KMeans(
        n_clusters=k,
        random_state=42,
        n_init=10
    )

    modelo_kmeans.fit(X_cluster_motos_scaled)

    inercias_motos.append(modelo_kmeans.inertia_)
    
    
plt.figure(figsize=(8, 5))

plt.plot(
    range(2, 11),
    inercias_motos,
    marker="o"
)

plt.xlabel("Número de clusters (k)")
plt.ylabel("Inercia")
plt.title("Método del codo - Motocicletas")
plt.xticks(range(2, 11))
plt.grid(True)

plt.show()



"""A través del metodo del codo apreciamos un cambio grande en el K=2, pero 
también en el k=3 y K=4 significativo asi que para serciorarnos vamos a aplicar
una segunda comprobacion a través de Silhouette Score"""


silhouettes_motos = []

for k in range(2, 11):
    modelo_kmeans = KMeans(
        n_clusters=k,
        random_state=42,
        n_init=10
    )

    clusters = modelo_kmeans.fit_predict(
        X_cluster_motos_scaled
    )

    score = silhouette_score(
        X_cluster_motos_scaled,
        clusters
    )

    silhouettes_motos.append(score)

for k, score in zip(range(2, 11), silhouettes_motos):
    print(f"k = {k}: {score:.3f}")
    
    
    

"""Vamos a comprobar con sólo 2 clusters"""

kmeans_motos_2 = KMeans(
    n_clusters=2,
    random_state=42,
    n_init=10)

clusters_motos_2 = kmeans_motos_2.fit_predict(X_cluster_motos_scaled)

comprobacion_clusters_motos = cambios_provincias_motos[["provincias"]].copy()

comprobacion_clusters_motos["cluster"] = clusters_motos_2

print(comprobacion_clusters_motos["cluster"].value_counts().sort_index())


for cluster in sorted(comprobacion_clusters_motos["cluster"].unique()):
    print(f"\nCLUSTER {cluster}")

    print(comprobacion_clusters_motos.loc[comprobacion_clusters_motos["cluster"] == cluster,"provincias"].tolist())
    
#Nos damos cuenta que haciendolo sólo con 2 closters uno de ellos esta formado
#solamente por Albacete, ya que tiene un dato un poco rora en la matriculaciones
#CERO del año 2021

"""Vamos a pasar a comprobar con 4 clusters"""    


kmeans_motos_4 = KMeans(
    n_clusters=4,
    random_state=42,
    n_init=10)

clusters_motos_4 = kmeans_motos_4.fit_predict(
    X_cluster_motos_scaled)

comprobacion_clusters_motos_4 = cambios_provincias_motos[
    ["provincias"]].copy()

comprobacion_clusters_motos_4["cluster"] = clusters_motos_4

print(comprobacion_clusters_motos_4["cluster"].value_counts().sort_index())




for cluster in sorted(
    comprobacion_clusters_motos_4["cluster"].unique()):
    print(f"\nCLUSTER {cluster}")

    print(comprobacion_clusters_motos_4.loc[comprobacion_clusters_motos_4["cluster"] == cluster,"provincias"].tolist())
    
    
    

cambios_provincias_motos["cluster"] = clusters_motos_4

perfil_clusters_motos = cambios_provincias_motos.groupby(
    "cluster")[["cambio_cero",
     "cambio_b",
     "cambio_c",
     "cambio_eco",
     "cambio_sin_distintivo"]].mean()

print(perfil_clusters_motos.round(2))




"""Vamos a pasar a comprobar con 3 clusters"""
kmeans_motos_3 = KMeans(
    n_clusters=3,
    random_state=42,
    n_init=10)

clusters_motos_3 = kmeans_motos_3.fit_predict(
    X_cluster_motos_scaled)

comprobacion_clusters_motos_3 = cambios_provincias_motos[
    ["provincias"]].copy()

comprobacion_clusters_motos_3["cluster"] = clusters_motos_3

print(
    comprobacion_clusters_motos_3["cluster"]
    .value_counts()
    .sort_index())


for cluster in sorted(
    comprobacion_clusters_motos_3["cluster"].unique()):
    print(f"\nCLUSTER {cluster}")

    print(
        comprobacion_clusters_motos_3.loc[
            comprobacion_clusters_motos_3["cluster"] == cluster,
            "provincias"].tolist())
    
    
cambios_provincias_motos["cluster"] = clusters_motos_3

perfil_clusters_motos_3 = cambios_provincias_motos.groupby(
    "cluster")[
    ["cambio_cero",
     "cambio_b",
     "cambio_c",
     "cambio_eco",
     "cambio_sin_distintivo"]].mean()

print(perfil_clusters_motos_3.round(2))


"""Ahora procedemos a hacer un grafico de la evolución de motocicletas para 3 clusters"""


perfil_grafico_motos = perfil_clusters_motos_3.rename(
    columns={
        "cambio_cero": "Cero",
        "cambio_b": "B",
        "cambio_c": "C",
        "cambio_eco": "ECO",
        "cambio_sin_distintivo": "Sin distintivo"})

perfil_grafico_motos.plot(
    kind="bar",
    figsize=(11, 6))

plt.axhline(
    0,
    color="black",
    linewidth=0.8)

plt.xlabel("Cluster")
plt.ylabel("Cambio en puntos porcentuales")
plt.title("Perfil de los clusters de motocicletas (2021-2025)")
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()








""" Procedemos a preparar el dataframe para hacer una regresión lineal, 
primero del apartado de turismos en números absolutos"""


from sklearn.linear_model import LinearRegression

X_turismos = turismos_por_ano[["ano"]]

y_turismos = turismos_por_ano["total_turismos"]

print(X_turismos)
print(y_turismos)




modelo_turismos = LinearRegression()

modelo_turismos.fit( X_turismos,y_turismos)

pendiente_turismos = modelo_turismos.coef_[0]

intercepto_turismos = modelo_turismos.intercept_

r2_turismos = modelo_turismos.score(
    X_turismos,
    y_turismos
)

print("Pendiente:", pendiente_turismos)
print("Intercepto:", intercepto_turismos)
print("R²:", r2_turismos)



predicciones_turismos = modelo_turismos.predict(X_turismos)

print(predicciones_turismos)



plt.figure(figsize=(9, 6))

plt.scatter(
    X_turismos["ano"],
    y_turismos,
    label="Datos reales")

plt.plot(
    X_turismos["ano"],
    predicciones_turismos,
    label="Regresión lineal")

plt.xlabel("Año")
plt.ylabel("Matriculaciones")
plt.title("Regresión lineal de matriculaciones de turismos (2021-2025)")
plt.xticks(X_turismos["ano"])
plt.legend()
plt.tight_layout()
plt.show()


""" Ahora vamos a empezar la regresión lineal de las etiquetas
medioambientales de turismos"""


etiquetas = [
    "cero",
    "b",
    "c",
    "eco",
    "sin_distintivo"]

modelos_etiquetas_turismos = {}
resultados_regresion_turismos = []

for etiqueta in etiquetas:

    y = etiquetas_por_ano[etiqueta]

    modelo = LinearRegression()
    modelo.fit(X_turismos, y)

    pendiente = modelo.coef_[0]
    r2 = modelo.score(X_turismos, y)

    modelos_etiquetas_turismos[etiqueta] = modelo

    resultados_regresion_turismos.append(
        {"etiqueta": etiqueta,
         "pendiente": pendiente,
         "r2": r2})
    
resultados_regresion_turismos = pd.DataFrame(
    resultados_regresion_turismos)

print(resultados_regresion_turismos.round(3))


"""Vamos a hacer un gráfico de líneas para ver la regresión lineal de las cinco 
etiquetas mediambientales """
plt.figure(figsize=(11, 7))

for etiqueta in etiquetas:

    y_real = etiquetas_por_ano[etiqueta]

    modelo = modelos_etiquetas_turismos[etiqueta]

    y_pred = modelo.predict(X_turismos)

    plt.scatter(X_turismos["ano"],y_real)

    plt.plot(X_turismos["ano"],y_pred,label=etiqueta)

plt.xlabel("Año")
plt.ylabel("Matriculaciones")
plt.title("Regresión lineal por etiqueta medioambiental - Turismos (2021-2025)")

plt.xticks(X_turismos["ano"])
plt.legend()
plt.tight_layout()
plt.show()




""" Ahora vamos a hacer lo mismo, preparar el dataframe para hacer una regresión
lineal, pero esta vez con los valores porcentuales del apartado de turismos"""

etiquetas_porcentaje = [
    "porcentaje_cero",
    "porcentaje_b",
    "porcentaje_c",
    "porcentaje_eco",
    "porcentaje_sin_distintivo"]

modelos_porcentaje_turismos = {}
resultados_porcentaje_turismos = []

for etiqueta in etiquetas_porcentaje:

    y = etiquetas_por_ano[etiqueta]

    modelo = LinearRegression()
    modelo.fit(X_turismos, y)

    pendiente = modelo.coef_[0]
    r2 = modelo.score(X_turismos, y)

    modelos_porcentaje_turismos[etiqueta] = modelo

    resultados_porcentaje_turismos.append(
        {"etiqueta": etiqueta,
         "pendiente": pendiente,
         "r2": r2})
    
resultados_porcentaje_turismos = pd.DataFrame(resultados_porcentaje_turismos)

print(resultados_porcentaje_turismos.round(3))


""" Vamos a verlo en un grafico de lineas"""

plt.figure(figsize=(11, 7))

for etiqueta in etiquetas_porcentaje:

    y_real = etiquetas_por_ano[etiqueta]

    modelo = modelos_porcentaje_turismos[etiqueta]

    y_pred = modelo.predict(X_turismos)

    nombre = etiqueta.replace("porcentaje_", "")

    plt.scatter(
        X_turismos["ano"],
        y_real)

    plt.plot(X_turismos["ano"],y_pred,label=nombre)

plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")

plt.title("Regresión lineal del peso de las etiquetas - Turismos (2021-2025)")

plt.xticks(X_turismos["ano"])
plt.legend()
plt.tight_layout()
plt.show()




""" Procedemos a hacer una estimación de la evulución para los siguientes 5 
años de los valores absolutos de matriculaciones"""

anos_futuros = pd.DataFrame({
    "ano": [2026, 2027, 2028, 2029, 2030]})

estimaciones_2030_turismos = []

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_turismos[etiqueta]

    predicciones = modelo.predict(
        anos_futuros)

    estimaciones_2030_turismos.append({"etiqueta": etiqueta,"estimacion_2030": predicciones[-1]})

estimaciones_2030_turismos = pd.DataFrame(
    estimaciones_2030_turismos)

print(estimaciones_2030_turismos.round(0))




proyecciones_absolutas_turismos = anos_futuros.copy()

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_turismos[etiqueta]

    proyecciones_absolutas_turismos[etiqueta] = modelo.predict(
        anos_futuros)

print(proyecciones_absolutas_turismos.round(0))

anos_proyeccion = pd.DataFrame({"ano": [2025, 2026, 2027, 2028, 2029, 2030]})
proyeccion_absoluta_grafico = anos_proyeccion.copy()

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_turismos[etiqueta]

    proyeccion_absoluta_grafico[etiqueta] = modelo.predict(
        anos_proyeccion)

print(proyeccion_absoluta_grafico.round(0))


plt.figure(figsize=(12, 7))

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_turismos[etiqueta]

    # Datos reales 2021-2025
    plt.scatter(
        X_turismos["ano"],
        etiquetas_por_ano[etiqueta])

    # Recta ajustada 2021-2025
    prediccion_historica = modelo.predict(
        X_turismos)

    plt.plot(
        X_turismos["ano"],
        prediccion_historica,
        label=etiqueta)

    # Proyección 2025-2030
    plt.plot(
        anos_proyeccion["ano"],
        proyeccion_absoluta_grafico[etiqueta],
        linestyle="--")

plt.axvline(
    x=2025,
    color="black",
    linestyle=":",
    linewidth=1)

plt.xlabel("Año")
plt.ylabel("Número de matriculaciones")

plt.title("Evolución y proyección lineal por etiqueta - Turismos (2021-2030)")

plt.xticks(range(2021, 2031))
plt.legend()
plt.tight_layout()
plt.show()




""" Comenzamos a hacer la regresion lineal para motocicletas con valores 
absolutos de matriculaciones"""

X_motos = motos_por_ano[["ano"]]
y_motos = motos_por_ano["total_motos"]

modelo_total_motos = LinearRegression()

modelo_total_motos.fit(
    X_motos,
    y_motos)

prediccion_total_motos = modelo_total_motos.predict(X_motos)

print("Pendiente:",modelo_total_motos.coef_[0])

print("Intercepto:",modelo_total_motos.intercept_)

print("R²:",modelo_total_motos.score(X_motos,y_motos))

print(pd.DataFrame({
        "ano": motos_por_ano["ano"],
        "real": y_motos,
        "estimado": prediccion_total_motos}).round(0))



plt.figure(figsize=(10, 6))

# Datos reales
plt.scatter(
    X_motos["ano"],
    y_motos,
    label="Datos reales")

# Recta de regresión
plt.plot(
    X_motos["ano"],
    prediccion_total_motos,
    label="Regresión lineal")

plt.xlabel("Año")
plt.ylabel("Número de matriculaciones")

plt.title("Regresión lineal del total de motocicletas (2021-2025)")

plt.xticks(range(2021, 2026))
plt.legend()
plt.tight_layout()
plt.show()

"""Ahora vamos a hacerlo con todas las etiquetas"""

modelos_etiquetas_motos = {}
resultados_motos = []

for etiqueta in etiquetas:

    y = etiquetas_motos[etiqueta]

    modelo = LinearRegression()

    modelo.fit(
        X_motos,
        y)

    modelos_etiquetas_motos[etiqueta] = modelo

    resultados_motos.append({
        "etiqueta": etiqueta,
        "pendiente": modelo.coef_[0],
        "R2": modelo.score(X_motos, y)})

resultados_regresion_motos = pd.DataFrame(resultados_motos)

print(resultados_regresion_motos.round(3))




plt.figure(figsize=(12, 7))

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_motos[etiqueta]

    y = etiquetas_motos[etiqueta]

    # Datos reales
    plt.scatter(
        X_motos["ano"],
        y)

    # Recta de regresión
    prediccion = modelo.predict(
        X_motos)

    plt.plot(
        X_motos["ano"],
        prediccion,
        label=etiqueta)

plt.xlabel("Año")
plt.ylabel("Número de matriculaciones")

plt.title("Regresión lineal por etiqueta - Motocicletas (2021-2025)")

plt.xticks(range(2021, 2026))
plt.legend()
plt.tight_layout()
plt.show()



""" Comenzamos a hacer la regresion lineal para motocicletas con valores 
porcentuales de matriculaciones"""

etiquetas_porcentaje = [
    "porcentaje_cero",
    "porcentaje_b",
    "porcentaje_c",
    "porcentaje_eco",
    "porcentaje_sin_distintivo"]

modelos_porcentaje_motos = {}
resultados_porcentaje_motos = []

for etiqueta in etiquetas_porcentaje:

    y = etiquetas_motos[etiqueta]

    modelo = LinearRegression()

    modelo.fit(
        X_motos,
        y)

    modelos_porcentaje_motos[etiqueta] = modelo

    resultados_porcentaje_motos.append({
        "etiqueta": etiqueta,
        "pendiente": modelo.coef_[0],
        "R2": modelo.score(X_motos, y)})

resultados_regresion_porcentaje_motos = pd.DataFrame(
    resultados_porcentaje_motos)

print(resultados_regresion_porcentaje_motos.round(3))



plt.figure(figsize=(12, 7))

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_motos[etiqueta]

    y = etiquetas_motos[etiqueta]

    # Datos reales
    plt.scatter(
        X_motos["ano"],
        y)

    # Recta de regresión
    prediccion = modelo.predict(
        X_motos)

    plt.plot(
        X_motos["ano"],
        prediccion,
        label=etiqueta.replace("porcentaje_", ""))

plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")

plt.title("Regresión lineal porcentual por etiqueta - Motocicletas (2021-2025)")

plt.xticks(range(2021, 2026))
plt.legend()
plt.tight_layout()
plt.show()



""" Vamos a hacer las predicciones a futuro de los valores absolutos de motocicletas"""
proyecciones_absolutas_motos = anos_futuros.copy()

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_motos[etiqueta]

    proyecciones_absolutas_motos[etiqueta] = modelo.predict(anos_futuros)

print(proyecciones_absolutas_motos.round(0))


""" Vamos a hacer las predicciones a futuro de los valores porcentuales de motocicletas"""


proyecciones_porcentaje_motos = anos_futuros.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_motos[etiqueta]

    proyecciones_porcentaje_motos[etiqueta] = modelo.predict(anos_futuros)

columnas_porcentaje = [
    "porcentaje_cero",
    "porcentaje_b",
    "porcentaje_c",
    "porcentaje_eco",
    "porcentaje_sin_distintivo"]

proyecciones_porcentaje_motos["suma"] = (proyecciones_porcentaje_motos[columnas_porcentaje].sum(axis=1))

print(proyecciones_porcentaje_motos.round(2))



"""Veamos la proyeccion a 5 años en valores porcentuales"""

proyeccion_porcentaje_grafico_motos = anos_proyeccion.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_motos[etiqueta]

    proyeccion_porcentaje_grafico_motos[etiqueta] = modelo.predict(
        anos_proyeccion)

print(proyeccion_porcentaje_grafico_motos.round(2))




plt.figure(figsize=(12, 7))

for etiqueta in etiquetas_porcentaje:

    nombre = etiqueta.replace(
        "porcentaje_", "")

    modelo = modelos_porcentaje_motos[etiqueta]

    # Datos reales 2021-2025
    plt.scatter(
        X_motos["ano"],
        etiquetas_motos[etiqueta])

    # Recta ajustada 2021-2025
    prediccion_historica = modelo.predict(
        X_motos)

    plt.plot(
        X_motos["ano"],
        prediccion_historica,
        label=nombre)

    # Proyección 2025-2030
    plt.plot(
        anos_proyeccion["ano"],
        proyeccion_porcentaje_grafico_motos[etiqueta],
        linestyle="--")

plt.axvline(
    x=2025,
    color="black",
    linestyle=":",
    linewidth=1)

plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")

plt.title("Evolución y proyección lineal por etiqueta - Motocicletas (2021-2030)")

plt.xticks(range(2021, 2031))
plt.legend()
plt.tight_layout()
plt.show()


"""Veamos la proyeccion a 5 años en total matriculaciones"""

proyeccion_absoluta_grafico_motos = anos_proyeccion.copy()

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_motos[etiqueta]

    proyeccion_absoluta_grafico_motos[etiqueta] = modelo.predict(anos_proyeccion)

print(proyeccion_absoluta_grafico_motos.round(0))

plt.figure(figsize=(12, 7))

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_motos[etiqueta]

    # Datos reales 2021-2025
    plt.scatter(
        X_motos["ano"],
        etiquetas_motos[etiqueta])

    # Recta ajustada 2021-2025
    prediccion_historica = modelo.predict(
        X_motos)

    plt.plot(
        X_motos["ano"],
        prediccion_historica,
        label=etiqueta)

    # Proyección 2025-2030
    plt.plot(
        anos_proyeccion["ano"],
        proyeccion_absoluta_grafico_motos[etiqueta],
        linestyle="--")

plt.axvline(
    x=2025,
    color="black",
    linestyle=":",
    linewidth=1)

plt.xlabel("Año")
plt.ylabel("Número de matriculaciones")

plt.title("Evolución y proyección lineal por etiqueta - Motocicletas (2021-2030)")

plt.xticks(range(2021, 2031))
plt.legend()
plt.tight_layout()
plt.show()

























""" Procedemos a hacer una estimación de la evulución para los siguientes 5 
años de los porcentajes de matriculaciones"""
estimaciones_porcentaje_2030 = []

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_turismos[etiqueta]

    prediccion = modelo.predict(pd.DataFrame({"ano": [2030]}))[0]

    estimaciones_porcentaje_2030.append({"etiqueta": etiqueta,"porcentaje_2030": prediccion})

estimaciones_porcentaje_2030 = pd.DataFrame(
    estimaciones_porcentaje_2030)

print(estimaciones_porcentaje_2030.round(2))

print("Suma:",estimaciones_porcentaje_2030["porcentaje_2030"].sum())




proyecciones_porcentaje_turismos = anos_futuros.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_turismos[etiqueta]

    proyecciones_porcentaje_turismos[etiqueta] = modelo.predict(
        anos_futuros)

print(proyecciones_porcentaje_turismos.round(2))


columnas_porcentaje = [
    "porcentaje_cero",
    "porcentaje_b",
    "porcentaje_c",
    "porcentaje_eco",
    "porcentaje_sin_distintivo"]

proyecciones_porcentaje_turismos["suma"] = (proyecciones_porcentaje_turismos[columnas_porcentaje].sum(axis=1))

print(proyecciones_porcentaje_turismos.round(2))




anos_proyeccion = pd.DataFrame({"ano": [2025, 2026, 2027, 2028, 2029, 2030]})

proyeccion_grafico_turismos = anos_proyeccion.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_turismos[etiqueta]

    proyeccion_grafico_turismos[etiqueta] = modelo.predict(anos_proyeccion)

print(proyeccion_grafico_turismos.round(2))







plt.figure(figsize=(12, 7))

for etiqueta in etiquetas_porcentaje:

    nombre = etiqueta.replace("porcentaje_", "")

    modelo = modelos_porcentaje_turismos[etiqueta]

    # Datos reales 2021-2025
    plt.scatter(X_turismos["ano"],
        etiquetas_por_ano[etiqueta])

    # Recta ajustada 2021-2025
    prediccion_historica = modelo.predict(
        X_turismos)

    plt.plot(
        X_turismos["ano"],
        prediccion_historica,
        label=nombre)

    # Proyección 2025-2030
    plt.plot(
        anos_proyeccion["ano"],
        proyeccion_grafico_turismos[etiqueta],
        linestyle="--")

plt.axvline(
    x=2025,
    color="black",
    linestyle=":",
    linewidth=1)

plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")

plt.title("Evolución y proyección lineal por etiqueta - Turismos (2021-2030)")

plt.xticks(range(2021, 2031))
plt.legend()
plt.tight_layout()
plt.show()




print(cambios_provincias_turismos.columns)
print(cambios_provincias_motos.columns)




"""Vamos a comenzar a crear archivos csv para poder trabajarlos en Power BI"""

"""Primero uno para el ML K-means clustering"""


clustering_powerbi = pd.concat([cambios_provincias_turismos,cambios_provincias_motos],ignore_index=True)

print(clustering_powerbi.shape)
print(clustering_powerbi.head())
print(clustering_powerbi["tipo"].value_counts())
print(clustering_powerbi.groupby(["tipo", "cluster"]).size())

clustering_powerbi.to_csv(OUTPUT_DIR / "clustering_powerbi.csv",index=False,encoding="utf-8-sig",sep=";",decimal=",")


"""Ahora vamos a crear el csv de ML para las regresiones lineales"""


etiquetas = ["cero", "b", "c", "eco", "sin_distintivo"]

proyecciones_lista = []

for tipo, df_abs, df_pct in [
    ("turismos",proyecciones_absolutas_turismos,proyecciones_porcentaje_turismos
    ),("motocicletas",proyecciones_absolutas_motos,proyecciones_porcentaje_motos)]:

    for etiqueta in etiquetas:

        temporal = pd.DataFrame({
            "ano": df_abs["ano"],
            "tipo": tipo,
            "etiqueta": etiqueta,
            "matriculaciones_proyectadas": df_abs[etiqueta],
            "porcentaje_proyectado": df_pct["porcentaje_" + etiqueta]})

        proyecciones_lista.append(temporal)


proyecciones_powerbi = pd.concat(proyecciones_lista,ignore_index=True)

# Ordenamos para facilitar su lectura
proyecciones_powerbi = proyecciones_powerbi.sort_values(["tipo", "ano", "etiqueta"]).reset_index(drop=True)

print(proyecciones_powerbi.shape)
print(proyecciones_powerbi.head(15))

print(proyecciones_powerbi.groupby(["tipo", "ano"]).size())

proyecciones_powerbi.to_csv(OUTPUT_DIR / "proyecciones_powerbi.csv",index=False,encoding="utf-8-sig",sep=";",decimal=",")






""" Nota Aclaratoria

Inicialmente se realizó una proyección hasta 2030. Sin embargo,
al disponer únicamente de datos históricos correspondientes al
periodo 2021-2025, se considera más prudente reducir el horizonte
de proyección a los años 2026 y 2027.

Los modelos de regresión siguen entrenados con los datos reales
de 2021-2025.
"""

# Nuevo horizonte de proyección
anos_futuros_2027 = pd.DataFrame({
    "ano": [2026, 2027]})

print(anos_futuros_2027)


# Proyección revisada de turismos para 2026-2027

proyeccion_turismos_2027 = modelo_turismos.predict(anos_futuros_2027[["ano"]])

proyeccion_turismos_2027 = pd.DataFrame({"ano": anos_futuros_2027["ano"],
    "total_turismos_proyectado": proyeccion_turismos_2027})

print(proyeccion_turismos_2027)




# Proyección revisada por etiqueta de turismos para 2026-2027

proyecciones_etiquetas_turismos_2027 = anos_futuros_2027.copy()

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_turismos[etiqueta]

    proyecciones_etiquetas_turismos_2027[etiqueta] = modelo.predict(anos_futuros_2027[["ano"]])

print(proyecciones_etiquetas_turismos_2027.round(0))




# Proyección revisada de los porcentajes por etiqueta
# de turismos para 2026-2027

proyecciones_porcentaje_turismos_2027 = anos_futuros_2027.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_turismos[etiqueta]

    proyecciones_porcentaje_turismos_2027[etiqueta] = modelo.predict(anos_futuros_2027[["ano"]])

# Comprobamos que los porcentajes suman 100 %
proyecciones_porcentaje_turismos_2027["suma"] = (
    proyecciones_porcentaje_turismos_2027[etiquetas_porcentaje].sum(axis=1))

print(proyecciones_porcentaje_turismos_2027.round(2))



""" Gráfico revisado de la proyección porcentual de turismos
para el periodo 2026-2027 """

# Incluimos 2025 para conectar la recta de regresión
# con la proyección 2026-2027
anos_proyeccion_2027 = pd.DataFrame({"ano": [2025, 2026, 2027]})

proyeccion_grafico_turismos_2027 = anos_proyeccion_2027.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_turismos[etiqueta]

    proyeccion_grafico_turismos_2027[etiqueta] = modelo.predict(
        anos_proyeccion_2027[["ano"]])

print(proyeccion_grafico_turismos_2027.round(2))



""" Gráfico final de la proyección porcentual de turismos
para el periodo 2026-2027 """

plt.figure(figsize=(12, 7))

for etiqueta in etiquetas_porcentaje:

    nombre = etiqueta.replace("porcentaje_", "")

    modelo = modelos_porcentaje_turismos[etiqueta]

    # Datos reales 2021-2025
    plt.scatter(
        X_turismos["ano"],
        etiquetas_por_ano[etiqueta])

    # Recta de regresión ajustada con 2021-2025
    prediccion_historica = modelo.predict(
        X_turismos)

    plt.plot(
        X_turismos["ano"],
        prediccion_historica,
        label=nombre)

    # Proyección de corto plazo 2025-2027
    plt.plot(
        anos_proyeccion_2027["ano"],
        proyeccion_grafico_turismos_2027[etiqueta],
        linestyle="--")

# Separación entre periodo histórico y proyectado
plt.axvline(
    x=2025,
    color="black",
    linestyle=":",
    linewidth=1)

plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")

plt.title(
    "Evolución y proyección lineal por etiqueta - Turismos (2021-2027)")





"""Ahora vamos a realizar el mismo proceso, pero con motocicletas"""
# Proyección revisada de motocicletas para 2026-2027

proyeccion_motos_2027 = modelo_total_motos.predict(
    anos_futuros_2027[["ano"]])

proyeccion_motos_2027 = pd.DataFrame({
    "ano": anos_futuros_2027["ano"],
    "total_motos_proyectado": proyeccion_motos_2027})

print(proyeccion_motos_2027)


# Proyección revisada por etiqueta de motocicletas para 2026-2027

proyecciones_etiquetas_motos_2027 = anos_futuros_2027.copy()

for etiqueta in etiquetas:

    modelo = modelos_etiquetas_motos[etiqueta]

    proyecciones_etiquetas_motos_2027[etiqueta] = modelo.predict(
        anos_futuros_2027[["ano"]])

print(proyecciones_etiquetas_motos_2027.round(0))



# Proyección revisada de los porcentajes por etiqueta
# de motocicletas para 2026-2027

proyecciones_porcentaje_motos_2027 = anos_futuros_2027.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_motos[etiqueta]

    proyecciones_porcentaje_motos_2027[etiqueta] = modelo.predict(
        anos_futuros_2027[["ano"]])

# Comprobamos que los porcentajes suman 100 %
proyecciones_porcentaje_motos_2027["suma"] = (
    proyecciones_porcentaje_motos_2027[etiquetas_porcentaje].sum(axis=1))

print(proyecciones_porcentaje_motos_2027.round(2))

""" Preparamos los datos para el gráfico revisado
de motocicletas 2025-2027 """

proyeccion_grafico_motos_2027 = anos_proyeccion_2027.copy()

for etiqueta in etiquetas_porcentaje:

    modelo = modelos_porcentaje_motos[etiqueta]

    proyeccion_grafico_motos_2027[etiqueta] = modelo.predict(
        anos_proyeccion_2027[["ano"]])

print(proyeccion_grafico_motos_2027.round(2))


""" Gráfico final de la proyección porcentual de motocicletas
para el periodo 2026-2027 """

plt.figure(figsize=(12, 7))

for etiqueta in etiquetas_porcentaje:

    nombre = etiqueta.replace("porcentaje_", "")

    modelo = modelos_porcentaje_motos[etiqueta]

    # Datos reales 2021-2025
    plt.scatter(
        X_motos["ano"],
        etiquetas_motos[etiqueta])

    # Recta de regresión ajustada con 2021-2025
    prediccion_historica = modelo.predict(
        X_motos)

    plt.plot(
        X_motos["ano"],
        prediccion_historica,
        label=nombre)

    # Proyección de corto plazo 2025-2027
    plt.plot(
        anos_proyeccion_2027["ano"],
        proyeccion_grafico_motos_2027[etiqueta],
        linestyle="--")

# Separación entre periodo histórico y proyectado
plt.axvline(
    x=2025,
    color="black",
    linestyle=":",
    linewidth=1)

plt.xlabel("Año")
plt.ylabel("Porcentaje de matriculaciones (%)")

plt.title("Evolución y proyección lineal por etiqueta - Motocicletas (2021-2027)")

plt.xticks(range(2021, 2028))

plt.legend()
plt.tight_layout()
plt.show()
