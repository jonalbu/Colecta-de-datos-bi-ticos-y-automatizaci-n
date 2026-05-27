# -*- coding: utf-8 -*-
"""
Plantilla de ingesta de base de datos de monitoreo ictiológico
================================================================

Procesa el XLS exportado de un formulario XLSForm de KoboToolbox y genera
las hojas de salida (Individuos/Peces por zona y Hábitat) según unas
plantillas de destino.

Esta es una versión ESTÁNDAR y anónima: no contiene nombres de proyecto,
códigos de estación, especies ni rutas reales. Todo lo específico de su
estudio se configura en la sección "CONFIGURACIÓN" más abajo, siguiendo las
instrucciones marcadas con  # >>> COMPLETAR.

Requisitos:
    pip install pandas openpyxl requests

Uso:
    python ingestar_BD_template.py
"""

import os
import time

import pandas as pd
import requests


# ===========================================================================
# CONFIGURACIÓN  — Edite SOLO esta sección.
# ===========================================================================

# --- 1. Archivo de entrada (exportado de KoboToolbox) ----------------------
# >>> COMPLETAR: nombre del archivo .xlsx exportado de Kobo y los nombres de
#     sus hojas. Una exportación de KoboToolbox crea una hoja por cada nivel
#     del formulario (el nivel principal y cada `begin_repeat`).
PATH_ENTRADA = "datos_export.xlsx"

HOJA_PRINCIPAL = "<NOMBRE_HOJA_PRINCIPAL>"   # nivel estación/muestreo
HOJA_MEDICIONES = "mediciones"               # repeat de mediciones de flujo
HOJA_INDIVIDUOS = "<NOMBRE_HOJA_INDIVIDUOS>"  # repeat de individuos/peces

# --- 2. Token de la API de IUCN (Red List API v4) --------------------------
# >>> COMPLETAR: solicite un token gratuito en https://api.iucnredlist.org/
#     ("Generate a token") o en https://api.iucnredlist.org/users/sign_up
#     y péguelo aquí. El token llega por correo, no de inmediato.
#     Si se deja el valor por defecto, la categoría IUCN quedará como 'NA'.
IUCN_TOKEN = "TU_TOKEN_AQUI"
IUCN_API_BASE = "https://api.iucnredlist.org/api/v4"

# --- 3. Sistemas / zonas y sus códigos de estación -------------------------
# >>> COMPLETAR: asocie el nombre de cada sistema (o zona) con la lista de
#     códigos de estación que le pertenecen. Estos códigos son los valores de
#     la columna de identificación de estación (p. ej. "ID_Sitio_0").
#     Ejemplo:
#         SISTEMAS_ACUATICOS = {
#             "Sistema A": ["A-1", "A-2", "A-3"],
#             "Sistema B": ["B-1", "B-2"],
#         }
SISTEMAS_ACUATICOS = {
    # "Nombre del sistema": ["CODIGO-1", "CODIGO-2"],
}
SISTEMA_DESCONOCIDO = "Desconocido"  # valor cuando el código no está en la tabla

# --- 4. Dietas por especie -------------------------------------------------
# >>> COMPLETAR: nombre científico -> categoría trófica.
#     Las especies que no estén aquí quedarán como 'Desconocida'.
#     Ejemplo:
#         DIETA = {"Genus species": "Omnívoro-insectívoro"}
DIETA = {
    # "Genus species": "Categoría trófica",
}

# --- 5. Correcciones manuales de familia por género ------------------------
# >>> COMPLETAR (opcional): cuando GBIF devuelve una familia que se desea
#     forzar. Ejemplo: {"Genero": "Familia_correcta"}
CORRECCIONES_FAMILIA = {
    # "Genero": "Familia",
}

# --- 6. Exportación de la hoja de Individuos por sistema --------------------
# >>> COMPLETAR: para cada sistema que se desea exportar, indique la plantilla
#     de destino (de la que se leen los encabezados) y el archivo de salida.
#     La clave debe coincidir con un nombre de SISTEMAS_ACUATICOS.
#     Ejemplo:
#         EXPORTES_INDIVIDUOS = {
#             "Sistema A": {"plantilla": "plantilla_A.xlsx",
#                           "salida":    "BD_Actualizada_A.xlsx"},
#         }
EXPORTES_INDIVIDUOS = {
    # "Nombre del sistema": {"plantilla": "plantilla.xlsx", "salida": "salida.xlsx"},
}
HOJA_PLANTILLA_INDIVIDUOS = "Sheet1"   # hoja dentro de las plantillas de destino

# --- 7. Exportación de la hoja de Hábitat ----------------------------------
PLANTILLA_HABITAT = "plantilla_habitat.xlsx"
SALIDA_HABITAT = "habitat_exportado.xlsx"
HOJA_HABITAT = "<NOMBRE_HOJA_HABITAT>"  # >>> COMPLETAR: hoja en la plantilla (colocar el nombre de la hoja que tiene los encabezados de la plantilla de hábitat)
# ver el Plantilla_habitat.xlsx para saber qué columnas se necesitan y cómo se llaman los encabezados

# ===========================================================================
# CONSTANTES / MAPEOS  — derivados del XLSForm, normalmente no requieren cambios.
# ===========================================================================

# Columnas a eliminar en cada hoja (metadatos de KoboToolbox y campos auxiliares).
# Ajuste según las columnas reales de su exportación.
COLS_DROP_PRINCIPAL = [
    'Coordenadas', 'fin', '_uuid', 'meta/rootUuid', '_index',
    'aparejo_actual', 'area',
]
COLS_DROP_MEDICIONES = [
    'Ver_vel', '_submission__id', '_submission__uuid',
    '_submission__submission_time', '_submission__validation_status',
    '_submission__notes', '_submission__status', '_submission__submitted_by',
    '_submission___version__', '_submission__tags', '_submission_meta/rootUuid',
]
COLS_DROP_INDIVIDUOS = [
    'ID_Registro', 'nuevo_aparejo', '_parent_table_name', '_parent_index',
    '_submission__id', '_submission__uuid', '_submission__submission_time',
    '_submission__validation_status', '_submission__notes', '_submission__status',
    '_submission__submitted_by', '_submission___version__', '_submission__tags',
    '_submission_meta/rootUuid',
]

# Indicadores de morfoespecies (sin epíteto válido para consultar APIs).
_SIN_ESPECIE = {'sp.', 'sp', 'gr.', 'gr', 'cf.', 'cf', 'aff.', 'aff'}

# Renombrado de columnas de la hoja de Individuos -> encabezados de la BD final.
# Ajuste las claves a los nombres de campo de su formulario y los valores a los
# encabezados de su plantilla de destino.
RENOMBRE_INDIVIDUOS = {
    'codigo_individuo': 'ID',
    'fecha_colecta': 'Fecha',
    'dia': 'Dia',
    'año': 'Año',
    'mes': 'Mes',
    'sistema_acuatico': 'Complejo hídrico',
    'Sitio_colecta': 'nombre estación',
    'ID_Sitio_2': 'Codigo (estacion)',
    'aparejo_pez': 'Aparejo',
    'orden': 'Orden',
    'familia': 'Familia',
    'genero': 'Genero',
    'Especie': 'Taxon sin Autor',
    'dieta': 'Dieta',
    'Talla': 'LST (mm)',
    'Peso': 'Wt (g)',
    'Peso_evicerado': 'We (g)',
    'Peso_tracto': 'Wtra (g)',
    'Peso_gonada': 'Wg (g)',
    'Sexo': 'SEX',
    'Estadio': 'EST.',
    'contados': 'Nº Liberados',
    'codigo_unico': 'Codigo Esp',
}

# Renombrado de columnas de la hoja Principal -> encabezados de Hábitat.
RENOMBRE_PRINCIPAL = {
    'ID_Sitio_0': 'CODIGO(NUEVO)',
    'sitio': 'Estación',
    'fecha': 'Fecha',
    'dia': 'Día',
    'año': 'Año',
    'mes': 'Mes',
    'ph': 'pH',
    'Ancho_cauce': 'ancho prom efec(mojado)',
    'caudal': 'Caudal Máximo',
    'Oxigeno': 'Oxigeno disuelto',
    'Saturación_oxigeno': '% de Saturación de Oxigeno',
    'Sechhi': 'Transparencia',
    'prof_promedio': 'Profundidad (mean)',
    'vel_promedio': 'Velocidad del agua (media)',
    'cv_velocidad': 'Velocidad del agua (coeficiente de Variación)',
    'cv_profundidad': 'Profundidad (Coeficiente de Variación)',
}

# Variables de hábitat: nombre en la plantilla -> columna en la hoja Principal.
VARIABLES_HABITAT = {
    'Caudal Máximo': 'Caudal Máximo',
    'Profundidad (mean)': 'Profundidad (mean)',
    'Velocidad del agua (media)': 'Velocidad del agua (media)',
    'Profundidad (Coeficiente de Variación)': 'Profundidad (Coeficiente de Variación)',
    'Velocidad del agua (coeficiente de Variación)': 'Velocidad del agua (coeficiente de Variación)',
    'Temperatura': 'Temperatura',
    'Conductividad': 'Conductividad',
    'pH': 'pH',
    'Oxigeno disuelto': 'Oxigeno disuelto',
    '% de Saturación de Oxigeno': '% de Saturación de Oxigeno',
    'Transparencia': 'Transparencia',
    'ancho prom efec(mojado)': 'ancho prom efec(mojado)',
}

MESES_ES = {
    '01': 'Enero', '1': 'Enero',
    '02': 'Febrero', '2': 'Febrero',
    '03': 'Marzo', '3': 'Marzo',
    '04': 'Abril', '4': 'Abril',
    '05': 'Mayo', '5': 'Mayo',
    '06': 'Junio', '6': 'Junio',
    '07': 'Julio', '7': 'Julio',
    '08': 'Agosto', '8': 'Agosto',
    '09': 'Septiembre', '9': 'Septiembre',
    '10': 'Octubre',
    '11': 'Noviembre',
    '12': 'Diciembre',
}


# ===========================================================================
# 1. Carga y limpieza
# ===========================================================================
def cargar_datos():
    """Carga las tres hojas del Excel exportado de KoboToolbox."""
    df_principal = pd.read_excel(PATH_ENTRADA, sheet_name=HOJA_PRINCIPAL)
    df_mediciones = pd.read_excel(PATH_ENTRADA, sheet_name=HOJA_MEDICIONES)
    df_individuos = pd.read_excel(PATH_ENTRADA, sheet_name=HOJA_INDIVIDUOS)

    print("Columnas de la hoja principal:")
    print(df_principal.columns)
    print("\nColumnas de la hoja 'mediciones':")
    print(df_mediciones.columns)
    print("\nColumnas de la hoja de individuos:")
    print(df_individuos.columns)

    return df_principal, df_mediciones, df_individuos


def _drop_existentes(df, columnas):
    """Elimina solo las columnas que realmente existen en el DataFrame."""
    return df.drop(columns=[c for c in columnas if c in df.columns])


def limpiar_columnas(df_principal, df_mediciones, df_individuos):
    """Elimina los metadatos y columnas auxiliares que no se necesitan."""
    df_principal = _drop_existentes(df_principal, COLS_DROP_PRINCIPAL)
    df_mediciones = _drop_existentes(df_mediciones, COLS_DROP_MEDICIONES)
    df_individuos = _drop_existentes(df_individuos, COLS_DROP_INDIVIDUOS)
    return df_principal, df_mediciones, df_individuos


def procesar_fechas(df_principal, df_mediciones, df_individuos):
    """Normaliza las fechas (descarta la hora) y crea dia, mes y año."""
    df_mediciones['fecha_medicion'] = df_mediciones['fecha_medicion'].str.split('T').str[0]

    df_principal['fecha'] = df_principal['fecha'].astype(str)
    df_principal['fecha'] = df_principal['fecha'].str.split('T').str[0]
    df_individuos['fecha_colecta'] = df_individuos['fecha_colecta'].str.split('T').str[0]

    df_individuos = df_individuos.assign(
        dia=df_individuos['fecha_colecta'].str.split('-').str[2],
        mes=df_individuos['fecha_colecta'].str.split('-').str[1],
        año=df_individuos['fecha_colecta'].str.split('-').str[0],
    )
    df_principal = df_principal.assign(
        dia=df_principal['fecha'].str.split('-').str[2],
        mes=df_principal['fecha'].str.split('-').str[1],
        año=df_principal['fecha'].str.split('-').str[0],
    )
    df_mediciones = df_mediciones.assign(
        dia=df_mediciones['fecha_medicion'].str.split('-').str[2],
        mes=df_mediciones['fecha_medicion'].str.split('-').str[1],
        año=df_mediciones['fecha_medicion'].str.split('-').str[0],
    )
    return df_principal, df_mediciones, df_individuos


# ===========================================================================
# 2. Sistema acuático
# ===========================================================================
def asignar_sistema_acuatico(codigo_sitio):
    """Devuelve el sistema al que pertenece un código de estación."""
    for sistema, codigos in SISTEMAS_ACUATICOS.items():
        if codigo_sitio in codigos:
            return sistema
    return SISTEMA_DESCONOCIDO


def agregar_sistema_acuatico(df_principal, df_mediciones, df_individuos):
    """Crea la columna 'sistema_acuatico' en cada hoja según su código de estación."""
    columnas_id = {
        'principal': ('ID_Sitio_0', df_principal),
        'mediciones': ('ID_Sitio_1', df_mediciones),
        'individuos': ('ID_Sitio_2', df_individuos),
    }
    for nombre, (col_id, df) in columnas_id.items():
        if col_id not in df.columns:
            print(f"La columna '{col_id}' no existe en la hoja de {nombre}")
        else:
            df['sistema_acuatico'] = df[col_id].apply(asignar_sistema_acuatico)
    return df_principal, df_mediciones, df_individuos


# ===========================================================================
# 3. Taxonomía (GBIF) y estado IUCN
# ===========================================================================
def get_taxonomy_gbif(scientific_name):
    """Consulta género, familia y orden de una especie en GBIF."""
    if pd.isna(scientific_name) or str(scientific_name).strip() == '':
        return {'genero': None, 'familia': None, 'orden': None}

    url = "https://api.gbif.org/v1/species/match"
    params = {"name": scientific_name.strip(), "kingdom": "Animalia"}
    try:
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        data = r.json()
        return {
            'genero': data.get('genus'),
            'familia': data.get('family'),
            'orden': data.get('order'),
        }
    except Exception as e:
        print(f"Error consultando '{scientific_name}': {e}")
        return {'genero': None, 'familia': None, 'orden': None}


def agregar_taxonomia(df_individuos):
    """Añade columnas genero, familia y orden usando GBIF (con caché por especie)."""
    especies_unicas = (set(df_individuos['Especie'].dropna().unique())
                       | set(df_individuos['Especie_nueva'].dropna().unique()))
    print(f"Consultando taxonomia para {len(especies_unicas)} especies unicas en GBIF...")

    taxonomia_cache = {}
    for especie in sorted(especies_unicas):
        taxonomia_cache[especie] = get_taxonomy_gbif(especie)
        print(f"  {especie} -> {taxonomia_cache[especie]}")
        time.sleep(0.2)

    def resolver_especie(row):
        nombre = (row['Especie_nueva']
                  if pd.notna(row['Especie_nueva']) and str(row['Especie_nueva']).strip() != ''
                  else row['Especie'])
        return taxonomia_cache.get(nombre, {'genero': None, 'familia': None, 'orden': None})

    df_individuos['genero'] = df_individuos.apply(lambda r: resolver_especie(r)['genero'], axis=1)
    df_individuos['familia'] = df_individuos.apply(lambda r: resolver_especie(r)['familia'], axis=1)
    df_individuos['orden'] = df_individuos.apply(lambda r: resolver_especie(r)['orden'], axis=1)

    # Correcciones manuales de familia por género (ver CONFIGURACIÓN).
    for genero, familia in CORRECCIONES_FAMILIA.items():
        df_individuos.loc[df_individuos['genero'] == genero, 'familia'] = familia

    print("\nColumnas de taxonomia agregadas:")
    print(df_individuos[['Especie', 'Especie_nueva', 'genero', 'familia', 'orden']]
          .drop_duplicates().to_string(index=False))
    return df_individuos


def get_iucn_status(scientific_name):
    """Consulta la categoría IUCN (Red List API v4) por nombre científico.

    Devuelve el código de categoría (LC, NT, VU, EN, CR, ...) de la evaluación
    más reciente, o 'NA' si no hay epíteto válido, token o resultados.
    """
    if pd.isna(scientific_name) or str(scientific_name).strip() == '':
        return 'NA'
    name = str(scientific_name).strip()
    partes = name.split()
    if len(partes) < 2 or partes[1].lower() in _SIN_ESPECIE:
        return 'NA'  # sin epíteto válido -> no se puede consultar
    genero, epiteto = partes[0], partes[1]
    url = f"{IUCN_API_BASE}/taxa/scientific_name"
    headers = {"Authorization": f"Bearer {IUCN_TOKEN}", "accept": "application/json"}
    params = {"genus_name": genero, "species_name": epiteto}
    try:
        r = requests.get(url, params=params, headers=headers, timeout=10)
        r.raise_for_status()
        assessments = r.json().get('assessments', [])
        if not assessments:
            return 'NA'
        # Preferir la evaluación marcada como 'latest'; si no, la primera.
        latest = next((a for a in assessments if a.get('latest')), assessments[0])
        return latest.get('red_list_category_code') or 'NA'
    except Exception as e:
        print(f"Error consultando IUCN '{scientific_name}': {e}")
        return 'NA'


def agregar_iucn(df_individuos):
    """Añade la columna IUCN a la hoja de individuos (con caché por especie)."""
    especies_unicas = (set(df_individuos['Especie'].dropna().unique())
                       | set(df_individuos['Especie_nueva'].dropna().unique()))
    print(f"Consultando estado IUCN para {len(especies_unicas)} especies...")

    iucn_cache = {}
    for especie in sorted(especies_unicas):
        iucn_cache[especie] = get_iucn_status(especie)
        print(f"  {especie} -> {iucn_cache[especie]}")
        time.sleep(0.3)

    def resolver_iucn(row):
        nombre = (row['Especie_nueva']
                  if pd.notna(row['Especie_nueva']) and str(row['Especie_nueva']).strip() != ''
                  else row['Especie'])
        return iucn_cache.get(nombre, 'NA')

    df_individuos['IUCN'] = df_individuos.apply(resolver_iucn, axis=1)
    print("\nEstados IUCN asignados:")
    print(df_individuos[['Especie', 'Especie_nueva', 'IUCN']].drop_duplicates().to_string(index=False))
    return df_individuos


# ===========================================================================
# 4. Dieta y código único
# ===========================================================================
def agregar_dieta(df_individuos):
    """Añade la columna de dieta según el nombre científico (ver DIETA)."""
    df_individuos['dieta'] = df_individuos['Especie'].map(DIETA).fillna('Desconocida')
    return df_individuos


def hacer_codigo_unico(row):
    """Genera un código único por especie, p. ej. 'GENER.ESP'."""
    genero = row['genero']
    especie = row['Especie']

    if genero is None or (isinstance(genero, float) and pd.isna(genero)):
        return ''

    gen_str = str(genero)[:5].upper()
    if pd.notna(especie) and str(especie).strip() != '':
        partes = str(especie).strip().split()
        if len(partes) >= 2 and partes[1].lower() not in _SIN_ESPECIE:
            return gen_str + '.' + partes[1][:3].upper()
    return gen_str


def agregar_codigo_unico(df_individuos):
    """Añade la columna codigo_unico a la hoja de individuos."""
    df_individuos = df_individuos.copy()
    df_individuos['codigo_unico'] = df_individuos.apply(hacer_codigo_unico, axis=1)
    print(df_individuos[['Especie', 'genero', 'codigo_unico']].drop_duplicates().to_string(index=False))
    return df_individuos


# ===========================================================================
# 5. Selección, renombrado, coordenadas y exportación de Individuos
# ===========================================================================
def construir_df_individuos(df_individuos):
    """Selecciona las columnas finales para la base de datos de individuos."""
    columnas = [
        'codigo_individuo', 'fecha_colecta', 'ID_Sitio_2', 'dia', 'año', 'mes',
        'complejo_hídrico_2', 'sistema_acuatico', 'Sitio_colecta', 'aparejo_pez',
        'orden', 'familia', 'genero', 'Especie', 'dieta', 'Talla', 'Peso',
        'Peso_evicerado', 'Peso_tracto', 'Peso_gonada', 'Sexo', 'Estadio',
        'contados', 'codigo_unico', 'IUCN',
    ]
    return df_individuos[[c for c in columnas if c in df_individuos.columns]]


def hacer_asignar_coordenadas_altitud(df_principal):
    """Devuelve una función que mapea un código de estación a (lat, lon, alt)."""
    def asignar_coordenadas_altitud(codigo_sitio):
        fila = df_principal[df_principal['ID_Sitio_0'] == codigo_sitio]
        if not fila.empty:
            return (fila.iloc[0]['_Coordenadas_latitude'],
                    fila.iloc[0]['_Coordenadas_longitude'],
                    fila.iloc[0]['_Coordenadas_altitude'])
        return None, None, None
    return asignar_coordenadas_altitud


def exportar_individuos(df_individuos, df_principal):
    """Filtra por sistema, renombra, agrega coordenadas y exporta cada zona."""
    if not EXPORTES_INDIVIDUOS:
        print("EXPORTES_INDIVIDUOS está vacío: no se exporta ninguna hoja de individuos.")
        return

    asignar = hacer_asignar_coordenadas_altitud(df_principal)
    for sistema, cfg in EXPORTES_INDIVIDUOS.items():
        df_sis = (df_individuos[df_individuos['sistema_acuatico'].isin([sistema])]
                  .rename(columns=RENOMBRE_INDIVIDUOS))

        if 'Codigo (estacion)' in df_sis.columns:
            df_sis['Latitud'], df_sis['Longitud'], df_sis['Altitud'] = zip(
                *df_sis['Codigo (estacion)'].apply(asignar))

        # Leer los encabezados de la plantilla de destino y pegar respetándolos.
        df_destino = pd.read_excel(cfg['plantilla'], sheet_name=HOJA_PLANTILLA_INDIVIDUOS)
        df_destino = df_sis.reindex(columns=df_destino.columns)
        df_destino.to_excel(cfg['salida'], sheet_name=HOJA_PLANTILLA_INDIVIDUOS, index=False)
        print(f"Exportado '{sistema}' -> {cfg['salida']} ({len(df_destino)} filas)")


# ===========================================================================
# 6. Mediciones + Principal -> Hábitat
# ===========================================================================
def calcular_cv_y_merge(df_principal, df_mediciones):
    """Calcula los coeficientes de variación de velocidad/profundidad y los une."""
    df_mediciones = df_mediciones.rename(columns={
        'Sitio_medicion': 'sitio',
        'fecha_medicion': 'fecha',
    })

    cv_velocidad = (df_mediciones[['sitio', 'velocidad']]
                    .groupby('sitio')['velocidad'].agg(['mean', 'std']).reset_index())
    cv_velocidad['cv'] = cv_velocidad['std'] / cv_velocidad['mean']

    cv_profundidad = (df_mediciones[['sitio', 'profundidad']]
                      .groupby('sitio')['profundidad'].agg(['mean', 'std']).reset_index())
    cv_profundidad['cv'] = cv_profundidad['std'] / cv_profundidad['mean']

    cv = pd.merge(cv_velocidad.set_index('sitio')[['cv']],
                  cv_profundidad.set_index('sitio')[['cv']],
                  left_index=True, right_index=True,
                  suffixes=('_velocidad', '_profundidad'))

    df_principal = pd.merge(df_principal, cv, left_on=['sitio'], right_on=['sitio'], how='left')
    return df_principal


def renombrar_principal(df_principal):
    """Renombra las columnas de la hoja Principal para la BD de Hábitat."""
    return df_principal.rename(columns=RENOMBRE_PRINCIPAL)


def exportar_habitat(df_principal_ren):
    """Genera la hoja de Hábitat en formato largo según la plantilla y exporta."""
    # Leer las variables de la plantilla en su orden original.
    try:
        tpl = pd.read_excel(PLANTILLA_HABITAT, sheet_name=HOJA_HABITAT)
        col_var = 'NOMBRE DE VARIABLE' if 'NOMBRE DE VARIABLE' in tpl.columns else 'VARIABLE'
        variables_maestras = tpl[col_var].dropna().astype(str).unique().tolist()
        print(f"Variables leídas de la plantilla ({len(variables_maestras)}):")
        for v in variables_maestras:
            print(f"  '{v}'")
    except Exception as e:
        print(f"No se pudo leer la plantilla: {e}")
        variables_maestras = list(VARIABLES_HABITAT.keys())

    df_base = df_principal_ren.copy()
    df_base['Mes'] = df_base['Mes'].map(
        lambda x: MESES_ES.get(str(x).strip(), x) if pd.notna(x) else pd.NA)
    df_base['Cod2019'] = df_base.get('ID_Sitio', pd.NA)

    id_vars = [c for c in ['Fecha', 'Día', 'Mes', 'Año', 'Estación', 'Cod2019', 'CODIGO(NUEVO)']
               if c in df_base.columns]
    cols_datos = [col for col in VARIABLES_HABITAT.values() if col in df_base.columns]
    cols_display = {col: name for name, col in VARIABLES_HABITAT.items() if col in df_base.columns}

    print("\nColumnas con datos encontradas en la hoja Principal:")
    for name, col in VARIABLES_HABITAT.items():
        estado = "OK " if col in df_base.columns else "NO ENCONTRADA"
        print(f"  [{estado}] '{name}' <- '{col}'")

    # wide -> long
    df_long = df_base.melt(id_vars=id_vars, value_vars=cols_datos,
                           var_name='_col_origen', value_name='VALOR')
    df_long['NOMBRE DE VARIABLE'] = df_long['_col_origen'].map(cols_display)
    df_long = df_long.drop(columns='_col_origen')

    # cross-join: sitios x variables maestras (orden de la plantilla)
    unique_ids = df_long[id_vars].drop_duplicates().reset_index(drop=True)
    vars_df = pd.DataFrame({'NOMBRE DE VARIABLE': variables_maestras})
    unique_ids['__key__'] = 1
    vars_df['__key__'] = 1
    full = unique_ids.merge(vars_df, on='__key__').drop(columns='__key__')

    df_full = full.merge(df_long, on=id_vars + ['NOMBRE DE VARIABLE'], how='left')

    columnas_salida = [
        'Fecha', 'Día', 'Mes', 'Año', 'Letra(campaña)', 'Ambiente', 'Cod2019', 'Estación',
        'Campaña', 'MOMENTO', 'Profundidad_', 'CODIGO(ANTES)', 'CODIGO(NUEVO)',
        'NOMBRE DE VARIABLE', 'VARIABLE', 'VALOR', 'VALOR AJUSTADO', 'OBSERVACIÓN', 'OBSERVACIÓN AJUSTADA',
    ]
    for c in ['Letra(campaña)', 'Ambiente', 'Campaña', 'MOMENTO', 'Profundidad_',
              'CODIGO(ANTES)', 'VALOR AJUSTADO', 'OBSERVACIÓN', 'OBSERVACIÓN AJUSTADA']:
        df_full[c] = pd.NA
    df_full['VARIABLE'] = df_full['NOMBRE DE VARIABLE']

    salida = df_full.reindex(columns=columnas_salida)

    order_map = {v: i for i, v in enumerate(variables_maestras)}
    salida['_var_order'] = salida['NOMBRE DE VARIABLE'].map(order_map)
    salida = (salida.sort_values(['Estación', 'Fecha', '_var_order'])
              .drop(columns='_var_order').reset_index(drop=True))

    if os.path.exists(SALIDA_HABITAT):
        os.remove(SALIDA_HABITAT)
    with pd.ExcelWriter(SALIDA_HABITAT, engine='openpyxl', mode='w') as writer:
        salida.to_excel(writer, sheet_name=HOJA_HABITAT, index=False)

    print(f'\nExportación completada -> {SALIDA_HABITAT}')
    print(f'Filas exportadas:    {len(salida)}')
    print(f'Sitios (Estación):   {salida["Estación"].nunique()}')
    print(f'Variables maestras:  {len(variables_maestras)}')
    print(f'Variables con datos: {len(cols_datos)} de {len(VARIABLES_HABITAT)}')


# ===========================================================================
# Flujo principal
# ===========================================================================
def main():
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)

    # 1. Carga y limpieza
    df_principal, df_mediciones, df_individuos = cargar_datos()
    df_principal, df_mediciones, df_individuos = limpiar_columnas(
        df_principal, df_mediciones, df_individuos)
    df_principal, df_mediciones, df_individuos = procesar_fechas(
        df_principal, df_mediciones, df_individuos)
    df_principal, df_mediciones, df_individuos = agregar_sistema_acuatico(
        df_principal, df_mediciones, df_individuos)

    # 2. Taxonomía, IUCN, dieta y código único (hoja de individuos)
    df_individuos = agregar_taxonomia(df_individuos)
    df_individuos = agregar_iucn(df_individuos)
    df_individuos = agregar_dieta(df_individuos)
    df_individuos = agregar_codigo_unico(df_individuos)

    # 3. Construir y exportar las hojas de individuos por sistema
    df_final = construir_df_individuos(df_individuos)
    exportar_individuos(df_final, df_principal)

    # 4. Hábitat (Mediciones + Principal)
    df_principal = calcular_cv_y_merge(df_principal, df_mediciones)
    df_principal_ren = renombrar_principal(df_principal)
    exportar_habitat(df_principal_ren)


if __name__ == "__main__":
    main()
