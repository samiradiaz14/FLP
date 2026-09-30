import streamlit as st
import pandas as pd
import sqlite3
import re
import datetime

# --- CONFIGURACIÓN DE PÁGINA (Tarea 5) ---
st.set_page_config(page_title="Sistema Fundación", layout="wide")
st.title("📊 Gestor de Consultas - Fundación")

DB_NAME = "fundacion.db"

# --- TAREA 1: REGLAS Y DICCIONARIO DE EQUIVALENCIAS ---
DICCIONARIO_LOCACIONES = {
    "prof park": "Parque",
    "park": "Parque",
    "parque": "Parque",
    "general": "Rancho",
    "lunch": "Rancho",
    "tier 3": "Rancho",
    "special events": "Rancho"
}

def categorizar_locacion(nombre_pestana):
    nombre_clean = str(nombre_pestana).lower().strip()
    for clave, locacion in DICCIONARIO_LOCACIONES.items():
        if clave in nombre_clean:
            return locacion
    return "Rancho"

def limpiar_monto(val):
    if pd.isna(val):
        return 0.0
    raw_monto = str(val).replace("$", "").replace(",", "").strip()
    parts = raw_monto.split(".")
    if len(parts) == 2 and len(parts[1]) == 3:
        raw_monto = raw_monto.replace(".", "")
    try:
        return float(raw_monto)
    except ValueError:
        return 0.0

def limpiar_texto(texto):
    if pd.isna(texto) or str(texto).lower().strip() in ["nan", "none", "null", ""]:
        return ""
    return str(texto).strip()

# --- TAREA 3: MODELO RELACIONAL (SQLite Normalizado) ---
def inicializar_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    
    # Tabla Personas
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS personas (
            id_persona INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE,
            nombre TEXT,
            apellido TEXT,
            direccion TEXT
        )
    ''')
    
    # Tabla Visitas
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS visitas (
            id_visita INTEGER PRIMARY KEY AUTOINCREMENT,
            id_persona INTEGER,
            fecha_visita DATE,
            cuarto TEXT,
            locacion TEXT,
            origen TEXT,
            FOREIGN KEY (id_persona) REFERENCES personas(id_persona) ON DELETE CASCADE
        )
    ''')
    
    # Tabla Donaciones
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS donaciones (
            id_donacion INTEGER PRIMARY KEY AUTOINCREMENT,
            id_persona INTEGER,
            monto REAL,
            fecha_donacion DATE,
            origen TEXT,
            FOREIGN KEY (id_persona) REFERENCES personas(id_persona) ON DELETE CASCADE
        )
    ''')
    conn.commit()
    conn.close()

def obtener_o_crear_persona(cursor, email, nombre, apellido, direccion=""):
    email_clean = email.lower().strip() if email else ""
    
    if email_clean:
        cursor.execute("SELECT id_persona FROM personas WHERE email = ?", (email_clean,))
        row = cursor.fetchone()
        if row:
            return row[0]
            
    cursor.execute(
        "INSERT INTO personas (email, nombre, apellido, direccion) VALUES (?, ?, ?, ?)",
        (email_clean if email_clean else None, nombre, apellido, direccion)
    )
    return cursor.lastrowid

# --- TAREA 2: PIPELINE ETL DINÁMICO ---
def procesar_excel(file):
    xls = pd.ExcelFile(file)
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Reiniciar base de datos para nueva carga
    cursor.execute("DELETE FROM visitas")
    cursor.execute("DELETE FROM donaciones")
    cursor.execute("DELETE FROM personas")
    
    POSIBLES_COLS_VISITA = ["date of visit", "visit date", "fecha de visita", "fecha visita", "date", "fecha"]
    POSIBLES_COLS_DONACION = ["donation date", "fecha donacion", "fecha de donacion", "date of donation"]
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        # Definición explícita de las columnas detectadas por pestaña
        col_fecha_visita = next((col for col in POSIBLES_COLS_VISITA if col in df.columns), None)
        col_fecha_donacion = next((col for col in POSIBLES_COLS_DONACION if col in df.columns), None)
        locacion_cat = categorizar_locacion(sheet_name)
        
        # 1. Procesar Visitas
        if col_fecha_visita:
            for _, row in df.iterrows():
                email = limpiar_texto(row.get("email"))
                nombre = limpiar_texto(row.get("name") or row.get("nombre"))
                apellido = limpiar_texto(row.get("last name") or row.get("apellido"))
                direccion = limpiar_texto(row.get("home address") or row.get("address"))
                fecha = pd.to_datetime(row.get(col_fecha_visita), errors='coerce')
                cuarto = limpiar_texto(row.get("room number") or row.get("cuarto") or row.get("room"))
                
                if pd.notnull(fecha) and (email or nombre):
                    id_p = obtener_o_crear_persona(cursor, email, nombre, apellido, direccion)
                    cursor.execute(
                        "INSERT INTO visitas (id_persona, fecha_visita, cuarto, locacion, origen) VALUES (?, ?, ?, ?, ?)",
                        (id_p, fecha.strftime('%Y-%m-%d'), cuarto, locacion_cat, sheet_name)
                    )
                    
        # 2. Procesar Donaciones
        if col_fecha_donacion:
            col_monto = next((col for col in ["monto", "amount", "donation", "donation amount", "donacion"] if col in df.columns), None)

            for _, row in df.iterrows():
                email = limpiar_texto(row.get("email"))
                nombre = limpiar_texto(row.get("name") or row.get("nombre"))
                apellido = limpiar_texto(row.get("last name") or row.get("apellido"))
                direccion = limpiar_texto(row.get("home address") or row.get("address"))
                fecha = pd.to_datetime(row.get(col_fecha_donacion), errors='coerce')
                monto_val = limpiar_monto(row.get(col_monto)) if col_monto else 0.0
                
                if pd.notnull(fecha) and (email or nombre):
                    id_p = obtener_o_crear_persona(cursor, email, nombre, apellido, direccion)
                    cursor.execute(
                        "INSERT INTO donaciones (id_persona, monto, fecha_donacion, origen) VALUES (?, ?, ?, ?)",
                        (id_p, monto_val, fecha.strftime('%Y-%m-%d'), sheet_name)
                    )
                    
    conn.commit()
    conn.close()

# --- TAREA 4: MOTOR DE CONSULTAS Y ANALÍTICA ---
def obtener_estacion(mes):
    if mes in [12, 1, 2]: return "Invierno"
    elif mes in [3, 4, 5]: return "Primavera"
    elif mes in [6, 7, 8]: return "Verano"
    else: return "Otoño"

def calcular_analitica(donaciones_df, visitas_df):
    if not donaciones_df.empty:
        fechas = pd.to_datetime(donaciones_df["fecha_donacion"])
        estaciones = fechas.dt.month.apply(obtener_estacion)
        epoca = estaciones.mode()[0] if not estaciones.empty else "Sin registros"
    else:
        epoca = "Sin registros"
        
    if visitas_df.empty or donaciones_df.empty:
        tiempo_prom = "N/A"
    else:
        v_df = visitas_df.copy()
        d_df = donaciones_df.copy()
        v_df["fecha_visita"] = pd.to_datetime(v_df["fecha_visita"])
        d_df["fecha_donacion"] = pd.to_datetime(d_df["fecha_donacion"])
        
        dias_list = []
        for _, don in d_df.iterrows():
            f_don = don["fecha_donacion"]
            prev = v_df[v_df["fecha_visita"] < f_don]
            if not prev.empty:
                dias_list.append((f_don - prev["fecha_visita"].max()).days)
                
        tiempo_prom = f"{int(sum(dias_list)/len(dias_list))} días" if dias_list else "Sin visitas previas"
        
    return epoca, tiempo_prom

# --- TAREA 5: INTERFAZ STREAMLIT ---
inicializar_db()

st.sidebar.header("📁 Carga de Datos")
archivo = st.sidebar.file_uploader("Sube el Excel semanal", type=["xlsx"])
if archivo:
    with st.spinner("Ejecutando ETL y guardando en SQLite..."):
        procesar_excel(archivo)
    st.sidebar.success("¡Base de Datos Actualizada!")

tab1, tab2 = st.tabs(["🔎 Consultas por Persona", "🏆 Rankings y Reportes"])

conn = sqlite3.connect(DB_NAME)

with tab1:
    personas_df = pd.read_sql_query("SELECT id_persona, email, nombre, apellido FROM personas", conn)
    
    if not personas_df.empty:
        personas_df["display"] = personas_df["email"].fillna("Sin Email") + " - " + personas_df["nombre"].fillna("") + " " + personas_df["apellido"].fillna("")
        dict_personas = dict(zip(personas_df["display"], personas_df["id_persona"]))
        opciones_busqueda = [""] + list(dict_personas.keys())
    else:
        dict_personas = {}
        opciones_busqueda = [""]

    opcion = st.selectbox("Buscar por Email o Nombre:", opciones_busqueda)
    
    if opcion:
        selected_id = dict_personas[opcion]
        visitas_df = pd.read_sql_query(f"SELECT * FROM visitas WHERE id_persona = {selected_id}", conn)
        donaciones_df = pd.read_sql_query(f"SELECT * FROM donaciones WHERE id_persona = {selected_id}", conn)
        
        if not visitas_df.empty:
            visitas_parque = len(visitas_df[visitas_df["locacion"] == "Parque"])
            visitas_rancho = len(visitas_df[visitas_df["locacion"] == "Rancho"])
        else:
            visitas_parque = 0
            visitas_rancho = 0

        total_visitas = len(visitas_df)
        total_donado = donaciones_df["monto"].sum() if not donaciones_df.empty else 0.0
        prom_donado = donaciones_df["monto"].mean() if not donaciones_df.empty and total_donado > 0 else 0.0
        
        anios = sorted(pd.to_datetime(donaciones_df["fecha_donacion"]).dt.year.unique().tolist()) if not donaciones_df.empty else []
        epoca, tiempo_prom = calcular_analitica(donaciones_df, visitas_df)
        
        st.subheader("📊 Tarjetas de Métricas KPI")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Visitas Totales", total_visitas)
        c2.metric("Visitas Parque / Rancho", f"{visitas_parque} / {visitas_rancho}")
        c3.metric("Total Donado", f"${total_donado:,.2f}")
        c4.metric("Promedio Donado", f"${prom_donado:,.2f}")
        
        c5, c6, c7 = st.columns(3)
        c5.metric("Años Activos", ", ".join(map(str, anios)) if anios else "Ninguno")
        c6.metric("Época de Mayor Frecuencia", epoca)
        c7.metric("Tiempo Prom. Visita ➔ Donación", tiempo_prom)
        
        st.markdown("---")
        st.subheader("📋 Historiales")
        col_v, col_d = st.columns(2)
        with col_v:
            st.write("**Visitas:**")
            if not visitas_df.empty:
                st.dataframe(visitas_df[["fecha_visita", "locacion", "origen", "cuarto"]], use_container_width=True)
            else:
                st.info("Sin registros de visitas.")
        with col_d:
            st.write("**Donaciones:**")
            if not donaciones_df.empty:
                st.dataframe(donaciones_df[["fecha_donacion", "monto", "origen"]], use_container_width=True)
            else:
                st.info("Sin registros de donaciones.")

with tab2:
    st.header("🏆 Rankings de Mayor a Menor")
    f_rank = st.radio("Filtrar Ranking:", ["Mayor Donación Total", "Más Visitas al Parque", "Más Visitas al Rancho"], horizontal=True)
    
    st.markdown("---")
    if f_rank == "Mayor Donación Total":
        query = '''
            SELECT p.email AS 'Correo', p.nombre || ' ' || p.apellido AS 'Nombre', 
                   COUNT(d.id_donacion) AS 'Total Aportes', SUM(d.monto) AS 'Total Donado ($)' 
            FROM donaciones d
            JOIN personas p ON d.id_persona = p.id_persona
            GROUP BY p.id_persona
            ORDER BY SUM(d.monto) DESC
        '''
        df_rank = pd.read_sql_query(query, conn)
        if not df_rank.empty:
            st.dataframe(df_rank.style.format({"Total Donado ($)": "${:,.2f}"}), use_container_width=True)
        else:
            st.info("Sin datos registrados.")

    elif f_rank == "Más Visitas al Parque":
        query = '''
            SELECT p.email AS 'Correo', p.nombre || ' ' || p.apellido AS 'Nombre', 
                   COUNT(v.id_visita) AS 'Visitas Parque' 
            FROM visitas v
            JOIN personas p ON v.id_persona = p.id_persona
            WHERE v.locacion = 'Parque'
            GROUP BY p.id_persona
            ORDER BY COUNT(v.id_visita) DESC
        '''
        df_rank = pd.read_sql_query(query, conn)
        if not df_rank.empty:
            st.dataframe(df_rank, use_container_width=True)
        else:
            st.info("Sin datos registrados.")

    else:
        query = '''
            SELECT p.email AS 'Correo', p.nombre || ' ' || p.apellido AS 'Nombre', 
                   COUNT(v.id_visita) AS 'Visitas Rancho' 
            FROM visitas v
            JOIN personas p ON v.id_persona = p.id_persona
            WHERE v.locacion = 'Rancho'
            GROUP BY p.id_persona
            ORDER BY COUNT(v.id_visita) DESC
        '''
        df_rank = pd.read_sql_query(query, conn)
        if not df_rank.empty:
            st.dataframe(df_rank, use_container_width=True)
        else:
            st.info("Sin datos registrados.")

conn.close()