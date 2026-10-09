import streamlit as st
import pandas as pd
from datetime import datetime
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from sqlalchemy import create_engine, text

# Configuración de la página web
st.set_page_config(
    page_title="Control de Portería & Recepción de Telas",
    layout="wide",
    page_icon="🚛"
)

# Conexión a la base de datos PostgreSQL de Supabase vía SQLAlchemy
def get_engine():
    db_url = st.secrets["postgres"]["url"]
    engine = create_engine(db_url)
    return engine

# Inicializar tablas en Supabase
def init_db():
    engine = get_engine()
    with engine.begin() as conn:
        # Tabla de Movimientos / Bitácora de Despachos
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS despachos (
                id SERIAL PRIMARY KEY,
                fecha_hora VARCHAR(50),
                tipo_movimiento VARCHAR(50),
                conductor VARCHAR(100),
                placa VARCHAR(20),
                empresa VARCHAR(100),
                guia VARCHAR(50),
                tipo_rollo VARCHAR(100),
                cantidad_rollos INT,
                peso_kg DOUBLE PRECISION,
                observaciones TEXT
            )
        """))
        
        # Tabla de Stock General
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS stock (
                tipo_rollo VARCHAR(100) PRIMARY KEY,
                cantidad INT,
                peso_total_kg DOUBLE PRECISION
            )
        """))

try:
    init_db()
except Exception as e:
    st.error(f"Error al conectar con la base de datos de Supabase: {e}")

st.title("🚛 Control de Entrada y Salida - Vigilancia")

# Menú lateral
opcion = st.sidebar.radio("Navegación", ["Registrar Movimiento", "Inventario y Stock", "Historial e Informes"])

# 1. REGISTRAR MOVIMIENTO
if opcion == "Registrar Movimiento":
    st.header("📝 Registrar Nuevo Movimiento (Entrada / Salida)")
    
    col1, col2 = st.columns(2)
    
    with col1:
        tipo_movimiento = st.selectbox("Tipo de Movimiento", ["Entrada (Ingreso)", "Salida (Despacho)"])
        conductor = st.text_input("Nombre del Conductor")
        placa = st.text_input("Placa del Vehículo")
        empresa = st.text_input("Empresa / Proveedor / Cliente")
        
    with col2:
        guia = st.text_input("N° de Guía / Documento")
        tipo_rollo = st.text_input("Tipo / Especificación del Rollo")
        cantidad_rollos = st.number_input("Cantidad de Rollos", min_value=1, step=1, value=1)
        peso_kg = st.number_input("Peso Total (kg)", min_value=0.0, step=0.1, value=0.0)
        
    observaciones = st.text_area("Observaciones Adicionales")
    
    if st.button("💾 Guardar Registro", use_container_width=True):
        if conductor and placa and tipo_rollo:
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            engine = get_engine()
            
            with engine.begin() as conn:
                # 1. Registrar el despacho
                conn.execute(text("""
                    INSERT INTO despachos (fecha_hora, tipo_movimiento, conductor, placa, empresa, guia, tipo_rollo, cantidad_rollos, peso_kg, observaciones)
                    VALUES (:fecha, :tipo_mov, :cond, :placa, :emp, :guia, :tipo_r, :cant, :peso, :obs)
                """), {
                    "fecha": fecha_actual, "tipo_mov": tipo_movimiento, "cond": conductor,
                    "placa": placa, "emp": empresa, "guia": guia, "tipo_r": tipo_rollo,
                    "cant": cantidad_rollos, "peso": peso_kg, "obs": observaciones
                })
                
                # 2. Consultar stock existente
                res = conn.execute(text("SELECT cantidad, peso_total_kg FROM stock WHERE tipo_rollo = :tipo_r"), {"tipo_r": tipo_rollo}).fetchone()
                
                if tipo_movimiento == "Entrada (Ingreso)":
                    if res:
                        nueva_cant = res[0] + cantidad_rollos
                        nuevo_peso = res[1] + peso_kg
                        conn.execute(text("UPDATE stock SET cantidad = :cant, peso_total_kg = :peso WHERE tipo_rollo = :tipo_r"), {
                            "cant": nueva_cant, "peso": nuevo_peso, "tipo_r": tipo_rollo
                        })
                    else:
                        conn.execute(text("INSERT INTO stock VALUES (:tipo_r, :cant, :peso)"), {
                            "tipo_r": tipo_rollo, "cant": cantidad_rollos, "peso": peso_kg
                        })
                else: # Salida
                    if res:
                        nueva_cant = max(0, res[0] - cantidad_rollos)
                        nuevo_peso = max(0.0, res[1] - peso_kg)
                        conn.execute(text("UPDATE stock SET cantidad = :cant, peso_total_kg = :peso WHERE tipo_rollo = :tipo_r"), {
                            "cant": nueva_cant, "peso": nuevo_peso, "tipo_r": tipo_rollo
                        })
                    else:
                        st.warning("⚠️ Se registró la salida de un material que no figuraba en stock registrado.")
            
            st.success("✅ ¡Registro guardado exitosamente en la nube!")
        else:
            st.error("❌ Por favor completa los campos obligatorios (Conductor, Placa y Tipo de Rollo).")

# 2. INVENTARIO Y STOCK
elif opcion == "Inventario y Stock":
    st.header("📦 Stock Actual en Planta")
    
    engine = get_engine()
    df_stock = pd.read_sql_query("SELECT * FROM stock", engine)
    
    if not df_stock.empty:
        df_stock.columns = ["Tipo de Rollo", "Cantidad Disponible", "Peso Total (kg)"]
        st.dataframe(df_stock, use_container_width=True)
    else:
        st.info("No hay registros de stock actualmente en la nube.")

# 3. HISTORIAL E INFORMES
elif opcion == "Historial e Informes":
    st.header("📋 Historial de Registros y Exportación")
    
    engine = get_engine()
    df_historial = pd.read_sql_query("SELECT * FROM despachos ORDER BY id DESC", engine)
    
    if not df_historial.empty:
        st.dataframe(df_historial, use_container_width=True)
        
        # Generar Excel con openpyxl
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reporte Vigilancia"
        
        for r in dataframe_to_rows(df_historial, index=False, header=True):
            ws.append(r)
            
        wb.save(output)
        excel_data = output.getvalue()
        
        st.download_button(
            label="📊 Descargar Reporte en Excel",
            data=excel_data,
            file_name=f"Reporte_Vigilancia_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("No hay registros en el historial.")
