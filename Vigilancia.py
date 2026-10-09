import streamlit as st
import pandas as pd
from datetime import datetime
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from sqlalchemy import create_engine, text

# 1. Configuración de la página web
st.set_page_config(
    page_title="Control de Portería & Recepción de Telas",
    layout="wide",
    page_icon="🚛"
)

# 2. Control de Acceso / Login
PASSWORD_CORRECTA = "garita123"  # Puedes cambiar la contraseña aquí

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

def check_password():
    st.title("🔐 Acceso al Sistema de Vigilancia")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd_input = st.text_input("Ingrese la contraseña de acceso:", type="password")
        if st.button("Iniciar Sesión", use_container_width=True):
            if pwd_input == PASSWORD_CORRECTA:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta. Intente nuevamente.")

# Bloquear la ejecución si el usuario no ha iniciado sesión
if not st.session_state["autenticado"]:
    check_password()
    st.stop()

# ----------------------------------------------------
# 3. Conexión y Gestión de Base de Datos (Supabase)
# ----------------------------------------------------
def get_engine():
    db_url = st.secrets["postgres"]["url"]
    engine = create_engine(db_url)
    return engine

def init_db():
    engine = get_engine()
    with engine.begin() as conn:
        # Tabla de Movimientos / Despachos
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
        
        # Tabla de Stock Consolidado
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
    st.error(f"⚠️ Error de conexión a la base de datos: {e}")

# ----------------------------------------------------
# 4. Encabezado principal y Menú Lateral
# ----------------------------------------------------
col_header, col_logout = st.columns([8, 2])
with col_header:
    st.title("🚛 Control de Entrada y Salida - Vigilancia")
with col_logout:
    st.write("")
    if st.button("🚪 Cerrar Sesión"):
        st.session_state["autenticado"] = False
        st.rerun()

opcion = st.sidebar.radio("Navegación", ["Registrar Movimiento", "Inventario y Stock", "Historial e Informes"])

# ----------------------------------------------------
# MÓDULO 1: REGISTRAR MOVIMIENTO (ENTRADA / SALIDA)
# ----------------------------------------------------
if opcion == "Registrar Movimiento":
    st.header("📝 Registrar Nuevo Movimiento de Telas")
    
    col1, col2 = st.columns(2)
    
    with col1:
        tipo_movimiento = st.selectbox("Tipo de Movimiento", ["Entrada (Ingreso)", "Salida (Despacho)"])
        conductor = st.text_input("Nombre del Conductor *")
        placa = st.text_input("Placa del Vehículo *")
        empresa = st.text_input("Empresa / Proveedor / Cliente")
        
    with col2:
        guia = st.text_input("N° de Guía de Remisión / Documento")
        tipo_rollo = st.text_input("Especificación / Código de Rollo *")
        cantidad_rollos = st.number_input("Cantidad de Rollos", min_value=1, step=1, value=1)
        peso_kg = st.number_input("Peso Total (kg)", min_value=0.0, step=0.1, value=0.0)
        
    observaciones = st.text_area("Observaciones o Incidencias")
    
    if st.button("💾 Guardar Registro", use_container_width=True):
        if conductor.strip() and placa.strip() and tipo_rollo.strip():
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            engine = get_engine()
            
            with engine.begin() as conn:
                # 1. Insertar el movimiento en la bitácora
                conn.execute(text("""
                    INSERT INTO despachos (fecha_hora, tipo_movimiento, conductor, placa, empresa, guia, tipo_rollo, cantidad_rollos, peso_kg, observaciones)
                    VALUES (:fecha, :tipo_mov, :cond, :placa, :emp, :guia, :tipo_r, :cant, :peso, :obs)
                """), {
                    "fecha": fecha_actual, "tipo_mov": tipo_movimiento, "cond": conductor,
                    "placa": placa, "emp": empresa, "guia": guia, "tipo_r": tipo_rollo,
                    "cant": cantidad_rollos, "peso": peso_kg, "obs": observaciones
                })
                
                # 2. Actualizar el inventario en tiempo real
                res = conn.execute(
                    text("SELECT cantidad, peso_total_kg FROM stock WHERE tipo_rollo = :tipo_r"), 
                    {"tipo_r": tipo_rollo}
                ).fetchone()
                
                if tipo_movimiento == "Entrada (Ingreso)":
                    if res:
                        nueva_cant = res[0] + cantidad_rollos
                        nuevo_peso = res[1] + peso_kg
                        conn.execute(text("""
                            UPDATE stock SET cantidad = :cant, peso_total_kg = :peso WHERE tipo_rollo = :tipo_r
                        """), {"cant": nueva_cant, "peso": nuevo_peso, "tipo_r": tipo_rollo})
                    else:
                        conn.execute(text("""
                            INSERT INTO stock (tipo_rollo, cantidad, peso_total_kg) VALUES (:tipo_r, :cant, :peso)
                        """), {"tipo_r": tipo_rollo, "cant": cantidad_rollos, "peso": peso_kg})
                else:  # Salida
                    if res:
                        nueva_cant = max(0, res[0] - cantidad_rollos)
                        nuevo_peso = max(0.0, res[1] - peso_kg)
                        conn.execute(text("""
                            UPDATE stock SET cantidad = :cant, peso_total_kg = :peso WHERE tipo_rollo = :tipo_r
                        """), {"cant": nueva_cant, "peso": nuevo_peso, "tipo_r": tipo_rollo})
                    else:
                        st.warning("⚠️ Se registró la salida de un tipo de rollo que no registraba stock previo.")
            
            st.success("✅ ¡Movimiento registrado correctamente y respaldado en la nube!")
        else:
            st.error("❌ Por favor completa los campos obligatorios: Conductor, Placa y Tipo de Rollo.")

# ----------------------------------------------------
# MÓDULO 2: INVENTARIO Y STOCK EN PLANTA
# ----------------------------------------------------
elif opcion == "Inventario y Stock":
    st.header("📦 Stock Actual Disponible en Planta")
    
    engine = get_engine()
    df_stock = pd.read_sql_query("SELECT * FROM stock ORDER BY tipo_rollo ASC", engine)
    
    if not df_stock.empty:
        df_stock.columns = ["Especificación del Rollo", "Cantidad Rollos", "Peso Total (kg)"]
        
        # Tarjetas de resumen métrico
        col_m1, col_m2 = st.columns(2)
        col_m1.metric("Total Rollos en Stock", int(df_stock["Cantidad Rollos"].sum()))
        col_m2.metric("Peso Total Acumulado (kg)", f"{df_stock['Peso Total (kg)'].sum():,.2f}")
        
        st.dataframe(df_stock, use_container_width=True)
    else:
        st.info("No hay stock registrado en el sistema.")

# ----------------------------------------------------
# MÓDULO 3: HISTORIAL COMPLETO, REPORTE Y EXPORTACIÓN
# ----------------------------------------------------
elif opcion == "Historial e Informes":
    st.header("📋 Bitácora de Movimientos y Generación de Informes")
    
    engine = get_engine()
    df_historial = pd.read_sql_query("SELECT * FROM despachos ORDER BY id DESC", engine)
    
    if not df_historial.empty:
        df_historial.columns = [
            "ID", "Fecha / Hora", "Tipo Movimiento", "Conductor", "Placa", 
            "Empresa", "N° Guía", "Tipo de Rollo", "Cantidad", "Peso (kg)", "Observaciones"
        ]
        
        st.dataframe(df_historial, use_container_width=True)
        
        # Generación de archivo Excel estructurado para descarga
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Historial Despachos"
        
        for r in dataframe_to_rows(df_historial, index=False, header=True):
            ws.append(r)
            
        wb.save(output)
        excel_data = output.getvalue()
        
        st.download_button(
            label="📊 Descargar Informe Completo en Excel (.xlsx)",
            data=excel_data,
            file_name=f"Reporte_Porteria_Vigilancia_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("No existen registros grabados en la bitácora.")
