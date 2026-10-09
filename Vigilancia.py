import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import io
import os
import glob
from sqlalchemy import create_engine, text

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="Control de Carga de Rollos",
    page_icon="🚛",
    layout="wide"
)

# Usuarios locales del sistema
USUARIOS_REGISTRADOS = {
    "operador1": {"password": "eiick", "nombre": "Erick Jimenez"},
    "operador2": {"password": "protex.vigilancia", "nombre": "Chipana"},
    "admin": {"password": "Josepkiri9651535447", "nombre": "admin"}
}

RUTA_RED_SYSTEM = r"\\192.168.1.4\system"

# 2. CONEXIÓN A LA BASE DE DATOS EN LA NUBE (SUPABASE / POSTGRESQL)
def get_engine():
    db_url = st.secrets["postgres"]["url"]
    engine = create_engine(db_url)
    return engine

def init_db():
    engine = get_engine()
    with engine.begin() as conn:
        # Tabla de Historial / Bitácora de Despachos
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS historial_despachos (
                id SERIAL PRIMARY KEY,
                fecha_hora VARCHAR(50),
                operario VARCHAR(100),
                partida_cargada VARCHAR(50),
                codigo_rollo VARCHAR(100),
                cliente VARCHAR(100),
                peso_kg DOUBLE PRECISION,
                estado VARCHAR(50),
                observacion TEXT
            )
        """))

try:
    init_db()
except Exception as e:
    st.error(f"⚠️ Error al conectar con Supabase: {e}")

def cargar_historial():
    engine = get_engine()
    try:
        df = pd.read_sql_query("SELECT * FROM historial_despachos ORDER BY id DESC", engine)
        if not df.empty:
            df.columns = ["ID", "Fecha_Hora", "Operario", "Partida_Cargada", "Codigo_Rollo", "Cliente", "Peso_KG", "Estado", "Observacion"]
            df["Peso_KG"] = df["Peso_KG"].astype(str)
            return df
        else:
            return pd.DataFrame(columns=["ID", "Fecha_Hora", "Operario", "Partida_Cargada", "Codigo_Rollo", "Cliente", "Peso_KG", "Estado", "Observacion"])
    except Exception:
        return pd.DataFrame(columns=["ID", "Fecha_Hora", "Operario", "Partida_Cargada", "Codigo_Rollo", "Cliente", "Peso_KG", "Estado", "Observacion"])

def registrar_despacho(operario, partida_sel, codigo, cliente, peso, estado, obs):
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    engine = get_engine()
    with engine.begin() as conn:
        conn.execute(text("""
            INSERT INTO historial_despachos (fecha_hora, operario, partida_cargada, codigo_rollo, cliente, peso_kg, estado, observacion)
            VALUES (:fecha, :operario, :partida, :codigo, :cliente, :peso, :estado, :obs)
        """), {
            "fecha": fecha_actual, "operario": operario, "partida": partida_sel,
            "codigo": codigo, "cliente": cliente, "peso": float(peso),
            "estado": estado, "obs": obs
        })

# 3. FUNCIONES DE AUDIO
def reproducir_sonido_alarma():
    st.components.v1.html("""
        <script>
            var ctx = new (window.AudioContext || window.webkitAudioContext)();
            function alertTone() {
                var osc = ctx.createOscillator();
                osc.type = 'sawtooth';
                osc.frequency.setValueAtTime(850, ctx.currentTime);
                osc.connect(ctx.destination);
                osc.start();
                osc.stop(ctx.currentTime + 0.8);
            }
            alertTone();
        </script>
    """, height=0)

def reproducir_sonido_ok():
    st.components.v1.html("""
        <script>
            var ctx = new (window.AudioContext || window.webkitAudioContext)();
            function okTone() {
                var osc = ctx.createOscillator();
                osc.type = 'sine';
                osc.frequency.setValueAtTime(523.25, ctx.currentTime);
                osc.connect(ctx.destination);
                osc.start();
                osc.stop(ctx.currentTime + 0.2);
            }
            okTone();
        </script>
    """, height=0)

def obtener_ultimo_reporte_xls():
    if os.path.exists(RUTA_RED_SYSTEM):
        patron = os.path.join(RUTA_RED_SYSTEM, "*.xls")
        archivos = glob.glob(patron)
        if archivos:
            return max(archivos, key=os.path.getmtime)
    return None

# 4. ESTADOS DE SESIÓN
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""
    st.session_state.nombre_operario = ""

if "df_rollos_master" not in st.session_state:
    st.session_state.df_rollos_master = None

if "partida_activa" not in st.session_state:
    st.session_state.partida_activa = None

# 5. PANTALLA DE LOGIN
if not st.session_state.autenticado:
    st.title("🔑 Control de Acceso al Sistema de Despacho")
    with st.form("form_login"):
        user_input = st.text_input("Usuario:").strip()
        pass_input = st.text_input("Contraseña:", type="password").strip()
        if st.form_submit_button("Ingresar", use_container_width=True):
            if user_input in USUARIOS_REGISTRADOS and USUARIOS_REGISTRADOS[user_input]["password"] == pass_input:
                st.session_state.autenticado = True
                st.session_state.usuario_actual = user_input
                st.session_state.nombre_operario = USUARIOS_REGISTRADOS[user_input]["nombre"]
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")

# 6. PANTALLA PRINCIPAL DE OPERACIONES
else:
    col_head1, col_head2 = st.columns([4, 1])
    col_head1.title("🚛 Control de Despacho de Rollos")
    col_head1.caption(f"Operario en sesión: **{st.session_state.nombre_operario}**")
    
    if col_head2.button("🚪 Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.df_rollos_master = None
        st.session_state.partida_activa = None
        st.rerun()

    st.markdown("---")

    # MÓDULO BARRA LATERAL (CARGA DE FUENTE DE DATOS Y NAVEGACIÓN)
    with st.sidebar:
        st.header("📂 Fuente de Datos de Planta")
        
        if st.button("🔄 Leer Último Reporte del Servidor", use_container_width=True):
            ultimo_archivo = obtener_ultimo_reporte_xls()
            if ultimo_archivo:
                try:
                    df_temp = pd.read_excel(ultimo_archivo, dtype=str)
                    st.session_state.df_rollos_master = df_temp
                    st.success(f"✅ Cargado: `{os.path.basename(ultimo_archivo)}`")
                except Exception as e:
                    st.error(f"Error al leer archivo de red: {e}")
            else:
                st.warning("⚠️ No se encontró la ruta de red o no hay archivos .xls.")

        st.markdown("---")
        archivo_subido = st.file_uploader("O subir archivo manualmente:", type=["xlsx", "xls", "csv"])
        if archivo_subido is not None:
            try:
                if archivo_subido.name.endswith('.csv'):
                    df_temp = pd.read_csv(archivo_subido, dtype=str)
                else:
                    df_temp = pd.read_excel(archivo_subido, dtype=str)
                
                st.session_state.df_rollos_master = df_temp
                st.success("✅ ¡Base cargada manualmente!")
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")

    # MODO DEMOSTRACIÓN O LECTURA AUTOMÁTICA
    if st.session_state.df_rollos_master is None:
        archivo_auto = obtener_ultimo_reporte_xls()
        if archivo_auto:
            try:
                st.session_state.df_rollos_master = pd.read_excel(archivo_auto, dtype=str)
                st.info(f"🌐 Conectado automáticamente al archivo del servidor: `{os.path.basename(archivo_auto)}`")
            except Exception:
                pass

    if st.session_state.df_rollos_master is None:
        st.info("💡 Modo Demostración activo. Carga un archivo en la barra lateral.")
        datos_demo = {
            "Codigo_Barras": ["A60933539", "A60933540", "A60933541", "A60933599"],
            "Partida": ["S122460", "S122460", "S122460", "S122900"],
            "Cliente": ["TOWERSPORT S.A.C.", "TOWERSPORT S.A.C.", "TOWERSPORT S.A.C.", "OTRO CLIENTE"],
            "Peso_KG": [18.95, 19.10, 20.00, 15.50],
            "Metros": [87.65, 88.00, 91.20, 70.00]
        }
        df_rollos = pd.DataFrame(datos_demo)
    else:
        df_rollos = st.session_state.df_rollos_master.copy()

    df_rollos['Peso_KG'] = pd.to_numeric(df_rollos['Peso_KG'], errors='coerce').fillna(0)
    df_despachos = cargar_historial()

    # PESTAÑAS DEL SISTEMA
    tab_escaneo, tab_historial = st.tabs(["📷 Escaneo de Rollos", "📋 Historial Completo en Nube"])

    with tab_escaneo:
        # FORMULARIO DE ESCANEO CONTINUO
        with st.form(key="form_escaneo", clear_on_submit=True):
            codigo_input = st.text_input("📷 ESCANEE CÓDIGO DEL ROLLO:").strip().upper()
            btn_validar = st.form_submit_button("Validar Rollo", use_container_width=True)

        if btn_validar and codigo_input:
            operario_logueado = st.session_state.nombre_operario
            rollo_encontrado = df_rollos[df_rollos["Codigo_Barras"] == codigo_input]

            if rollo_encontrado.empty:
                st.error(f"🚨 ¡ALARMA! El código '{codigo_input}' NO EXISTE en la base de datos.")
                reproducir_sonido_alarma()
                partida_reg = st.session_state.partida_activa if st.session_state.partida_activa else "N/A"
                registrar_despacho(operario_logueado, partida_reg, codigo_input, "N/A", 0, "RECHAZADO", "Código inexistente")
            else:
                partida_real = str(rollo_encontrado.iloc[0]["Partida"])
                cliente = str(rollo_encontrado.iloc[0]["Cliente"])
                peso = float(rollo_encontrado.iloc[0]["Peso_KG"])

                # Auto-detección de la primera partida
                if st.session_state.partida_activa is None:
                    st.session_state.partida_activa = partida_real
                    st.info(f"🎯 Partida detectada automáticamente: **{partida_real}** ({cliente})")

                if partida_real != st.session_state.partida_activa:
                    st.error(f"🚨 ¡ALARMA! Este rollo pertenece a la PARTIDA {partida_real}. Se está cargando la PARTIDA {st.session_state.partida_activa}.")
                    reproducir_sonido_alarma()
                    registrar_despacho(operario_logueado, st.session_state.partida_activa, codigo_input, cliente, peso, "RECHAZADO", f"Pertenece a {partida_real}")
                else:
                    despachos_partida = df_despachos[
                        (df_despachos["Partida_Cargada"] == st.session_state.partida_activa) & 
                        (df_despachos["Estado"] == "CARGADO_OK")
                    ]
                    
                    if codigo_input in despachos_partida["Codigo_Rollo"].tolist():
                        st.warning(f"⚠️ El rollo '{codigo_input}' YA FUE ESCANEADO previamente.")
                        reproducir_sonido_alarma()
                        registrar_despacho(operario_logueado, st.session_state.partida_activa, codigo_input, cliente, peso, "DUPLICADO", "Rollo repetido")
                    else:
                        st.success(f"✅ ¡ROLLO OK! Código: {codigo_input} | Peso: {peso} KG")
                        reproducir_sonido_ok()
                        registrar_despacho(operario_logueado, st.session_state.partida_activa, codigo_input, cliente, peso, "CARGADO_OK", "Validado")
                        st.rerun()

        # CHECKLIST Y PROGRESO EN TIEMPO REAL
        if st.session_state.partida_activa:
            st.markdown(f"### 📋 Checklist de Carga - Partida: `{st.session_state.partida_activa}`")
            
            if st.button("🏁 Finalizar Carga de Partida Actual"):
                st.session_state.partida_activa = None
                st.rerun()

            rollos_totales = df_rollos[df_rollos["Partida"] == st.session_state.partida_activa].copy()
            
            despachos_ok = df_despachos[
                (df_despachos["Partida_Cargada"] == st.session_state.partida_activa) & 
                (df_despachos["Estado"] == "CARGADO_OK")
            ]["Codigo_Rollo"].tolist()

            rollos_totales["Estado_Carga"] = rollos_totales["Codigo_Barras"].apply(
                lambda x: "🟢 CARGADO" if x in despachos_ok else "🔴 PENDIENTE"
            )

            cargados_cnt = len(despachos_ok)
            total_cnt = len(rollos_totales)
            kilos_cargados = rollos_totales[rollos_totales["Estado_Carga"] == "🟢 CARGADO"]["Peso_KG"].sum()
            kilos_totales = rollos_totales["Peso_KG"].sum()

            m1, m2 = st.columns(2)
            m1.metric("Progreso de Rollos", f"{cargados_cnt} de {total_cnt}")
            m2.metric("Kilos Cargados", f"{kilos_cargados:.2f} / {kilos_totales:.2f} KG")

            st.dataframe(
                rollos_totales[["Codigo_Barras", "Cliente", "Peso_KG", "Metros", "Estado_Carga"]],
                use_container_width=True
            )

    with tab_historial:
        st.header("📊 Bitácora de Despachos Registrados en Supabase")
        if not df_despachos.empty:
            st.dataframe(df_despachos, use_container_width=True)
            
            # Exportar a Excel
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Historial Despachos"
            
            # Cabeceras
            ws.append(list(df_despachos.columns))
            for row in df_despachos.itertuples(index=False):
                ws.append(list(row))
                
            wb.save(output)
            excel_bytes = output.getvalue()
            
            st.download_button(
                label="📥 Descargar Historial Completo en Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"Reporte_Despachos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("No hay registros almacenados en la nube.")
