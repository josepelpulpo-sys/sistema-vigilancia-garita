import datetime
import io
import sqlite3
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import streamlit as st

# Configuración de la página web
st.set_page_config(
    page_title="Control de Portería & Recepción de Telas",
    layout="wide",
    page_icon="🚛",
)


# -------------------------------------------------------------------
# BASE DE DATOS SQLITE (HISTORIAL Y BITÁCORA DE MOVIMIENTOS)
# -------------------------------------------------------------------
def conectar_db():
    return sqlite3.connect("registro_porteria.db")


def inicializar_db():
    conn = conectar_db()
    cursor = conn.cursor()

    # 1. Tabla de Movimientos / Bitácora Completa de Eventos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historial_movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_evento TEXT,
            hora_evento TEXT,
            tipo_operacion TEXT,
            num_guia_entrada TEXT,
            num_guia_salida TEXT DEFAULT '-',
            num_partida TEXT DEFAULT '-',
            cliente TEXT,
            tipo_tela TEXT,
            rollos_movimiento INTEGER,
            peso_kg_movimiento REAL,
            rollos_saldo_planta INTEGER,
            chofer TEXT,
            dni_chofer TEXT,
            empresa_transporte TEXT,
            placa TEXT,
            vigilante TEXT
        )
    """)

    # 2. Tabla de Stock Actual / Estado de Guías en Planta
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS stock_telas_planta (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            num_guia_entrada TEXT,
            num_partida TEXT,
            cliente TEXT,
            tipo_tela TEXT,
            rollos_actuales INTEGER,
            peso_kg_inicial REAL,
            fecha_ingreso TEXT,
            chofer_ingreso TEXT,
            dni_chofer_ingreso TEXT,
            transporte_ingreso TEXT,
            placa_ingreso TEXT
        )
    """)

    conn.commit()
    conn.close()


inicializar_db()


# -------------------------------------------------------------------
# SESIÓN DE OPERADOR / VIGILANTE
# -------------------------------------------------------------------
if "vigilante" not in st.session_state:
    st.session_state["vigilante"] = None

if not st.session_state["vigilante"]:
    st.markdown("## 🛡️ Control de Portería & Recepción de Telas")
    st.info("Por favor, ingrese el nombre del vigilante u operador de turno para iniciar.")

    with st.form(key="form_vigilante"):
        nombre_input = st.text_input("Nombre del Vigilante:")
        submit_vigilante = st.form_submit_button("Ingresar al Sistema")

        if submit_vigilante:
            if nombre_input.strip():
                st.session_state["vigilante"] = nombre_input.strip()
                st.success(f"Bienvenido/a, **{st.session_state['vigilante']}**.")
                st.rerun()
            else:
                st.error("⚠️ Debe ingresar su nombre para continuar.")
    st.stop()


# -------------------------------------------------------------------
# BARRA LATERAL
# -------------------------------------------------------------------
with st.sidebar:
    st.title("👤 Turno Activo")
    st.write(f"**Vigilante:** {st.session_state['vigilante']}")
    st.write(f"**Fecha:** {datetime.datetime.now().strftime('%Y-%m-%d')}")
    st.markdown("---")
    if st.button("🔴 Cambiar de Vigilante / Salir"):
        st.session_state["vigilante"] = None
        st.rerun()


# -------------------------------------------------------------------
# GENERACIÓN DE EXCEL 100% COMPATIBLE Y SIN ERRORES DE REPARACIÓN
# -------------------------------------------------------------------
def generar_excel_profesional():
    conn = conectar_db()
    query = """
        SELECT 
            id AS 'N° Movimiento',
            fecha_evento AS 'Fecha Evento',
            hora_evento AS 'Hora Evento',
            tipo_operacion AS 'Tipo de Operación',
            num_guia_entrada AS 'N° Guía Entrada',
            num_guia_salida AS 'N° Guía Salida',
            num_partida AS 'N° Partida',
            cliente AS 'Cliente',
            tipo_tela AS 'Tipo de Tela',
            rollos_movimiento AS 'Rollos Movidis / Retirados',
            peso_kg_movimiento AS 'Peso Movidi (Kg)',
            rollos_saldo_planta AS 'Saldo Restante en Planta',
            chofer AS 'Chófer Registrado',
            dni_chofer AS 'DNI Chófer',
            empresa_transporte AS 'Empresa Transportista',
            placa AS 'Placa Vehículo',
            vigilante AS 'Vigilante en Turno'
        FROM historial_movimientos
        ORDER BY id DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    # Limpiar valores nulos para evitar corrupción de celdas
    df = df.fillna("-")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Historial Completo"

    # Estilos de formato
    navy_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
    )
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    # Poblar filas de forma segura mediante dataframe_to_rows
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(row)
        for c_idx in range(1, len(row) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            
            if r_idx == 1:
                cell.fill = navy_header
                cell.font = white_bold
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = Font(name="Calibri", size=10)
                if r_idx % 2 == 0:
                    cell.fill = zebra_fill
                
                if c_idx in [1, 2, 3, 4, 5, 6, 7, 13, 14, 16, 17]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx in [10, 11, 12]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

    # Autoajuste de anchos de columna
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# -------------------------------------------------------------------
# INTERFAZ PRINCIPAL
# -------------------------------------------------------------------
st.title("🚛 Control de Portería - Bitácora Completa de Movimientos")
st.caption(f"Vigilante en turno: **{st.session_state['vigilante']}**")
st.markdown("---")

tab1, tab2, tab3 = st.tabs([
    "📥 1. Registro de Portería (Dejar / Recoger)",
    "📦 2. Stock Actual en Almacén",
    "📋 3. Historial de Todos los Movimientos"
])

# -------------------------------------------------------------------
# PESTAÑA 1: REGISTRO DE MOVIMIENTO (ENTRADA / RECOJO)
# -------------------------------------------------------------------
with tab1:
    st.subheader("1. Datos Completos del Vehículo, Chófer y Transporte")
    
    col_v1, col_v2, col_v3, col_v4, col_v5 = st.columns([2.5, 2, 2.5, 2, 3])
    with col_v1:
        chofer_nom = st.text_input("Nombre del Chófer *").strip()
    with col_v2:
        dni_chofer_in = st.text_input("DNI / Documento del Chófer *").strip()
    with col_v3:
        empresa_trans = st.text_input("Empresa de Transporte / Tercero *", placeholder="Ej. Servitrans").strip()
    with col_v4:
        placa_veh = st.text_input("Placa del Carro/Camión *").upper().strip()
    with col_v5:
        operacion = st.radio("¿Qué acción realiza el vehículo? *", ["Dejar Tela (Ingreso a Planta)", "Recoger Tela (Despacho de Planta)"])

    st.markdown("---")

    # OPERACIÓN A: DEJAR TELA
    if operacion == "Dejar Tela (Ingreso a Planta)":
        st.subheader("2. Registro de Guía de Entrada y Cliente")
        
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            num_guia_e = st.text_input("N° Guía de Remisión de Entrada *").strip()
        with col_g2:
            cliente_e = st.text_input("Cliente de la Tela *").strip()

        st.markdown("##### 📝 Lista de telas que vienen en esta Guía:")
        
        if "telas_dejar" not in st.session_state:
            st.session_state["telas_dejar"] = []

        with st.form(key="form_add_tela_dejar", clear_on_submit=True):
            col_d1, col_d2, col_d3, col_d4 = st.columns([2, 3, 2, 2])
            with col_d1:
                partida_t = st.text_input("N° Partida (Opcional)", placeholder="Ej. P-102")
            with col_d2:
                tipo_t = st.text_input("Tipo de Tela (ej. Jersey 18/1, Rib Veteado 24/1)")
            with col_d3:
                cant_r = st.number_input("Cantidad de Rollos", min_value=1, step=1)
            with col_d4:
                peso_k = st.number_input("Peso de Entrada (Kg)", min_value=0.1, step=0.1)
            
            btn_add = st.form_submit_button("➕ Agregar Tela a la Lista")
            if btn_add:
                if tipo_t.strip():
                    st.session_state["telas_dejar"].append({
                        "num_partida": partida_t.strip() if partida_t.strip() else "-",
                        "tipo_tela": tipo_t.strip(),
                        "rollos": cant_r,
                        "peso_kg": peso_k
                    })
                    st.success(f"Agregado: {tipo_t} - {cant_r} rollos ({peso_k} Kg)")
                else:
                    st.warning("Escriba el tipo de tela.")

        if st.session_state["telas_dejar"]:
            st.dataframe(pd.DataFrame(st.session_state["telas_dejar"]), use_container_width=True)
            if st.button("🗑️ Limpiar Telas"):
                st.session_state["telas_dejar"] = []
                st.rerun()

        if st.button("💾 Registrar Ingreso de Tela", type="primary"):
            if not chofer_nom or not dni_chofer_in or not empresa_trans or not placa_veh or not num_guia_e or not cliente_e:
                st.error("⚠️ Complete todos los campos obligatorios (*): Chófer, DNI, Empresa de Transporte, Placa, Guía de Entrada y Cliente.")
            elif not st.session_state["telas_dejar"]:
                st.error("⚠️ Debe agregar al menos un tipo de tela a la lista.")
            else:
                fecha_f = datetime.datetime.now().strftime("%Y-%m-%d")
                hora_f = datetime.datetime.now().strftime("%H:%M:%S")

                conn = conectar_db()
                cursor = conn.cursor()

                for item in st.session_state["telas_dejar"]:
                    cursor.execute("""
                        INSERT INTO historial_movimientos (
                            fecha_evento, hora_evento, tipo_operacion, num_guia_entrada, num_guia_salida,
                            num_partida, cliente, tipo_tela, rollos_movimiento, peso_kg_movimiento,
                            rollos_saldo_planta, chofer, dni_chofer, empresa_transporte, placa, vigilante
                        ) VALUES (?, ?, 'Ingreso a Planta', ?, '-', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        fecha_f, hora_f, num_guia_e, item["num_partida"], cliente_e, item["tipo_tela"],
                        item["rollos"], item["peso_kg"], item["rollos"], chofer_nom, dni_chofer_in,
                        empresa_trans, placa_veh, st.session_state["vigilante"]
                    ))

                    cursor.execute("""
                        INSERT INTO stock_telas_planta (
                            num_guia_entrada, num_partida, cliente, tipo_tela, rollos_actuales,
                            peso_kg_inicial, fecha_ingreso, chofer_ingreso, dni_chofer_ingreso,
                            transporte_ingreso, placa_ingreso
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        num_guia_e, item["num_partida"], cliente_e, item["tipo_tela"], item["rollos"],
                        item["peso_kg"], fecha_f, chofer_nom, dni_chofer_in, empresa_trans, placa_veh
                    ))

                conn.commit()
                conn.close()

                st.success(f"✅ Ingreso registrado con éxito. Guía **{num_guia_e}** almacenada en historial y planta.")
                st.session_state["telas_dejar"] = []
                st.rerun()

    # OPERACIÓN B: RECOGER TELA (CON LIMPIEZA AUTOMÁTICA DE ESTADO)
    else:
        st.subheader("2. Selección de Guía / Tela a Recoger")
        
        conn = conectar_db()
        df_disponibles = pd.read_sql_query("""
            SELECT num_guia_entrada, cliente, fecha_ingreso,
                   SUM(rollos_actuales) as total_rollos, SUM(peso_kg_inicial) as total_peso
            FROM stock_telas_planta
            WHERE rollos_actuales > 0
            GROUP BY num_guia_entrada
        """, conn)
        conn.close()

        if not df_disponibles.empty:
            opciones_guias = {
                f"Guía Ent.: {row['num_guia_entrada']} | Cliente: {row['cliente']} ({row['total_rollos']} Rollos dispon. en planta)": row['num_guia_entrada']
                for _, row in df_disponibles.iterrows()
            }
            
            guia_sel = st.selectbox("Seleccione la Guía de Entrada sobre la que retirará telas:", list(opciones_guias.keys()))
            num_guia_retirar = opciones_guias[guia_sel]

            conn = conectar_db()
            df_det_orig = pd.read_sql_query(
                "SELECT id, num_partida, cliente, tipo_tela, rollos_actuales, peso_kg_inicial FROM stock_telas_planta WHERE num_guia_entrada = ? AND rollos_actuales > 0",
                conn, params=(num_guia_retirar,)
            )
            conn.close()

            # Reiniciar la memoria de selección si se cambia de guía o si la lista quedó vacía
            if ("telas_despacho_ids" not in st.session_state 
                or st.session_state.get("guia_actual_sel") != num_guia_retirar 
                or not st.session_state["telas_despacho_ids"]):
                
                st.session_state["guia_actual_sel"] = num_guia_retirar
                st.session_state["telas_despacho_ids"] = list(df_det_orig["id"])

            st.markdown("---")
            st.write("##### 📋 Indique los datos de salida para este movimiento:")

            datos_salida_telas = []
            
            for _, row in df_det_orig.iterrows():
                item_id = row["id"]
                rollos_disponibles = int(row['rollos_actuales'])
                
                if item_id in st.session_state["telas_despacho_ids"]:
                    st.markdown(f"🧵 **Tela:** `{row['tipo_tela']}` | **Stock actual en almacén:** **{rollos_disponibles} rollos**")
                    
                    col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns([2, 2, 2, 2, 1])
                    with col_s1:
                        guia_sal_ind = st.text_input("N° Guía Salida *", key=f"gs_{item_id}", placeholder="Ej. GS-501")
                    with col_s2:
                        partida_sal_ind = st.text_input("N° Partida", value=str(row['num_partida']), key=f"part_{item_id}")
                    with col_s3:
                        r_out = st.number_input("Rollos que Retira hoy", min_value=1, max_value=rollos_disponibles, value=rollos_disponibles, key=f"r_{item_id}")
                    with col_s4:
                        p_out = st.number_input("Peso Salida (Kg)", min_value=0.1, value=float(row['peso_kg_inicial']), step=0.1, key=f"p_{item_id}")
                    with col_s5:
                        st.write("")
                        st.write("")
                        if st.button("❌ Quitar", key=f"btn_del_{item_id}"):
                            st.session_state["telas_despacho_ids"].remove(item_id)
                            st.rerun()

                    datos_salida_telas.append({
                        "id": item_id,
                        "cliente": row["cliente"],
                        "tipo_tela": row["tipo_tela"],
                        "guia_salida": guia_sal_ind.strip(),
                        "num_partida": partida_sal_ind.strip(),
                        "rollos_salida": r_out,
                        "rollos_disponibles_antes": rollos_disponibles,
                        "peso_salida": p_out
                    })
                    st.markdown("---")

            if st.button("🚪 Confirmar Recojo y Registrar Movimiento en Historial", type="primary"):
                guias_vacias = [d for d in datos_salida_telas if not d["guia_salida"]]
                
                if not chofer_nom or not dni_chofer_in or not empresa_trans or not placa_veh:
                    st.error("⚠️ Complete los campos obligatorios del vehículo y chófer: Nombre, DNI, Empresa de Transporte y Placa.")
                elif guias_vacias:
                    st.error("⚠️ Debe ingresar el N° de Guía de Salida para cada tela que va a despachar.")
                elif not datos_salida_telas:
                    st.error("⚠️ Ha quitado todas las telas. Seleccione al menos una tela para despachar.")
                else:
                    fecha_f = datetime.datetime.now().strftime("%Y-%m-%d")
                    hora_f = datetime.datetime.now().strftime("%H:%M:%S")

                    conn = conectar_db()
                    cursor = conn.cursor()

                    for d_out in datos_salida_telas:
                        rollos_restantes = d_out["rollos_disponibles_antes"] - d_out["rollos_salida"]
                        tipo_mov = "Despacho Final" if rollos_restantes == 0 else "Despacho Parcial"

                        cursor.execute("""
                            INSERT INTO historial_movimientos (
                                fecha_evento, hora_evento, tipo_operacion, num_guia_entrada, num_guia_salida,
                                num_partida, cliente, tipo_tela, rollos_movimiento, peso_kg_movimiento,
                                rollos_saldo_planta, chofer, dni_chofer, empresa_transporte, placa, vigilante
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            fecha_f, hora_f, tipo_mov, num_guia_retirar, d_out["guia_salida"],
                            d_out["num_partida"], d_out["cliente"], d_out["tipo_tela"], d_out["rollos_salida"],
                            d_out["peso_salida"], rollos_restantes, chofer_nom, dni_chofer_in,
                            empresa_trans, placa_veh, st.session_state["vigilante"]
                        ))

                        cursor.execute("""
                            UPDATE stock_telas_planta
                            SET rollos_actuales = ?, num_partida = ?
                            WHERE id = ?
                        """, (rollos_restantes, d_out["num_partida"], d_out["id"]))

                    conn.commit()
                    conn.close()

                    # Limpieza explícita de variables de sesión
                    if "telas_despacho_ids" in st.session_state:
                        del st.session_state["telas_despacho_ids"]
                    if "guia_actual_sel" in st.session_state:
                        del st.session_state["guia_actual_sel"]

                    st.success(f"✅ Nuevo movimiento registrado en el historial. Saldo en planta actualizado.")
                    st.rerun()
        else:
            st.info("No hay telas almacenadas en planta actualmente para retirar.")

# -------------------------------------------------------------------
# PESTAÑA 2: TELAS ACTUALMENTE EN PLANTA (STOCK REAL)
# -------------------------------------------------------------------
with tab2:
    st.subheader("📦 Stock Actual de Telas Almacenadas en Empresa")
    
    conn = conectar_db()
    df_stock = pd.read_sql_query("""
        SELECT 
            num_guia_entrada AS 'N° Guía Entrada',
            num_partida AS 'N° Partida',
            cliente AS 'Cliente',
            tipo_tela AS 'Tipo de Tela',
            rollos_actuales AS 'Rollos Restantes en Planta',
            peso_kg_inicial AS 'Peso Inicial Entrante (Kg)',
            fecha_ingreso AS 'Fecha Ingreso',
            chofer_ingreso AS 'Chófer que la Dejó',
            dni_chofer_ingreso AS 'DNI Chófer',
            transporte_ingreso AS 'Empresa Transportista'
        FROM stock_telas_planta
        WHERE rollos_actuales > 0
        ORDER BY id DESC
    """, conn)
    conn.close()

    if not df_stock.empty:
        st.dataframe(df_stock, use_container_width=True)
    else:
        st.info("El almacén no tiene telas registradas en planta actualmente.")

# -------------------------------------------------------------------
# PESTAÑA 3: BITÁCORA / HISTORIAL COMPLETO DE TODOS LOS MOVIMIENTOS
# -------------------------------------------------------------------
with tab3:
    st.subheader("📋 Historial Completo de Movimientos (Bitácora de Entradas y Salidas)")

    conn = conectar_db()
    df_todo = pd.read_sql_query("""
        SELECT 
            id AS 'N° Mov.',
            fecha_evento AS 'Fecha',
            hora_evento AS 'Hora',
            tipo_operacion AS 'Operación',
            num_guia_entrada AS 'Guía Ent.',
            num_guia_salida AS 'Guía Sal.',
            num_partida AS 'N° Partida',
            cliente AS 'Cliente',
            tipo_tela AS 'Tipo Tela',
            rollos_movimiento AS 'Rollos Movidis',
            peso_kg_movimiento AS 'Peso (Kg)',
            rollos_saldo_planta AS 'Saldo Restante Planta',
            chofer AS 'Chófer',
            dni_chofer AS 'DNI Chófer',
            empresa_transporte AS 'Empresa Transp.',
            placa AS 'Placa',
            vigilante AS 'Vigilante'
        FROM historial_movimientos
        ORDER BY id DESC
    """, conn)
    conn.close()

    if not df_todo.empty:
        excel_bytes = generar_excel_profesional()
        st.download_button(
            label="📊 Descargar Historial Completo en Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Historial_Movimientos_Porteria_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown("---")
        st.dataframe(df_todo, use_container_width=True)
    else:
        st.write("Aún no hay movimientos registrados en la bitácora.")