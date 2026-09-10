#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
 SISTEMA DE CONTROL LOGÍSTICO INTEGRAL DE COMBUSTIBLE - OPERACIÓN RS1
============================================================================
Aplicación autosuficiente en Streamlit + SQLite para la gestión integral de
combustible de una operación minera: recepción de cisterna, despacho a
maquinaria, préstamos/salidas auxiliares, consumo vehicular y kardex
valorizado con Costo Promedio Ponderado (CPP).

Ejecución:
    streamlit run app.py

Autor: Desarrollado como Desarrollador Senior Full-Stack (Logística Minera)
============================================================================
"""

import io
import os
import sqlite3
import datetime as dt
from contextlib import closing

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

try:
    import openpyxl
    from openpyxl.utils.datetime import to_excel
except ImportError:  # pragma: no cover
    openpyxl = None


# ============================================================================
# 1. CONFIGURACIÓN GLOBAL Y CONSTANTES DE NEGOCIO
# ============================================================================

DB_PATH = "kardex_rs1.db"

STOCK_CRITICO_GLN = 600        # Umbral crítico permanente (galones)
STOCK_ALERTA_GLN = 900         # Umbral de alerta temprana (galones)

# Umbrales de desviación de rendimiento (Gln/Hr) respecto a la referencia
DESV_AMARILLA = 0.15           # +/- 15% -> Amarillo
DESV_ROJA = 0.35                # +/- 35% -> Rojo

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Setiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}

LOGO_PATH = "logo.png"  # Coloca aquí tu logo (mismo nombre, en la raíz del proyecto)

st.set_page_config(
    page_title="Control de Combustible RS1",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded",
)


def aplicar_estilo_profesional():
    """Inyecta una tipografía corporativa (Google Fonts 'Inter') y pequeños
    ajustes visuales sobre el tema azul/gris definido en .streamlit/config.toml."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', 'Segoe UI', sans-serif !important;
        }
        h1, h2, h3, h4, h5, h6 {
            font-family: 'Inter', 'Segoe UI', sans-serif !important;
            font-weight: 700 !important;
            color: #1F4E8C;
            letter-spacing: -0.3px;
        }
        [data-testid="stSidebar"] {
            background-color: #EEF1F5;
            border-right: 1px solid #D6DCE5;
        }
        [data-testid="stMetricValue"] {
            color: #1F4E8C;
            font-weight: 700;
        }
        .stButton > button, .stFormSubmitButton > button {
            border-radius: 6px;
            font-weight: 600;
        }
        [data-testid="stTabs"] button [data-testid="stMarkdownContainer"] p {
            font-weight: 600;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# 2. CAPA DE BASE DE DATOS (SQLite)
# ============================================================================

@st.cache_resource(show_spinner=False)
def get_conn():
    """Conexión única y persistente a SQLite, compartida entre reruns."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    init_db(conn)
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    """Crea el esquema relacional si no existe (idempotente)."""
    cur = conn.cursor()

    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS tanques (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            tipo TEXT DEFAULT 'Auxiliar',
            capacidad_gln REAL,
            stock_minimo_gln REAL,
            ubicacion TEXT
        );

        CREATE TABLE IF NOT EXISTS parametros_equipo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            modelo TEXT,
            consumo_est_gln_hr REAL,
            ubicacion TEXT
        );

        CREATE TABLE IF NOT EXISTS ingresos_cisterna (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            factura_vale TEXT,
            proveedor TEXT,
            tanque_destino TEXT,
            galones REAL NOT NULL,
            precio_unitario REAL NOT NULL,
            importe_total REAL NOT NULL,
            forma_pago TEXT,
            soporte TEXT,
            observaciones TEXT,
            origen TEXT DEFAULT 'manual'
        );

        CREATE TABLE IF NOT EXISTS despacho_maquinaria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            equipo TEXT NOT NULL,
            horometro_anterior REAL,
            horometro_actual REAL,
            cambio_horometro INTEGER DEFAULT 0,
            horas_operadas REAL,
            galones REAL NOT NULL,
            rendimiento REAL,
            abastecedor TEXT,
            dni TEXT,
            observaciones TEXT,
            estado TEXT DEFAULT 'OK',
            origen TEXT DEFAULT 'manual'
        );

        CREATE TABLE IF NOT EXISTS prestamos_salidas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            tipo_operacion TEXT,
            destino TEXT,
            cantidad_original REAL,
            unidad TEXT DEFAULT 'Galones',
            galones_equivalentes REAL NOT NULL,
            responsable TEXT,
            receptor TEXT,
            vale TEXT,
            observaciones TEXT,
            origen TEXT DEFAULT 'manual'
        );

        CREATE TABLE IF NOT EXISTS consumo_vehicular (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            placa TEXT,
            n_vale TEXT,
            galones REAL NOT NULL,
            precio REAL,
            total REAL,
            origen TEXT DEFAULT 'manual'
        );

        CREATE TABLE IF NOT EXISTS kardex_diario (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT UNIQUE NOT NULL,
            concepto TEXT,
            ingresos_gln REAL DEFAULT 0,
            salidas_maquinaria_gln REAL DEFAULT 0,
            salidas_prestamos_gln REAL DEFAULT 0,
            salidas_vehicular_gln REAL DEFAULT 0,
            total_salidas_gln REAL DEFAULT 0,
            stock_gln REAL DEFAULT 0,
            nivel_stock TEXT,
            costo_promedio_gln REAL DEFAULT 0,
            valor_ingreso_dia REAL DEFAULT 0,
            saldo_valorizado REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS app_metadata (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        """
    )
    conn.commit()

    # Semilla mínima de tanques conocidos (se amplía automáticamente con uso)
    tanques_base = [
        ("Cisterna Principal Mina", "Cisterna", 1500, 300, "Patio Superficie"),
        ("Tanque Estacionario RS1 - 1", "IBC/Estacionario", 300, 60, "Zona de Tanques RS1"),
        ("Tanque Estacionario RS1 - 2", "IBC/Estacionario", 200, 40, "Zona de Tanques RS1"),
    ]
    for nombre, tipo, cap, minimo, ubic in tanques_base:
        cur.execute(
            "INSERT OR IGNORE INTO tanques (nombre, tipo, capacidad_gln, stock_minimo_gln, ubicacion) "
            "VALUES (?,?,?,?,?)",
            (nombre, tipo, cap, minimo, ubic),
        )
    conn.commit()


def get_or_create_tanque(conn, nombre: str, tipo: str = "Auxiliar"):
    """Devuelve el tanque; si no existe (nombre libre proveniente de un
    vale/factura histórico) lo registra automáticamente."""
    if not nombre:
        nombre = "Sin Especificar"
    cur = conn.cursor()
    cur.execute("SELECT id FROM tanques WHERE nombre = ?", (nombre,))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute(
        "INSERT INTO tanques (nombre, tipo) VALUES (?, ?)", (nombre, tipo)
    )
    conn.commit()
    return cur.lastrowid


def get_or_create_parametro_equipo(conn, nombre: str, consumo_est=None):
    cur = conn.cursor()
    cur.execute("SELECT id FROM parametros_equipo WHERE nombre = ?", (nombre,))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute(
        "INSERT INTO parametros_equipo (nombre, consumo_est_gln_hr) VALUES (?, ?)",
        (nombre, consumo_est),
    )
    conn.commit()
    return cur.lastrowid


# ============================================================================
# 3. UTILIDADES DE LECTURA / SANEAMIENTO DE DATOS
# ============================================================================

def _to_float(value, default=None):
    """Convierte celdas heterogéneas de Excel ('-', None, texto, número) a
    float, sin romper la importación."""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        v = value.strip().replace(",", "")
        if v in ("", "-", "S/", "S/P", "N/A"):
            return default
        try:
            return float(v)
        except ValueError:
            return default
    return default


def _to_date_str(value):
    """Normaliza fechas de Excel (datetime, date o texto) a 'YYYY-MM-DD'."""
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, str):
        v = value.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return dt.datetime.strptime(v, fmt).date().isoformat()
            except ValueError:
                continue
        return None
    return None


EQUIPO_ALIAS = {
    # Normaliza variantes de nombre del mismo equipo detectadas en el
    # historial fuente, para no romper la continuidad de horómetros.
    "generador gsw-6 perkins": "Generador Perkins GSW-6",
    "generador perkins gsw-6": "Generador Perkins GSW-6",
    "compresor sullair 375-h": "Compresor Sullair 375-H",
}


def normalizar_equipo(nombre: str) -> str:
    if not nombre:
        return nombre
    limpio = " ".join(str(nombre).strip().split())
    return EQUIPO_ALIAS.get(limpio.lower(), limpio)


def _horometro_a_numero(value):
    """El archivo fuente guarda la lectura de horómetro en celdas con
    formato de hora, por lo que openpyxl la entrega como datetime. El
    número de serie de Excel subyacente ES la lectura real (en horas)."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, dt.datetime):
        try:
            return round(float(to_excel(value)), 2)
        except Exception:
            return None
    if isinstance(value, str):
        return _to_float(value)
    return None


def mes_label(fecha_iso: str) -> str:
    if not fecha_iso:
        return ""
    d = dt.date.fromisoformat(fecha_iso)
    return f"{MESES_ES[d.month]} {d.year}"


# ============================================================================
# 4. MIGRACIÓN DESDE LOS ARCHIVOS EXCEL ORIGINALES
# ============================================================================

def _leer_ingresos_cisterna(wb):
    """Hoja 'Ingresos_Cisterna' del archivo CONTROL_LOGISTICO."""
    if "Ingresos_Cisterna" not in wb.sheetnames:
        return []
    ws = wb["Ingresos_Cisterna"]
    filas = []
    for row in ws.iter_rows(min_row=6, values_only=True):
        fecha = row[1] if len(row) > 1 else None
        if not isinstance(fecha, (dt.datetime, dt.date)):
            continue  # salta encabezados repetidos / fila TOTAL GENERAL
        galones = _to_float(row[6]) or 0.0
        precio = _to_float(row[7]) or 0.0
        importe = _to_float(row[8])
        if importe is None:
            importe = round(galones * precio, 2)
        filas.append(
            dict(
                fecha=_to_date_str(fecha),
                factura_vale=row[3],
                proveedor=row[4],
                tanque_destino=row[5] or "Sin Especificar",
                galones=galones,
                precio_unitario=precio,
                importe_total=importe,
                forma_pago=row[9],
                soporte=row[10],
                observaciones=row[11] if len(row) > 11 else None,
            )
        )
    return filas


def _leer_despacho_maquinaria(wb):
    """Hoja 'Despacho_Maquinaria'. Recalculamos horómetro anterior/actual y
    horas operadas de forma propia y consistente, recorriendo por equipo en
    orden cronológico (misma lógica que usará la app en adelante)."""
    if "Despacho_Maquinaria" not in wb.sheetnames:
        return []
    ws = wb["Despacho_Maquinaria"]
    crudos = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        fecha = row[1] if len(row) > 1 else None
        equipo = row[3] if len(row) > 3 else None
        if not isinstance(fecha, (dt.datetime, dt.date)) or not equipo:
            continue
        horometro = _horometro_a_numero(row[4])
        galones = _to_float(row[6]) or 0.0
        crudos.append(
            dict(
                fecha=_to_date_str(fecha),
                orden=len(crudos),
                equipo=normalizar_equipo(equipo),
                horometro=horometro,
                galones=galones,
                abastecedor=row[8],
                dni=row[9],
                observaciones=row[10] if len(row) > 10 else None,
                estado=row[11] if len(row) > 11 else "OK",
            )
        )
    # Orden cronológico estable por equipo
    crudos.sort(key=lambda r: (r["fecha"] or "", r["orden"]))

    ultimo_horometro = {}
    filas = []
    for r in crudos:
        equipo = r["equipo"]
        anterior = ultimo_horometro.get(equipo)
        actual = r["horometro"]
        cambio = 0
        horas = None
        if actual is not None:
            if anterior is not None:
                if actual <= anterior:
                    cambio = 1
                    horas = None  # requiere validación manual en uso futuro
                else:
                    horas = round(actual - anterior, 2)
            ultimo_horometro[equipo] = actual
        rendimiento = None
        if horas and horas > 0:
            rendimiento = round(r["galones"] / horas, 2)
        filas.append(
            dict(
                fecha=r["fecha"],
                equipo=equipo,
                horometro_anterior=anterior,
                horometro_actual=actual,
                cambio_horometro=cambio,
                horas_operadas=horas,
                galones=r["galones"],
                rendimiento=rendimiento,
                abastecedor=r["abastecedor"],
                dni=str(r["dni"]) if r["dni"] is not None else None,
                observaciones=r["observaciones"],
                estado=r["estado"] or "OK",
            )
        )
    return filas


def _leer_prestamos(wb):
    if "Prestamos" not in wb.sheetnames:
        return []
    ws = wb["Prestamos"]
    filas = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        fecha = row[1] if len(row) > 1 else None
        if not isinstance(fecha, (dt.datetime, dt.date)):
            continue
        gal_eq = _to_float(row[7]) or 0.0
        filas.append(
            dict(
                fecha=_to_date_str(fecha),
                tipo_operacion=row[3] or "Préstamo de Combustible",
                destino=row[4],
                cantidad_original=_to_float(row[5]),
                unidad=row[6] or "Galones",
                galones_equivalentes=gal_eq,
                responsable=row[8],
                receptor=row[9],
                vale=None,
                observaciones=row[10] if len(row) > 10 else None,
            )
        )
    return filas


def _leer_consumo_vehicular(wb):
    """Hoja 'BD' del archivo Comsumo_de_combustible."""
    if "BD" not in wb.sheetnames:
        return []
    ws = wb["BD"]
    filas = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        placa = row[1] if len(row) > 1 else None
        fecha = row[2] if len(row) > 2 else None
        if not placa or not isinstance(fecha, (dt.datetime, dt.date)):
            continue
        galones = _to_float(row[3]) or 0.0
        precio = _to_float(row[4])
        total = _to_float(row[5])
        if total is None and precio is not None:
            total = round(galones * precio, 2)
        filas.append(
            dict(
                fecha=_to_date_str(fecha),
                placa=str(placa).strip(),
                n_vale=str(row[6]) if len(row) > 6 and row[6] is not None else None,
                galones=galones,
                precio=precio,
                total=total,
            )
        )
    return filas


def _leer_parametros(wb):
    """Hoja 'Parametros' del archivo CONTROL_LOGISTICO -> alimenta tanques y
    parametros_equipo (consumo estimado usado como referencia de rendimiento)."""
    if "Parametros" not in wb.sheetnames:
        return [], []
    ws = wb["Parametros"]
    tanques, equipos = [], []
    for row in ws.iter_rows(min_row=6, values_only=True):
        nombre = row[1] if len(row) > 1 else None
        if not nombre:
            continue
        capacidad = _to_float(row[3])
        stock_min = _to_float(row[4])
        consumo_est = _to_float(row[5])
        ubicacion = row[6]
        low = str(nombre).lower()
        if "tanque" in low or "cisterna" in low:
            tanques.append(
                dict(nombre=str(nombre).strip(), capacidad_gln=capacidad,
                     stock_minimo_gln=stock_min, ubicacion=ubicacion)
            )
        else:
            equipos.append(
                dict(nombre=str(nombre).strip(), consumo_est_gln_hr=consumo_est,
                     ubicacion=ubicacion)
            )
    return tanques, equipos


def migrar_desde_excel(conn, archivo_control_bytes=None, archivo_consumo_bytes=None) -> dict:
    """Importa el historial completo desde los dos libros Excel originales.
    Es segura de re-ejecutar: primero limpia únicamente los registros con
    origen='migracion', preservando cualquier registro cargado manualmente
    desde la propia aplicación."""
    if openpyxl is None:
        raise RuntimeError("openpyxl no está instalado en el entorno.")

    resumen = {
        "tanques": 0, "parametros_equipo": 0, "ingresos_cisterna": 0,
        "despacho_maquinaria": 0, "prestamos_salidas": 0, "consumo_vehicular": 0,
    }
    cur = conn.cursor()

    # Limpieza de una migración previa (idempotencia)
    for tabla in ["ingresos_cisterna", "despacho_maquinaria", "prestamos_salidas", "consumo_vehicular"]:
        cur.execute(f"DELETE FROM {tabla} WHERE origen = 'migracion'")
    conn.commit()

    if archivo_control_bytes:
        wb1 = openpyxl.load_workbook(io.BytesIO(archivo_control_bytes), data_only=True)

        tanques, equipos = _leer_parametros(wb1)
        for t in tanques:
            cur.execute(
                "INSERT INTO tanques (nombre, tipo, capacidad_gln, stock_minimo_gln, ubicacion) "
                "VALUES (?,?,?,?,?) "
                "ON CONFLICT(nombre) DO UPDATE SET capacidad_gln=excluded.capacidad_gln, "
                "stock_minimo_gln=excluded.stock_minimo_gln, ubicacion=excluded.ubicacion",
                (t["nombre"], "Tanque", t["capacidad_gln"], t["stock_minimo_gln"], t["ubicacion"]),
            )
            resumen["tanques"] += 1
        for e in equipos:
            cur.execute(
                "INSERT INTO parametros_equipo (nombre, consumo_est_gln_hr, ubicacion) VALUES (?,?,?) "
                "ON CONFLICT(nombre) DO UPDATE SET consumo_est_gln_hr=excluded.consumo_est_gln_hr, "
                "ubicacion=excluded.ubicacion",
                (e["nombre"], e["consumo_est_gln_hr"], e["ubicacion"]),
            )
            resumen["parametros_equipo"] += 1
        conn.commit()

        for f in _leer_ingresos_cisterna(wb1):
            get_or_create_tanque(conn, f["tanque_destino"])
            cur.execute(
                """INSERT INTO ingresos_cisterna
                   (fecha, factura_vale, proveedor, tanque_destino, galones, precio_unitario,
                    importe_total, forma_pago, soporte, observaciones, origen)
                   VALUES (?,?,?,?,?,?,?,?,?,?, 'migracion')""",
                (f["fecha"], f["factura_vale"], f["proveedor"], f["tanque_destino"],
                 f["galones"], f["precio_unitario"], f["importe_total"], f["forma_pago"],
                 f["soporte"], f["observaciones"]),
            )
            resumen["ingresos_cisterna"] += 1

        for f in _leer_despacho_maquinaria(wb1):
            get_or_create_parametro_equipo(conn, f["equipo"])
            cur.execute(
                """INSERT INTO despacho_maquinaria
                   (fecha, equipo, horometro_anterior, horometro_actual, cambio_horometro,
                    horas_operadas, galones, rendimiento, abastecedor, dni, observaciones,
                    estado, origen)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?, 'migracion')""",
                (f["fecha"], f["equipo"], f["horometro_anterior"], f["horometro_actual"],
                 f["cambio_horometro"], f["horas_operadas"], f["galones"], f["rendimiento"],
                 f["abastecedor"], f["dni"], f["observaciones"], f["estado"]),
            )
            resumen["despacho_maquinaria"] += 1

        for f in _leer_prestamos(wb1):
            cur.execute(
                """INSERT INTO prestamos_salidas
                   (fecha, tipo_operacion, destino, cantidad_original, unidad,
                    galones_equivalentes, responsable, receptor, vale, observaciones, origen)
                   VALUES (?,?,?,?,?,?,?,?,?,?, 'migracion')""",
                (f["fecha"], f["tipo_operacion"], f["destino"], f["cantidad_original"],
                 f["unidad"], f["galones_equivalentes"], f["responsable"], f["receptor"],
                 f["vale"], f["observaciones"]),
            )
            resumen["prestamos_salidas"] += 1
        conn.commit()

    if archivo_consumo_bytes:
        wb2 = openpyxl.load_workbook(io.BytesIO(archivo_consumo_bytes), data_only=True)
        for f in _leer_consumo_vehicular(wb2):
            cur.execute(
                """INSERT INTO consumo_vehicular (fecha, placa, n_vale, galones, precio, total, origen)
                   VALUES (?,?,?,?,?,?, 'migracion')""",
                (f["fecha"], f["placa"], f["n_vale"], f["galones"], f["precio"], f["total"]),
            )
            resumen["consumo_vehicular"] += 1
        conn.commit()

    cur.execute(
        "INSERT INTO app_metadata (key, value) VALUES ('ultima_migracion', ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (dt.datetime.now().isoformat(timespec="seconds"),),
    )
    conn.commit()

    recompute_kardex(conn)
    return resumen


# ============================================================================
# 5. LÓGICA DE NEGOCIO: HORÓMETROS, RENDIMIENTO Y KARDEX CPP
# ============================================================================

def obtener_ultimo_horometro(conn, equipo: str):
    row = conn.execute(
        "SELECT horometro_actual FROM despacho_maquinaria "
        "WHERE equipo = ? AND horometro_actual IS NOT NULL "
        "ORDER BY fecha DESC, id DESC LIMIT 1",
        (equipo,),
    ).fetchone()
    return row[0] if row else None


def obtener_rendimiento_referencia(conn, equipo: str):
    """Rendimiento (Gln/Hr) de referencia: promedio histórico del propio
    equipo; si no hay historial, cae al consumo estimado de parámetros."""
    row = conn.execute(
        "SELECT AVG(rendimiento) FROM despacho_maquinaria "
        "WHERE equipo = ? AND rendimiento IS NOT NULL AND rendimiento > 0",
        (equipo,),
    ).fetchone()
    if row and row[0]:
        return round(row[0], 2)
    row = conn.execute(
        "SELECT consumo_est_gln_hr FROM parametros_equipo WHERE nombre = ?",
        (equipo,),
    ).fetchone()
    return round(row[0], 2) if row and row[0] else None


def clasificar_rendimiento(rendimiento, referencia):
    """Devuelve ('Verde'|'Amarillo'|'Rojo'|'Sin datos', desviación%)."""
    if rendimiento is None:
        return "Sin datos", None
    if not referencia:
        return "Sin datos", None
    desv = (rendimiento - referencia) / referencia
    if abs(desv) <= DESV_AMARILLA:
        return "Verde", desv
    if abs(desv) <= DESV_ROJA:
        return "Amarillo", desv
    return "Rojo", desv


def lista_equipos(conn):
    rows = conn.execute(
        "SELECT DISTINCT equipo FROM despacho_maquinaria "
        "UNION SELECT nombre FROM parametros_equipo ORDER BY 1"
    ).fetchall()
    return [r[0] for r in rows if r[0]]


def lista_tanques(conn):
    rows = conn.execute("SELECT nombre FROM tanques ORDER BY nombre").fetchall()
    return [r[0] for r in rows]


def recompute_kardex(conn) -> None:
    """Recalcula íntegramente el Kardex diario valorizado con Costo Promedio
    Ponderado (CPP), recorriendo cronológicamente TODOS los movimientos:
    ingresos (aumentan stock y recalculan el CPP) y salidas de maquinaria,
    préstamos y consumo vehicular (disminuyen stock al CPP vigente, sin
    alterarlo). Se consolida un registro por día."""
    cur = conn.cursor()

    ingresos = cur.execute(
        "SELECT fecha, galones, precio_unitario FROM ingresos_cisterna ORDER BY fecha, id"
    ).fetchall()
    desp = cur.execute(
        "SELECT fecha, galones FROM despacho_maquinaria ORDER BY fecha, id"
    ).fetchall()
    prest = cur.execute(
        "SELECT fecha, galones_equivalentes FROM prestamos_salidas ORDER BY fecha, id"
    ).fetchall()
    veh = cur.execute(
        "SELECT fecha, galones FROM consumo_vehicular ORDER BY fecha, id"
    ).fetchall()

    fechas = sorted(set(
        [r[0] for r in ingresos] + [r[0] for r in desp] +
        [r[0] for r in prest] + [r[0] for r in veh]
    ))

    stock = 0.0
    cpp = 0.0
    filas = []
    for fecha in fechas:
        ing_dia = [g for f, g, p in ingresos if f == fecha]
        ing_dia_full = [(g, p) for f, g, p in ingresos if f == fecha]
        desp_dia = sum(g for f, g in desp if f == fecha)
        prest_dia = sum(g for f, g in prest if f == fecha)
        veh_dia = sum(g for f, g in veh if f == fecha)

        ingreso_gln_dia = sum(ing_dia)
        valor_ingreso_dia = 0.0

        # 1) Procesar ingresos del día -> recalcula CPP (promedio ponderado)
        for galones, precio in ing_dia_full:
            if galones <= 0:
                continue
            valor_actual = stock * cpp
            valor_nuevo = galones * (precio or 0)
            stock_nuevo = stock + galones
            cpp = (valor_actual + valor_nuevo) / stock_nuevo if stock_nuevo > 0 else 0
            stock = stock_nuevo
            valor_ingreso_dia += valor_nuevo

        # 2) Procesar salidas del día (no alteran el CPP, solo el stock)
        total_salidas = desp_dia + prest_dia + veh_dia
        stock -= total_salidas
        if stock < 0:
            stock = 0.0  # blindaje: el stock físico nunca es negativo

        saldo_valorizado = round(stock * cpp, 2)

        if stock < STOCK_CRITICO_GLN:
            nivel = "Crítico"
        elif stock < STOCK_ALERTA_GLN:
            nivel = "Moderado"
        else:
            nivel = "Óptimo"

        concepto = []
        if ingreso_gln_dia > 0:
            concepto.append("Llegada Cisterna")
        if total_salidas > 0:
            concepto.append("Salidas de Combustible")
        if not concepto:
            concepto.append("Sin movimientos")

        filas.append((
            fecha, " / ".join(concepto), round(ingreso_gln_dia, 2),
            round(desp_dia, 2), round(prest_dia, 2), round(veh_dia, 2),
            round(total_salidas, 2), round(stock, 2), nivel,
            round(cpp, 4), round(valor_ingreso_dia, 2), saldo_valorizado,
        ))

    cur.execute("DELETE FROM kardex_diario")
    cur.executemany(
        """INSERT INTO kardex_diario
           (fecha, concepto, ingresos_gln, salidas_maquinaria_gln, salidas_prestamos_gln,
            salidas_vehicular_gln, total_salidas_gln, stock_gln, nivel_stock,
            costo_promedio_gln, valor_ingreso_dia, saldo_valorizado)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        filas,
    )
    conn.commit()


def stock_y_cpp_actual(conn):
    row = conn.execute(
        "SELECT stock_gln, costo_promedio_gln, saldo_valorizado FROM kardex_diario "
        "ORDER BY fecha DESC LIMIT 1"
    ).fetchone()
    return row if row else (0.0, 0.0, 0.0)


# ============================================================================
# 6. OPERACIONES DE INSERCIÓN (uso manual desde los formularios)
# ============================================================================

def insertar_ingreso(conn, **kw):
    conn.execute(
        """INSERT INTO ingresos_cisterna
           (fecha, factura_vale, proveedor, tanque_destino, galones, precio_unitario,
            importe_total, forma_pago, soporte, observaciones, origen)
           VALUES (?,?,?,?,?,?,?,?,?,?, 'manual')""",
        (kw["fecha"], kw["factura_vale"], kw["proveedor"], kw["tanque_destino"],
         kw["galones"], kw["precio_unitario"], kw["importe_total"], kw["forma_pago"],
         kw["soporte"], kw["observaciones"]),
    )
    conn.commit()
    recompute_kardex(conn)


def insertar_despacho(conn, **kw):
    conn.execute(
        """INSERT INTO despacho_maquinaria
           (fecha, equipo, horometro_anterior, horometro_actual, cambio_horometro,
            horas_operadas, galones, rendimiento, abastecedor, dni, observaciones,
            estado, origen)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?, 'manual')""",
        (kw["fecha"], kw["equipo"], kw["horometro_anterior"], kw["horometro_actual"],
         kw["cambio_horometro"], kw["horas_operadas"], kw["galones"], kw["rendimiento"],
         kw["abastecedor"], kw["dni"], kw["observaciones"], kw["estado"]),
    )
    conn.commit()
    recompute_kardex(conn)


def insertar_prestamo(conn, **kw):
    conn.execute(
        """INSERT INTO prestamos_salidas
           (fecha, tipo_operacion, destino, cantidad_original, unidad,
            galones_equivalentes, responsable, receptor, vale, observaciones, origen)
           VALUES (?,?,?,?,?,?,?,?,?,?, 'manual')""",
        (kw["fecha"], kw["tipo_operacion"], kw["destino"], kw["cantidad_original"],
         kw["unidad"], kw["galones_equivalentes"], kw["responsable"], kw["receptor"],
         kw["vale"], kw["observaciones"]),
    )
    conn.commit()
    recompute_kardex(conn)


def insertar_consumo_vehicular(conn, **kw):
    conn.execute(
        """INSERT INTO consumo_vehicular (fecha, placa, n_vale, galones, precio, total, origen)
           VALUES (?,?,?,?,?,?, 'manual')""",
        (kw["fecha"], kw["placa"], kw["n_vale"], kw["galones"], kw["precio"], kw["total"]),
    )
    conn.commit()
    recompute_kardex(conn)


# ============================================================================
# 7. COMPONENTES DE INTERFAZ (Streamlit)
# ============================================================================

def badge_nivel(nivel: str) -> str:
    colores = {"Óptimo": "🟢", "Moderado": "🟡", "Crítico": "🔴"}
    return f"{colores.get(nivel, '⚪')} {nivel}"


def render_sidebar(conn):
    st.sidebar.title("⛽ RS1 · Control de Combustible")
    st.sidebar.caption("Sistema integral de gestión y kardex de combustible")

    stock, cpp, saldo = stock_y_cpp_actual(conn)
    st.sidebar.metric("Stock actual (Gln)", f"{stock:,.1f}")
    if stock < STOCK_CRITICO_GLN:
        st.sidebar.error(f"🔴 STOCK CRÍTICO: por debajo de {STOCK_CRITICO_GLN} Gln")
    elif stock < STOCK_ALERTA_GLN:
        st.sidebar.warning("🟡 Nivel de stock moderado, planifique reabastecimiento")
    else:
        st.sidebar.success("🟢 Nivel de stock óptimo")

    st.sidebar.divider()
    with st.sidebar.expander("📥 Migración de Datos Históricos", expanded=False):
        st.caption(
            "Carga los dos archivos Excel originales para poblar automáticamente "
            "toda la base de datos histórica."
        )
        f1 = st.file_uploader("CONTROL_LOGISTICO_COMBUSTIBLE_RS1.xlsx", type=["xlsx"], key="f1")
        f2 = st.file_uploader("Consumo_de_combustible_RS1.xlsx", type=["xlsx"], key="f2")
        if st.button("🚀 Importar / Reimportar Historial", use_container_width=True):
            if not f1 and not f2:
                st.error("Sube al menos uno de los dos archivos.")
            else:
                with st.spinner("Migrando datos históricos a la base de datos..."):
                    resumen = migrar_desde_excel(
                        conn,
                        archivo_control_bytes=f1.getvalue() if f1 else None,
                        archivo_consumo_bytes=f2.getvalue() if f2 else None,
                    )
                st.success("Migración completada correctamente.")
                st.json(resumen)
                st.rerun()

        meta = conn.execute(
            "SELECT value FROM app_metadata WHERE key='ultima_migracion'"
        ).fetchone()
        if meta:
            st.caption(f"Última migración: {meta[0]}")

    st.sidebar.divider()
    st.sidebar.caption("Desarrollado para el control logístico de la operación minera RS1.")


def tab_dashboard(conn):
    st.subheader("📊 Dashboard Ejecutivo")

    stock, cpp, saldo = stock_y_cpp_actual(conn)
    df_kardex = pd.read_sql_query(
        "SELECT * FROM kardex_diario ORDER BY fecha", conn
    )

    hoy = dt.date.today()
    inicio_mes = hoy.replace(day=1).isoformat()
    consumo_mes = conn.execute(
        "SELECT "
        "  COALESCE((SELECT SUM(galones) FROM despacho_maquinaria WHERE fecha >= ?),0) + "
        "  COALESCE((SELECT SUM(galones_equivalentes) FROM prestamos_salidas WHERE fecha >= ?),0) + "
        "  COALESCE((SELECT SUM(galones) FROM consumo_vehicular WHERE fecha >= ?),0)",
        (inicio_mes, inicio_mes, inicio_mes),
    ).fetchone()[0] or 0.0

    # Autonomía: promedio de consumo diario de los últimos 30 días con movimiento
    df_recientes = df_kardex[df_kardex["fecha"] >= (hoy - dt.timedelta(days=30)).isoformat()]
    consumo_prom_dia = df_recientes["total_salidas_gln"].mean() if not df_recientes.empty else 0
    autonomia_dias = (stock / consumo_prom_dia) if consumo_prom_dia and consumo_prom_dia > 0 else None

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Stock en Tanque (Gln)", f"{stock:,.1f}",
              delta=badge_nivel(df_kardex.iloc[-1]["nivel_stock"]) if not df_kardex.empty else "")
    c2.metric("Saldo Valorizado (S/)", f"S/ {saldo:,.2f}", help=f"CPP vigente: S/ {cpp:,.4f} / Gln")
    c3.metric("Autonomía Estimada", f"{autonomia_dias:,.1f} días" if autonomia_dias else "Sin datos",
              help="Basado en el consumo promedio diario de los últimos 30 días")
    c4.metric("Consumo del Mes (Gln)", f"{consumo_mes:,.1f}")

    if stock < STOCK_CRITICO_GLN:
        st.error(f"🔴 **ALERTA DE STOCK CRÍTICO** — El stock ({stock:,.1f} Gln) está por debajo "
                  f"del umbral mínimo de {STOCK_CRITICO_GLN} Gln. Coordine reabastecimiento urgente.")

    st.divider()

    col1, col2 = st.columns([1.3, 1])
    with col1:
        st.markdown("##### Curva de Stock en Tanque (Gln)")
        if not df_kardex.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_kardex["fecha"], y=df_kardex["stock_gln"],
                                      mode="lines", name="Stock (Gln)", line=dict(color="#1f77b4", width=2),
                                      fill="tozeroy"))
            fig.add_hline(y=STOCK_CRITICO_GLN, line_dash="dash", line_color="red",
                          annotation_text="Umbral Crítico")
            fig.add_hline(y=STOCK_ALERTA_GLN, line_dash="dot", line_color="orange",
                          annotation_text="Umbral de Alerta")
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10),
                               yaxis_title="Galones", xaxis_title=None)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Aún no hay movimientos registrados. Importa el historial o registra un ingreso.")

    with col2:
        st.markdown("##### Saldo Valorizado (S/)")
        if not df_kardex.empty:
            fig2 = px.area(df_kardex, x="fecha", y="saldo_valorizado")
            fig2.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10),
                                yaxis_title="Soles (S/)", xaxis_title=None)
            fig2.update_traces(line_color="#2ca02c")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Sin datos.")

    st.divider()
    col3, col4 = st.columns(2)
    with col3:
        st.markdown("##### Consumo Diario por Máquina (Gln)")
        df_desp = pd.read_sql_query(
            "SELECT fecha, equipo, galones FROM despacho_maquinaria ORDER BY fecha", conn
        )
        if not df_desp.empty:
            fig3 = px.bar(df_desp, x="fecha", y="galones", color="equipo", barmode="stack")
            fig3.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10), legend_title="Equipo")
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("Sin despachos registrados aún.")

    with col4:
        st.markdown("##### Comparativa Mensual: Ingresos vs. Salidas (Gln)")
        if not df_kardex.empty:
            df_kardex["mes"] = df_kardex["fecha"].apply(mes_label)
            df_mensual = df_kardex.groupby("mes", sort=False).agg(
                Ingresos=("ingresos_gln", "sum"), Salidas=("total_salidas_gln", "sum")
            ).reset_index()
            fig4 = go.Figure()
            fig4.add_bar(x=df_mensual["mes"], y=df_mensual["Ingresos"], name="Ingresos", marker_color="#2ca02c")
            fig4.add_bar(x=df_mensual["mes"], y=df_mensual["Salidas"], name="Salidas", marker_color="#d62728")
            fig4.update_layout(barmode="group", height=380, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.info("Sin datos.")


def tab_despacho(conn):
    st.subheader("⛽ Registro de Despacho Diario a Maquinaria")
    st.caption("Interior mina y superficie — con validación en vivo de horómetro y rendimiento.")

    equipos = lista_equipos(conn)
    col_eq, col_new = st.columns([2, 1])
    with col_eq:
        modo_nuevo = st.checkbox("Registrar equipo nuevo (no está en la lista)")
        if modo_nuevo or not equipos:
            equipo = normalizar_equipo(st.text_input("Nombre del nuevo equipo"))
        else:
            equipo = st.selectbox("Equipo / Máquina", equipos)

    ultimo_horometro = obtener_ultimo_horometro(conn, equipo) if equipo else None
    referencia_rend = obtener_rendimiento_referencia(conn, equipo) if equipo else None

    info_cols = st.columns(3)
    info_cols[0].info(f"🕐 Último horómetro registrado: **"
                       f"{ultimo_horometro:,.2f} Hr**" if ultimo_horometro is not None else
                       "🕐 Sin lecturas previas para este equipo.")
    info_cols[1].info(f"⚙️ Rendimiento de referencia: **{referencia_rend:,.2f} Gln/Hr**"
                       if referencia_rend else "⚙️ Sin rendimiento de referencia aún.")

    with st.form("form_despacho", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        fecha = c1.date_input("Fecha de Despacho", value=dt.date.today())
        galones = c2.number_input("Cantidad Suministrada (Gln)", min_value=0.0, step=0.1)
        abastecedor = c3.text_input("Abastecedor (Nombre completo)")

        c4, c5, c6 = st.columns(3)
        horometro_actual = c4.number_input(
            "Lectura de Horómetro Actual (Hr)", min_value=0.0, step=0.1,
            value=float(ultimo_horometro) if ultimo_horometro else 0.0,
        )
        dni = c5.text_input("DNI del Abastecedor")
        tanque_origen = c6.selectbox("Tanque / Cisterna de Origen", lista_tanques(conn))

        cambio_horometro_manual = False
        horas_manual = None
        if ultimo_horometro is not None and horometro_actual <= ultimo_horometro:
            st.warning(
                "⚠️ El horómetro ingresado es **menor o igual** al último registrado "
                f"({ultimo_horometro:,.2f} Hr). Esto solo es válido si hubo un "
                "**cambio o mantenimiento del horómetro**."
            )
            cambio_horometro_manual = st.checkbox(
                "Confirmo Cambio / Mantenimiento de Horómetro (evita horas negativas)"
            )
            horas_manual = st.number_input(
                "Horas Operadas (ingreso manual, ya que el horómetro fue reiniciado)",
                min_value=0.0, step=0.1,
            )

        observaciones = st.text_area("Observaciones Operativas")
        submitted = st.form_submit_button("💾 Registrar Despacho", use_container_width=True)

    if submitted:
        if not equipo:
            st.error("Debes indicar el equipo.")
            return
        if galones <= 0:
            st.error("Los galones deben ser mayores a cero.")
            return

        horas_operadas = None
        cambio_flag = 0
        if ultimo_horometro is not None:
            if horometro_actual <= ultimo_horometro:
                if not cambio_horometro_manual:
                    st.error(
                        "🚫 No se puede registrar: el horómetro es menor o igual al anterior. "
                        "Marca la casilla de confirmación de cambio/mantenimiento para continuar."
                    )
                    return
                cambio_flag = 1
                horas_operadas = horas_manual
            else:
                horas_operadas = round(horometro_actual - ultimo_horometro, 2)

        rendimiento = round(galones / horas_operadas, 2) if horas_operadas and horas_operadas > 0 else None
        nivel_rend, desv = clasificar_rendimiento(rendimiento, referencia_rend)

        estado = "OK"
        obs_final = observaciones or ""
        if nivel_rend == "Rojo":
            estado = "Alerta"
            obs_final += " [ALERTA: consumo anómalo / posible fuga]"
        elif nivel_rend == "Amarillo":
            obs_final += " [Desviación de rendimiento respecto a la referencia]"

        insertar_despacho(
            conn, fecha=fecha.isoformat(), equipo=equipo,
            horometro_anterior=ultimo_horometro, horometro_actual=horometro_actual,
            cambio_horometro=cambio_flag, horas_operadas=horas_operadas, galones=galones,
            rendimiento=rendimiento, abastecedor=abastecedor, dni=dni,
            observaciones=obs_final.strip(), estado=estado,
        )
        get_or_create_parametro_equipo(conn, equipo)
        get_or_create_tanque(conn, tanque_origen)

        if nivel_rend == "Rojo":
            st.error(f"🔴 Despacho registrado, pero el rendimiento ({rendimiento} Gln/Hr) se desvía "
                      f"{desv:+.0%} de la referencia. Posible consumo anómalo o fuga — revisar el equipo.")
        elif nivel_rend == "Amarillo":
            st.warning(f"🟡 Despacho registrado con desviación de rendimiento de {desv:+.0%} "
                        "respecto a la referencia histórica.")
        else:
            st.success("🟢 Despacho registrado correctamente.")
        st.rerun()

    st.divider()
    st.markdown("##### Últimos Despachos Registrados")
    df = pd.read_sql_query(
        "SELECT fecha AS Fecha, equipo AS Equipo, horometro_anterior AS 'Horómetro Ant.', "
        "horometro_actual AS 'Horómetro Act.', horas_operadas AS 'Horas', galones AS Galones, "
        "rendimiento AS 'Rend. Gln/Hr', abastecedor AS Abastecedor, estado AS Estado, "
        "observaciones AS Observaciones "
        "FROM despacho_maquinaria ORDER BY fecha DESC, id DESC LIMIT 100", conn
    )
    def _color_estado(row):
        color = "#ffe6e6" if row["Estado"] == "Alerta" else ""
        return [f"background-color: {color}"] * len(row)
    if not df.empty:
        st.dataframe(df.style.apply(_color_estado, axis=1), use_container_width=True, hide_index=True)
    else:
        st.info("Aún no hay despachos registrados.")


def tab_recepcion(conn):
    st.subheader("🚚 Recepción de Cisterna & Proveedores")
    st.caption("Registro de compras, facturas/vales e ingresos a tanques y cisterna.")

    tanques = lista_tanques(conn)
    with st.form("form_ingreso", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        fecha = c1.date_input("Fecha de Llegada", value=dt.date.today())
        factura = c2.text_input("N° Factura / Vale")
        proveedor = c3.text_input("Proveedor", value="Servicios Multiples El Panita E.I.R.L.")

        c4, c5, c6 = st.columns(3)
        destino_existente = c4.selectbox("Tanque Destino", tanques + ["+ Nuevo tanque..."])
        if destino_existente == "+ Nuevo tanque...":
            destino = c4.text_input("Nombre del nuevo tanque")
        else:
            destino = destino_existente
        galones = c5.number_input("Cantidad (Gln)", min_value=0.0, step=0.1)
        precio_unit = c6.number_input("P. Unitario (S/)", min_value=0.0, step=0.01, format="%.2f")

        importe = round(galones * precio_unit, 2)
        st.metric("Importe Total (S/)", f"S/ {importe:,.2f}")

        c7, c8 = st.columns(2)
        forma_pago = c7.selectbox("Forma de Pago", ["Efectivo", "Transferencia BCP", "Crédito / Efectivo", "Otro"])
        soporte = c8.text_input("Soporte / N° de Operación")
        observaciones = st.text_area("Observaciones")

        submitted = st.form_submit_button("💾 Registrar Ingreso a Tanque", use_container_width=True)

    if submitted:
        if galones <= 0 or precio_unit <= 0:
            st.error("Galones y precio unitario deben ser mayores a cero.")
        elif not destino:
            st.error("Debes indicar el tanque destino.")
        else:
            insertar_ingreso(
                conn, fecha=fecha.isoformat(), factura_vale=factura, proveedor=proveedor,
                tanque_destino=destino, galones=galones, precio_unitario=precio_unit,
                importe_total=importe, forma_pago=forma_pago, soporte=soporte,
                observaciones=observaciones,
            )
            get_or_create_tanque(conn, destino)
            st.success(f"✅ Ingreso registrado: {galones:,.1f} Gln al tanque '{destino}'. "
                        "El Costo Promedio Ponderado (CPP) del Kardex fue actualizado.")
            st.rerun()

    st.divider()
    st.markdown("##### Historial de Ingresos a Cisterna")
    df = pd.read_sql_query(
        "SELECT fecha AS Fecha, factura_vale AS 'Factura/Vale', proveedor AS Proveedor, "
        "tanque_destino AS 'Tanque Destino', galones AS Galones, precio_unitario AS 'P.Unit S/', "
        "importe_total AS 'Importe S/', forma_pago AS 'Forma de Pago' "
        "FROM ingresos_cisterna ORDER BY fecha DESC, id DESC LIMIT 100", conn
    )
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.caption(f"Total histórico: {df['Galones'].sum():,.1f} Gln  ·  "
                    f"S/ {df['Importe S/'].sum():,.2f}")
    else:
        st.info("Aún no hay ingresos registrados.")


def tab_prestamos_flota(conn):
    st.subheader("📋 Préstamos & Flota Vehicular")

    sub1, sub2 = st.tabs(["🔁 Préstamos y Salidas Auxiliares", "🚗 Consumo por Placa"])

    with sub1:
        with st.form("form_prestamo", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            fecha = c1.date_input("Fecha", value=dt.date.today())
            tipo = c2.selectbox("Tipo de Operación",
                                 ["Préstamo de Combustible", "Préstamo Directo", "Otro"])
            destino = c3.text_input("Destino / Labor / Beneficiario")

            c4, c5, c6 = st.columns(3)
            cantidad = c4.number_input("Cantidad Original", min_value=0.0, step=0.1)
            unidad = c5.selectbox("Unidad", ["Galones", "Litros"])
            vale = c6.text_input("N° Vale")

            galones_eq = cantidad if unidad == "Galones" else round(cantidad / 3.78541, 2)
            st.metric("Galones Equivalentes (Gln)", f"{galones_eq:,.2f}")

            c7, c8 = st.columns(2)
            responsable = c7.text_input("Responsable de Entrega")
            receptor = c8.text_input("Receptor / Autorizado")
            observaciones = st.text_area("Observaciones / Estado")

            submitted = st.form_submit_button("💾 Registrar Préstamo / Salida", use_container_width=True)

        if submitted:
            if galones_eq <= 0:
                st.error("La cantidad debe ser mayor a cero.")
            else:
                insertar_prestamo(
                    conn, fecha=fecha.isoformat(), tipo_operacion=tipo, destino=destino,
                    cantidad_original=cantidad, unidad=unidad, galones_equivalentes=galones_eq,
                    responsable=responsable, receptor=receptor, vale=vale,
                    observaciones=observaciones,
                )
                st.success("✅ Préstamo / salida registrada. El Kardex se actualizó.")
                st.rerun()

        st.divider()
        df = pd.read_sql_query(
            "SELECT fecha AS Fecha, tipo_operacion AS Tipo, destino AS Destino, "
            "cantidad_original AS Cantidad, unidad AS Unidad, "
            "galones_equivalentes AS 'Gln Equiv.', responsable AS Responsable, "
            "receptor AS Receptor, observaciones AS Observaciones "
            "FROM prestamos_salidas ORDER BY fecha DESC, id DESC LIMIT 100", conn
        )
        if not df.empty:
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("Sin préstamos registrados.")

    with sub2:
        with st.form("form_vehicular", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            fecha = c1.date_input("Fecha", value=dt.date.today(), key="fecha_veh")
            placa = c2.text_input("Placa")
            n_vale = c3.text_input("N° Vale")

            c4, c5 = st.columns(2)
            galones = c4.number_input("Galones", min_value=0.0, step=0.1, key="gal_veh")
            precio = c5.number_input("Precio (S/)", min_value=0.0, step=0.01, format="%.2f", key="precio_veh")
            total = round(galones * precio, 2)
            st.metric("Total (S/)", f"S/ {total:,.2f}")

            submitted_v = st.form_submit_button("💾 Registrar Consumo Vehicular", use_container_width=True)

        if submitted_v:
            if galones <= 0 or not placa:
                st.error("Debes indicar la placa y una cantidad de galones mayor a cero.")
            else:
                insertar_consumo_vehicular(
                    conn, fecha=fecha.isoformat(), placa=placa.upper(), n_vale=n_vale,
                    galones=galones, precio=precio, total=total,
                )
                st.success("✅ Consumo vehicular registrado.")
                st.rerun()

        st.divider()
        df_v = pd.read_sql_query(
            "SELECT fecha AS Fecha, placa AS Placa, n_vale AS 'N° Vale', galones AS Galones, "
            "precio AS Precio, total AS 'Total S/' FROM consumo_vehicular "
            "ORDER BY fecha DESC, id DESC LIMIT 150", conn
        )
        if not df_v.empty:
            st.dataframe(df_v, use_container_width=True, hide_index=True)
            resumen_placa = df_v.groupby("Placa", as_index=False)["Galones"].sum().sort_values(
                "Galones", ascending=False
            )
            st.markdown("###### Consumo Acumulado por Placa (Gln)")
            fig = px.bar(resumen_placa, x="Placa", y="Galones")
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sin consumo vehicular registrado.")


def tab_kardex(conn):
    st.subheader("📑 Kardex Oficial & Reportes")

    if st.button("🔄 Recalcular Kardex"):
        with st.spinner("Recalculando Kardex con Costo Promedio Ponderado..."):
            recompute_kardex(conn)
        st.success("Kardex recalculado.")

    df = pd.read_sql_query("SELECT * FROM kardex_diario ORDER BY fecha", conn)
    if df.empty:
        st.info("Aún no hay movimientos para construir el Kardex.")
        return

    df["fecha_dt"] = pd.to_datetime(df["fecha"])
    min_f, max_f = df["fecha_dt"].min().date(), df["fecha_dt"].max().date()

    c1, c2 = st.columns(2)
    rango = c1.date_input("Rango de Fechas", value=(min_f, max_f), min_value=min_f, max_value=max_f)
    meses_disp = sorted(df["fecha"].apply(mes_label).unique().tolist())
    mes_sel = c2.multiselect("Filtrar por Mes", meses_disp, default=[])

    df_f = df.copy()
    if isinstance(rango, tuple) and len(rango) == 2:
        df_f = df_f[(df_f["fecha_dt"].dt.date >= rango[0]) & (df_f["fecha_dt"].dt.date <= rango[1])]
    if mes_sel:
        df_f = df_f[df_f["fecha"].apply(mes_label).isin(mes_sel)]

    df_show = df_f.rename(columns={
        "fecha": "Fecha", "concepto": "Concepto / Movimiento",
        "ingresos_gln": "Ingresos Cisterna (Gln)", "salidas_maquinaria_gln": "Salidas Maquinaria (Gln)",
        "salidas_prestamos_gln": "Salidas Préstamos (Gln)", "salidas_vehicular_gln": "Salidas Vehicular (Gln)",
        "total_salidas_gln": "Total Salidas (Gln)", "stock_gln": "Stock en Tanque (Gln)",
        "nivel_stock": "Nivel de Stock", "costo_promedio_gln": "CPP (S/ / Gln)",
        "valor_ingreso_dia": "Valor Ingreso del Día (S/)", "saldo_valorizado": "Saldo Valorizado (S/)",
    })[[
        "Fecha", "Concepto / Movimiento", "Ingresos Cisterna (Gln)", "Salidas Maquinaria (Gln)",
        "Salidas Préstamos (Gln)", "Salidas Vehicular (Gln)", "Total Salidas (Gln)",
        "Stock en Tanque (Gln)", "Nivel de Stock", "CPP (S/ / Gln)",
        "Valor Ingreso del Día (S/)", "Saldo Valorizado (S/)",
    ]]

    st.dataframe(df_show, use_container_width=True, hide_index=True, height=460)

    c3, c4, c5 = st.columns(3)
    c3.metric("Ingresos del período (Gln)", f"{df_f['ingresos_gln'].sum():,.1f}")
    c4.metric("Salidas del período (Gln)", f"{df_f['total_salidas_gln'].sum():,.1f}")
    c5.metric("Saldo Valorizado Final (S/)", f"S/ {df_f['saldo_valorizado'].iloc[-1]:,.2f}" if not df_f.empty else "S/ 0.00")

    st.divider()
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_show.to_excel(writer, index=False, sheet_name="Kardex Valorizado")
        ws = writer.sheets["Kardex Valorizado"]
        for col_cells in ws.columns:
            length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 12), 40)
    buffer.seek(0)

    st.download_button(
        "⬇️ Descargar Kardex Valorizado a Excel (.xlsx)",
        data=buffer,
        file_name=f"Kardex_Valorizado_RS1_{dt.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# ============================================================================
# 8. PUNTO DE ENTRADA
# ============================================================================

def main():
    conn = get_conn()
    aplicar_estilo_profesional()

    if os.path.exists(LOGO_PATH):
        st.logo(LOGO_PATH, size="large")

    st.title("⛽ Control Logístico Integral de Combustible — Operación RS1")
    st.caption(
        "Gestión de kardex valorizado, despacho a maquinaria, recepción de cisterna, "
        "préstamos y flota vehicular, con Costo Promedio Ponderado (CPP) en tiempo real."
    )

    render_sidebar(conn)

    tabs = st.tabs([
        "📊 Dashboard Ejecutivo",
        "⛽ Despacho Diario",
        "🚚 Recepción de Cisterna & Proveedores",
        "📋 Préstamos & Flota Vehicular",
        "📑 Kardex Oficial & Reportes",
    ])
    with tabs[0]:
        tab_dashboard(conn)
    with tabs[1]:
        tab_despacho(conn)
    with tabs[2]:
        tab_recepcion(conn)
    with tabs[3]:
        tab_prestamos_flota(conn)
    with tabs[4]:
        tab_kardex(conn)


if __name__ == "__main__":
    main()
