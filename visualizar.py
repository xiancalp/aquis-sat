"""
Panel en tempo real para sensores.xlsx
Executar con:  streamlit run visualizar.py
Actualízase automaticamente cada 5 segundos.
"""

import time
import os
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

EXCEL_FILE = "sensores.xlsx"
INTERVALO  = 5   # segundos entre actualizacións

PARAMETROS = [
    ("Temperatura", "°C",  "#E05C5C"),
    ("Humidade",    "%",   "#4A90D9"),
    ("Altura",      "m",   "#6BBF59"),
    ("Presión",     "hPa", "#A070D0"),
]


def hex_a_rgba(cor_hex: str, alpha: float) -> str:
    """Converte '#RRGGBB' a 'rgba(r,g,b,alpha)'."""
    h = cor_hex.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


# ---------------------------------------------------------------------------
# Configuración da páxina
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Monitor de sensores",
    page_icon="📡",
    layout="wide",
)

st.title("📡 Monitor de sensores en tempo real")

# ---------------------------------------------------------------------------
# Carga de datos dende o ficheiro Excel
# ---------------------------------------------------------------------------

def cargar_datos():
    if not os.path.exists(EXCEL_FILE):
        return None, {}
    try:
        df_reg = pd.read_excel(EXCEL_FILE, sheet_name="Rexistro")
        follas = {}
        for nome, _, _ in PARAMETROS:
            follas[nome] = pd.read_excel(EXCEL_FILE, sheet_name=nome)
        return df_reg, follas
    except Exception:
        return None, {}


df_reg, follas = cargar_datos()

if df_reg is None or df_reg.empty:
    st.warning("Agardando datos… ¿está en execución `stream_to_excel.py`?")
    time.sleep(INTERVALO)
    st.rerun()

# ---------------------------------------------------------------------------
# Métricas superiores (última lectura)
# ---------------------------------------------------------------------------

ultima   = df_reg.iloc[-1]
anterior = df_reg.iloc[-2] if len(df_reg) > 1 else ultima

col_t, col_h, col_a, col_p, col_n = st.columns(5)
col_t.metric("🌡 Temperatura", f"{ultima['Temperatura (°C)']} °C",
             f"{ultima['Temperatura (°C)'] - anterior['Temperatura (°C)']:+.1f}")
col_h.metric("💧 Humidade",    f"{ultima['Humidade (%)']} %",
             f"{ultima['Humidade (%)'] - anterior['Humidade (%)']:+.1f}")
col_a.metric("⛰ Altura",       f"{ultima['Altura (m)']} m",
             f"{ultima['Altura (m)'] - anterior['Altura (m)']:+.0f}")
col_p.metric("🌬 Presión",      f"{ultima['Presión (hPa)']} hPa",
             f"{ultima['Presión (hPa)'] - anterior['Presión (hPa)']:+.1f}")
col_n.metric("📊 Lecturas",     len(df_reg))

st.caption(f"Última lectura: **{ultima['Timestamp']}**  —  actualizando cada {INTERVALO} s")
st.divider()

# ---------------------------------------------------------------------------
# Grid 2×2 de gráficas por parámetro
# ---------------------------------------------------------------------------

def facer_grafica(df, nome: str, unidade: str, cor: str) -> go.Figure:
    """Constrúe a gráfica de liñas dun parámetro con banda mín-máx e media."""
    col_val   = df.columns[2]   # valor do parámetro
    col_min   = df.columns[3]   # Mín acum.
    col_max   = df.columns[4]   # Máx acum.
    col_media = df.columns[5]   # Media acum.

    fig = go.Figure()

    # Banda mín-máx sombreada
    fig.add_trace(go.Scatter(
        x=pd.concat([df["#"], df["#"][::-1]], ignore_index=True),
        y=pd.concat([df[col_max], df[col_min][::-1]], ignore_index=True),
        fill="toself",
        fillcolor=hex_a_rgba(cor, 0.15),
        line=dict(color="rgba(0,0,0,0)"),
        name="Rango mín-máx",
        hoverinfo="skip",
        showlegend=True,
    ))

    # Valor real
    fig.add_trace(go.Scatter(
        x=df["#"], y=df[col_val],
        mode="lines+markers",
        name=f"Valor ({unidade})",
        line=dict(color=cor, width=2.5),
        marker=dict(size=5, color=cor),
    ))

    # Media acumulada
    fig.add_trace(go.Scatter(
        x=df["#"], y=df[col_media],
        mode="lines",
        name="Media",
        line=dict(color="rgba(200,200,200,0.8)", width=1.5, dash="dash"),
    ))

    fig.update_layout(
        title=dict(text=f"<b>{nome}</b> ({unidade})", font=dict(size=15)),
        xaxis_title="Lectura #",
        yaxis_title=f"{unidade}",
        height=340,
        margin=dict(l=50, r=20, t=50, b=40),
        legend=dict(
            orientation="h",
            yanchor="bottom", y=1.01,
            xanchor="left",   x=0,
            font=dict(size=11),
        ),
        # Fondo transparente para adaptarse ao tema de Streamlit
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#cccccc"),
        xaxis=dict(showgrid=True, gridcolor="rgba(150,150,150,0.2)", zeroline=False),
        yaxis=dict(showgrid=True, gridcolor="rgba(150,150,150,0.2)", zeroline=False),
    )

    return fig


# Distribuír as 4 gráficas nun grid de 2 filas × 2 columnas
fila1 = st.columns(2)
fila2 = st.columns(2)
contedores = [fila1[0], fila1[1], fila2[0], fila2[1]]

for (nome, unidade, cor), contedor in zip(PARAMETROS, contedores):
    df = follas.get(nome)
    if df is not None and not df.empty:
        fig = facer_grafica(df, nome, unidade, cor)
        contedor.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# Táboa das últimas lecturas
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Últimas 20 lecturas")
st.dataframe(
    df_reg.tail(20).sort_values("#", ascending=False),
    use_container_width=True,
    hide_index=True,
)

# ---------------------------------------------------------------------------
# Auto-actualización
# ---------------------------------------------------------------------------

time.sleep(INTERVALO)
st.rerun()
