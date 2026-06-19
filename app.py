"""
MycoRoots — HyphaPod Generator
Interfaz web con Streamlit
Autora: Ana Velasco Márquez · TFG Ingeniería Agroalimentaria · UCO
"""
import sys, os, io, tempfile

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# ── Rutas ──────────────────────────────────────────────────────────────
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_SCR_DIR = os.path.join(_APP_DIR, "scripts")
if _SCR_DIR not in sys.path:
    sys.path.insert(0, _SCR_DIR)

import hyphapod_generator_v4 as gen
import hyphapod_plano_v5     as plano
import hyphapod_glb          as glb_mod

# ── Configuración de página ────────────────────────────────────────────
st.set_page_config(
    page_title="MycoRoots — HyphaPod",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Catálogos (cacheados al arrancar) ──────────────────────────────────
@st.cache_resource
def _cat_gen():
    return gen.cargar_catalogo()

@st.cache_resource
def _cat_plano():
    return plano.cargar_catalogo(plano.EXCEL_CATALOGO)

@st.cache_resource
def _cat_glb():
    return glb_mod.cargar_catalogo(glb_mod.EXCEL_CATALOGO)

try:
    CAT_GEN   = _cat_gen()
    CAT_PLANO = _cat_plano()
    CAT_GLB   = _cat_glb()
except Exception as e:
    st.error(f"No se pudo cargar el catálogo de especies: {e}")
    st.stop()

# ── Paletas de color ───────────────────────────────────────────────────
COLOR_EST = {"CA": "#D6EAD0", "MI": "#9FCD88", "ME": "#4E9940", "MG": "#3A6B27"}
COLOR_EST_DARK = {"CA": False, "MI": False, "ME": True, "MG": True}
COLOR_COMB = {
    frozenset(["CA"]):        "#D6EAD0",
    frozenset(["CA","MI"]):   "#BBDCA8",
    frozenset(["MI"]):        "#9FCD88",
    frozenset(["CA","ME"]):   "#83BE70",
    frozenset(["MI","ME"]):   "#68AF58",
    frozenset(["ME"]):        "#4E9940",
    frozenset(["MG"]):        "#3A6B27",
}
NOMBRES_EST = {
    "CA": "Caméfitas (<1 m)",
    "MI": "Microfaner. (1–4 m)",
    "ME": "Mesofaner. (3–10 m)",
    "MG": "Megafaner. (8–30 m)",
}

# ── Funciones de conversión y generación ──────────────────────────────
def grid_to_celdas(grid):
    """Convierte la lista de listas del generador a dict {(col, fila): [cods]}."""
    celdas = {}
    for fi, fila in enumerate(grid):
        for ci, celda in enumerate(fila):
            cods = [e["cod"] for e in celda["especies"]]
            if cods:
                celdas[(ci, fi)] = cods
    return celdas

def generar_excel_bytes(grid, params, coste, total_plantas, desglose, n_anillos):
    buf = io.BytesIO()
    gen.exportar_excel(grid, params, coste, total_plantas, desglose, buf, n_anillos)
    buf.seek(0)
    return buf.getvalue()

def generar_pdf_bytes(celdas, n_cols, n_filas):
    buf = io.BytesIO()
    plano.generar_pdf(celdas, n_cols, n_filas, CAT_PLANO, buf)
    buf.seek(0)
    return buf.getvalue()

def generar_dxf_bytes(celdas, n_cols, n_filas):
    with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as f:
        tmp = f.name
    try:
        plano.generar_dxf(celdas, n_cols, n_filas, CAT_PLANO, tmp)
        with open(tmp, "rb") as f:
            return f.read()
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass

def generar_glb_bytes(celdas, n_cols, n_filas):
    return glb_mod.generar_glb_bytes(celdas, n_cols, n_filas, CAT_GLB)

# ── Visualización de la cuadrícula ────────────────────────────────────
def dibujar_cuadricula(celdas, n_cols, n_filas):
    fw = min(15, max(8, n_cols * 0.70 + 3))
    fh = min(11, max(6, n_filas * 0.70 + 2))
    fig, ax = plt.subplots(figsize=(fw, fh), facecolor="#F5F0E8")
    ax.set_facecolor("#DDD0A8")
    ax.set_xlim(0, n_cols)
    ax.set_ylim(0, n_filas)
    ax.set_aspect("equal")
    ax.set_title("Plano de plantación — HyphaPod", fontsize=10, fontweight="bold", pad=8)
    ax.set_xlabel("Largo (m)", fontsize=8)
    ax.set_ylabel("Ancho (m)", fontsize=8)
    ax.tick_params(labelsize=7)
    ax.set_xticks(range(n_cols + 1))
    ax.set_yticks(range(n_filas + 1))

    for (col, fila), cods in celdas.items():
        y = n_filas - fila - 1
        ests = frozenset(c[:2] for c in cods)
        bg = COLOR_COMB.get(ests, "#D9D9D9")
        ax.add_patch(plt.Rectangle((col, y), 1, 1, color=bg, linewidth=0, zorder=1))

        n = len(cods)
        if n == 1:
            pos = [(col + .5, y + .5)]; r = 0.39
        elif n == 2:
            pos = [(col + .27, y + .73), (col + .73, y + .27)]; r = 0.19
        else:
            cx, cy = col + .5, y + .5
            pos = [(cx - .20, cy + .19), (cx + .20, cy + .19), (cx, cy - .25)]; r = 0.18

        for k, (px, py) in enumerate(pos[:len(cods)]):
            est = cods[k][:2]
            fc  = COLOR_EST.get(est, "#888888")
            circle = plt.Circle((px, py), r, color=fc, linewidth=0.4,
                                 edgecolor="#1A3A1A", zorder=2)
            ax.add_patch(circle)
            fs = max(3.0, r * 7.5)
            tc = "white" if COLOR_EST_DARK.get(est) else "#111111"
            ax.text(px, py, cods[k], ha="center", va="center",
                    fontsize=fs, color=tc, fontweight="bold", zorder=3)

    for i in range(n_cols + 1):
        ax.axvline(i, color="#7A6A4A", linewidth=0.25, alpha=0.5, zorder=0)
    for j in range(n_filas + 1):
        ax.axhline(j, color="#7A6A4A", linewidth=0.25, alpha=0.5, zorder=0)

    handles = [
        mpatches.Patch(facecolor=COLOR_EST[e], edgecolor="#1A3A1A", label=NOMBRES_EST[e])
        for e in ["CA", "MI", "ME", "MG"]
    ]
    ax.legend(handles=handles, loc="upper left", bbox_to_anchor=(1.01, 1.0),
              fontsize=7, title="Estratos", title_fontsize=7.5, framealpha=0.9)
    plt.tight_layout()
    return fig

# ── Session state ──────────────────────────────────────────────────────
if "resultado" not in st.session_state:
    st.session_state.resultado = None
if "prefs" not in st.session_state:
    st.session_state.prefs = []

# ══════════════════════════════════════════════════════════════════════
# SIDEBAR — Configuración
# ══════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🌿 MycoRoots")
    st.caption("HyphaPod Generator · Microbosque urbano · Miyawaki")
    st.divider()

    # ── 1. Dimensiones ────────────────────────────────────────────────
    st.markdown("### 📐 Parcela")
    largo = st.slider("Largo (m)", min_value=10, max_value=20, value=12)
    ancho = st.slider("Ancho (m)", min_value=10, max_value=20, value=10)
    sup   = largo * ancho
    if 100 <= sup <= 300:
        st.success(f"{largo} × {ancho} m = **{sup} m²**")
    else:
        st.error(f"{sup} m² — debe estar entre 100 y 300 m²")

    # ── 2. Presupuesto ────────────────────────────────────────────────
    st.markdown("### 💶 Presupuesto")
    presupuesto = st.number_input(
        "Máximo (€)", min_value=100.0, value=2000.0, step=100.0, format="%.0f"
    )

    # ── 3. Personalización ────────────────────────────────────────────
    st.markdown("### 🌱 Personalización")
    personalizar = st.checkbox("Personalizar el diseño")

    estratos_sel = ["CA", "MI", "ME", "MG"]
    excluidas    = []
    prefs_ok     = []

    if personalizar:
        # Estratos
        with st.expander("Estratos a incluir"):
            all_ests = {
                "CA — Caméfitas (<1 m)":      "CA",
                "MI — Microfaner. (1–4 m)":   "MI",
                "ME — Mesofaner. (3–10 m)":   "ME",
                "MG — Megafaner. (>8 m)":     "MG",
            }
            sel_est = st.multiselect(
                "Estratos", list(all_ests.keys()),
                default=list(all_ests.keys()), label_visibility="collapsed"
            )
            estratos_sel = [all_ests[k] for k in sel_est] if sel_est else ["CA","MI","ME","MG"]

        # Exclusiones
        with st.expander("Excluir especies"):
            opt_exc = {f"{e['cod']} — {e['nombre_comun']}": e["cod"] for e in CAT_GEN}
            sel_exc = st.multiselect("Especies", list(opt_exc.keys()), label_visibility="collapsed")
            excluidas = [opt_exc[k] for k in sel_exc]

            st.markdown("**Por categoría de precio:**")
            c1, c2 = st.columns(2)
            exc_pre = c1.checkbox("PRE (caras)")
            exc_med = c2.checkbox("MED (medias)")
            if exc_pre or exc_med:
                cats = (["PRE"] if exc_pre else []) + (["MED"] if exc_med else [])
                extras = [e["cod"] for e in CAT_GEN
                          if e["cat_precio"] in cats and e["cod"] not in excluidas]
                excluidas = list(set(excluidas + extras))

        # Preferencias de diseño
        n_an_prev = (min(ancho, largo) - 1) // 2 + 1
        with st.expander("Preferencias de diseño"):
            opt_pref = {
                f"{e['cod']} — {e['nombre_comun']} [{e['cod_estrato']}]": e["cod"]
                for e in CAT_GEN
            }
            esp_p  = st.selectbox("Especie", ["— elige —"] + list(opt_pref.keys()), key="ps_esp")
            zona_p = st.slider(f"Anillo (1 = borde, {n_an_prev} = centro)",
                               1, n_an_prev, 1, key="ps_zona")
            prob_p = st.slider("Probabilidad (%)", 1, 100, 50, key="ps_prob")

            if st.button("➕ Añadir preferencia", key="ps_add") and esp_p != "— elige —":
                cod_p = opt_pref[esp_p]
                if not any(p["cod"] == cod_p and p["zona"] == zona_p
                           for p in st.session_state.prefs):
                    st.session_state.prefs.append(
                        {"cod": cod_p, "zona": zona_p, "prob": prob_p}
                    )

            if st.session_state.prefs:
                st.markdown("**Activas:**")
                for i, p in enumerate(list(st.session_state.prefs)):
                    ca, cb = st.columns([5, 1])
                    ca.caption(f"`{p['cod']}` · Anillo {p['zona']} · {p['prob']}%")
                    if cb.button("✕", key=f"del_p_{i}"):
                        st.session_state.prefs.pop(i)
                        st.rerun()
                if st.button("🗑 Limpiar", key="clear_prefs"):
                    st.session_state.prefs = []
                    st.rerun()

        # Convertir prefs al formato del generador
        for p in st.session_state.prefs:
            esp = next((e for e in CAT_GEN if e["cod"] == p["cod"]), None)
            if esp:
                prefs_ok.append({"especie": esp, "zona": p["zona"], "prob": p["prob"]})

    st.divider()
    btn_gen = st.button(
        "🌿 Generar diseño", type="primary",
        use_container_width=True,
        disabled=not (100 <= sup <= 300),
    )

# ══════════════════════════════════════════════════════════════════════
# LÓGICA DE GENERACIÓN
# ══════════════════════════════════════════════════════════════════════
if btn_gen:
    params = {
        "largo":       largo,
        "ancho":       ancho,
        "superficie":  sup,
        "presupuesto": presupuesto,
        "estratos":    estratos_sel,
        "excluidas":   list(set(excluidas)),
        "preferencias": prefs_ok,
    }

    prog = st.progress(0, text="Calculando distribución de plantas…")
    try:
        grid, err = gen.generar_grid(params, CAT_GEN)
        if err:
            st.error(f"Error al generar el diseño: {err}")
            prog.empty()
            st.stop()

        coste, total_plantas, desglose = gen.calcular_coste(grid)
        n_an   = (min(ancho, largo) - 1) // 2 + 1
        celdas = grid_to_celdas(grid)
        prog.progress(15, "Generando Excel…")

        xls = generar_excel_bytes(grid, params, coste, total_plantas, desglose, n_an)
        prog.progress(35, "Generando plano PDF…")

        pdf = generar_pdf_bytes(celdas, largo, ancho)
        prog.progress(60, "Generando plano DXF…")

        dxf = generar_dxf_bytes(celdas, largo, ancho)
        prog.progress(80, "Generando modelo 3D…")

        glb = generar_glb_bytes(celdas, largo, ancho)
        prog.progress(100, "✅ Listo")

    except Exception as e:
        st.error(f"Error inesperado: {e}")
        prog.empty()
        st.stop()

    prog.empty()

    st.session_state.resultado = {
        "grid":          grid,
        "params":        params,
        "coste":         coste,
        "total_plantas": total_plantas,
        "desglose":      desglose,
        "n_anillos":     n_an,
        "celdas":        celdas,
        "n_cols":        largo,
        "n_filas":       ancho,
        "xls":           xls,
        "pdf":           pdf,
        "dxf":           dxf,
        "glb":           glb,
    }

# ══════════════════════════════════════════════════════════════════════
# ÁREA PRINCIPAL
# ══════════════════════════════════════════════════════════════════════
st.markdown("# 🌿 MycoRoots — HyphaPod Generator")
st.caption(
    "Microbosque urbano · Método Miyawaki · "
    "TFG Ingeniería Agroalimentaria — Universidad de Córdoba · "
    "Ana Velasco Márquez"
)
st.divider()

# ── Pantalla de bienvenida ────────────────────────────────────────────
if st.session_state.resultado is None:
    c1, c2, c3 = st.columns(3)
    c1.info("**1. Configura** la parcela y el presupuesto en el panel izquierdo")
    c2.info("**2. Personaliza** estratos, exclusiones y preferencias (opcional)")
    c3.info("**3. Pulsa Generar** y descarga el plano, Excel y modelo 3D")

    st.markdown("""
    ### ¿Qué es un HyphaPod?
    Sistema de diseño de microbosques urbanos basado en el **Método Miyawaki**,
    adaptado a las condiciones climáticas y florísticas de **Córdoba** (clima mediterráneo continental).
    Genera automáticamente la distribución de plantas en una parcela rectangular,
    respetando el gradiente de estratos concéntrico del método.

    | Estrato | Código | Altura aprox. | Posición en el diseño |
    |---------|--------|---------------|----------------------|
    | Caméfitas | **CA** | < 1 m | Anillo exterior (borde) |
    | Microfanerófitas | **MI** | 1–4 m | Anillos intermedios |
    | Mesofanerófitas | **ME** | 3–10 m | Anillos interiores |
    | Megafanerófitas | **MG** | 8–30 m | Núcleo central |
    """)

    # ── Catálogo de especies ──────────────────────────────────────────
    st.markdown("### 📖 Catálogo de especies")
    st.caption("Consulta las especies disponibles con fichas fotográficas y datos agronómicos.")

    _CAT_DIR = os.path.join(_APP_DIR, "catálogos")
    _pdf_path = os.path.join(_CAT_DIR, "cátalogo_mycoroots.pdf")
    _xls_path = os.path.join(_CAT_DIR, "catálogo_especies_mycoroots.xlsx")

    cat_col1, cat_col2 = st.columns(2)
    if os.path.exists(_pdf_path):
        with open(_pdf_path, "rb") as f:
            cat_col1.download_button(
                label="⬇️ Catálogo con fotos (PDF)",
                data=f.read(),
                file_name="catálogo_mycoroots.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
    if os.path.exists(_xls_path):
        with open(_xls_path, "rb") as f:
            cat_col2.download_button(
                label="⬇️ Catálogo de especies (Excel)",
                data=f.read(),
                file_name="catálogo_especies_mycoroots.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ── Resultados ────────────────────────────────────────────────────────
else:
    r = st.session_state.resultado

    # Aviso de sobrecoste
    if r["coste"] > r["params"]["presupuesto"]:
        diff = r["coste"] - r["params"]["presupuesto"]
        st.warning(
            f"⚠️ El diseño cuesta **{r['coste']:.2f} €** y supera el presupuesto "
            f"en **{diff:.2f} €**. "
            "Puedes reducir las dimensiones o excluir especies PRE/MED y volver a generar."
        )

    tab_res, tab_grid, tab_dl = st.tabs(["📊 Resumen", "🗺️ Cuadrícula", "📥 Descargas"])

    # ── TAB: RESUMEN ──────────────────────────────────────────────────
    with tab_res:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Superficie",      f"{r['params']['superficie']} m²")
        m2.metric("Total plantas",   r["total_plantas"])
        m3.metric("Coste estimado",  f"{r['coste']:.2f} €")
        m4.metric("Anillos",         r["n_anillos"])

        st.markdown("#### Distribución por estrato")
        est_counts = {}
        for cod, d in r["desglose"].items():
            est_counts[cod[:2]] = est_counts.get(cod[:2], 0) + d["n_plantas"]
        cols_est = st.columns(4)
        for i, est in enumerate(["CA", "MI", "ME", "MG"]):
            n   = est_counts.get(est, 0)
            pct = n / r["total_plantas"] * 100 if r["total_plantas"] else 0
            cols_est[i].metric(NOMBRES_EST[est], f"{n} ud.", f"{pct:.1f} %")

        st.markdown("#### Desglose por especie")
        rows = [
            {
                "Cód.":       cod,
                "Especie":    d["nombre"],
                "Est.":       cod[:2],
                "N.º":        d["n_plantas"],
                "€/ud":       round(d["precio_ud"], 2),
                "Subtotal €": round(d["subtotal"], 2),
            }
            for cod, d in sorted(r["desglose"].items())
        ]
        df = pd.DataFrame(rows)

        def _col_est(val):
            bg = COLOR_EST.get(val, "")
            fg = "white" if COLOR_EST_DARK.get(val) else "#111111"
            return f"background-color: {bg}; color: {fg}" if bg else ""

        st.dataframe(
            df.style.map(_col_est, subset=["Est."]),
            use_container_width=True,
            hide_index=True,
        )

    # ── TAB: CUADRÍCULA ───────────────────────────────────────────────
    with tab_grid:
        st.caption(
            "Vista previa del plano. Cada celda representa 1 × 1 m². "
            "Los círculos indican las plantas con su código de especie."
        )
        fig = dibujar_cuadricula(r["celdas"], r["n_cols"], r["n_filas"])
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    # ── TAB: DESCARGAS ────────────────────────────────────────────────
    with tab_dl:
        st.markdown("### 📥 Archivos generados")
        fname = f"HyphaPod_{r['params']['largo']}x{r['params']['ancho']}"

        col_a, col_b = st.columns(2)

        with col_a:
            st.download_button(
                label="⬇️ Excel — Cuadrícula + Resumen económico",
                data=r["xls"],
                file_name=f"{fname}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption("Cuadrícula del diseño por celdas y desglose de costes por especie.")

            st.markdown("")
            st.download_button(
                label="⬇️ PDF — Plano de plantación (DIN A1)",
                data=r["pdf"],
                file_name=f"{fname}_Plano.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            st.caption("Plano técnico con leyenda, escala gráfica y cajetín.")

        with col_b:
            st.download_button(
                label="⬇️ DXF — Plano CAD",
                data=r["dxf"],
                file_name=f"{fname}_Plano.dxf",
                mime="application/octet-stream",
                use_container_width=True,
            )
            st.caption("Plano vectorial compatible con AutoCAD, QGIS y LibreCAD.")

            st.markdown("")
            st.download_button(
                label="⬇️ GLB — Modelo 3D volumétrico",
                data=r["glb"],
                file_name=f"{fname}_3D.glb",
                mime="model/gltf-binary",
                use_container_width=True,
            )
            st.caption(
                "Modelo 3D con troncos y copas. "
                "Ábrelo arrastrándolo a [3dviewer.net](https://3dviewer.net)."
            )

# ── Pie de página ─────────────────────────────────────────────────────
st.divider()
st.caption(
    "MycoRoots · TFG Ingeniería Agroalimentaria · Universidad de Córdoba · "
    "Autora: Ana Velasco Márquez"
)
