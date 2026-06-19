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

    st.html("""
    <style>
    *{box-sizing:border-box;margin:0;padding:0}
    .h-block{background:var(--color-background-primary);border:0.5px solid var(--color-border-tertiary);border-radius:var(--border-radius-lg);padding:1.25rem 1.5rem;margin:0 0 1rem}
    .h-head{display:flex;align-items:center;gap:10px;margin:0 0 1rem;padding-bottom:0.75rem;border-bottom:0.5px solid var(--color-border-tertiary)}
    .h-icon{width:32px;height:32px;border-radius:8px;background:#EAF3DE;display:flex;align-items:center;justify-content:center;flex-shrink:0}
    .h-icon svg{display:block}
    .h-title{font-size:15px;font-weight:500;color:var(--color-text-primary)}
    .h-desc{font-size:13px;color:var(--color-text-secondary);line-height:1.7;margin:0 0 1.25rem}
    .h-table{width:100%;border-collapse:collapse;font-size:13px}
    .h-table th{text-align:left;font-weight:500;font-size:11px;color:var(--color-text-secondary);padding:5px 10px;border-bottom:0.5px solid var(--color-border-tertiary);text-transform:uppercase;letter-spacing:0.05em}
    .h-table td{padding:9px 10px;border-bottom:0.5px solid var(--color-border-tertiary);color:var(--color-text-primary);vertical-align:middle;font-size:13px}
    .h-table tr:last-child td{border-bottom:none}
    .h-dot{width:9px;height:9px;border-radius:50%;display:inline-block;margin-right:8px;vertical-align:middle;flex-shrink:0}
    .h-cod{font-family:monospace;font-size:11px;font-weight:500;background:#EAF3DE;color:#27500A;border-radius:4px;padding:2px 7px;letter-spacing:0.03em}
    .h-pos{font-size:12px;color:var(--color-text-secondary)}
    .h-cat-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:0.75rem 0 0}
    .h-cat-card{background:#F5F9F0;border:0.5px solid #C0DD97;border-radius:8px;padding:12px 14px;display:flex;align-items:center;gap:12px}
    .h-cat-icon{width:36px;height:36px;border-radius:7px;background:#EAF3DE;display:flex;align-items:center;justify-content:center;flex-shrink:0}
    .h-cat-label{font-size:13px;font-weight:500;color:#27500A}
    .h-cat-sub{font-size:11px;color:#3B6D11;margin-top:2px}
    .h-step-grid{display:flex;flex-direction:column;gap:12px}
    .h-step{display:flex;gap:14px;align-items:flex-start}
    .h-step-num{width:24px;height:24px;border-radius:50%;background:#EAF3DE;border:0.5px solid #C0DD97;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:500;color:#27500A;flex-shrink:0;margin-top:1px}
    .h-step-body p{font-size:13px;font-weight:500;color:var(--color-text-primary);margin:0 0 2px}
    .h-step-body span{font-size:12px;color:var(--color-text-secondary);line-height:1.6}
    .h-pref-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:0.75rem 0 0}
    .h-pref-box{background:#F5F9F0;border:0.5px solid #C0DD97;border-radius:8px;padding:12px 13px}
    .h-pref-letter{font-size:10px;font-weight:500;color:#3B6D11;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:5px}
    .h-pref-box p{font-size:13px;font-weight:500;color:var(--color-text-primary);margin:0 0 3px}
    .h-pref-box span{font-size:12px;color:var(--color-text-secondary);line-height:1.5}
    .h-viz-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:0.75rem 0 1rem}
    .h-viz-box{background:#F5F9F0;border:0.5px solid #C0DD97;border-radius:8px;padding:12px 14px;display:flex;gap:12px;align-items:flex-start}
    .h-viz-box p{font-size:13px;font-weight:500;color:var(--color-text-primary);margin:0 0 3px}
    .h-viz-box span{font-size:12px;color:var(--color-text-secondary);line-height:1.5}
    .h-strat-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:0 0 1rem}
    .h-strat{border-radius:8px;padding:10px 12px;border:0.5px solid transparent}
    .h-strat p{font-size:12px;font-weight:500;margin:0 0 2px}
    .h-strat span{font-size:11px;line-height:1.4}
    .h-note{border-left:2px solid #C0DD97;padding:8px 14px}
    .h-note span{font-size:12px;color:var(--color-text-secondary);line-height:1.6}
    .h-dl-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin:0.75rem 0 0}
    .h-dl{background:#F5F9F0;border:0.5px solid #C0DD97;border-radius:8px;padding:11px 14px;display:flex;gap:10px;align-items:center}
    .h-dl p{font-size:13px;font-weight:500;color:#27500A;margin:0 0 1px}
    .h-dl span{font-size:11px;color:#3B6D11}
    </style>

    <div style="padding:0.25rem 0 1rem">

    <div class="h-block">
      <div class="h-head">
        <div class="h-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M8 2C5.5 2 3 4 3 7c0 2 1 3.5 2.5 4.5V13h5v-1.5C12 10.5 13 9 13 7c0-3-2.5-5-5-5z" stroke="#3B6D11" stroke-width="1.1" stroke-linejoin="round"/><path d="M6 13h4" stroke="#3B6D11" stroke-width="1.1" stroke-linecap="round"/></svg></div>
        <span class="h-title">¿Qué es un HyphaPod?</span>
      </div>
      <p class="h-desc">Sistema de diseño de microbosques urbanos basado en el <strong style="font-weight:500;color:var(--color-text-primary)">método Miyawaki</strong>, adaptado a las condiciones climáticas y florísticas de <strong style="font-weight:500;color:var(--color-text-primary)">Córdoba</strong> (clima mediterráneo continental). Genera automáticamente la distribución de plantas en una parcela rectangular, respetando el gradiente de estratos concéntrico del método.</p>
      <table class="h-table">
        <thead><tr><th>Estrato</th><th>Código</th><th>Altura</th><th>Posición</th></tr></thead>
        <tbody>
          <tr><td><span class="h-dot" style="background:#D6EAD0;border:0.5px solid #aaa"></span>Caméfitas</td><td><span class="h-cod">CA</span></td><td>&lt; 1 m</td><td><span class="h-pos">Borde exterior</span></td></tr>
          <tr><td><span class="h-dot" style="background:#9FCD88"></span>Microfanerófitas</td><td><span class="h-cod">MI</span></td><td>1–4 m</td><td><span class="h-pos">Anillos intermedios</span></td></tr>
          <tr><td><span class="h-dot" style="background:#4E9940"></span>Mesofanerófitas</td><td><span class="h-cod">ME</span></td><td>3–10 m</td><td><span class="h-pos">Anillos interiores</span></td></tr>
          <tr><td><span class="h-dot" style="background:#3A6B27"></span>Megafanerófitas</td><td><span class="h-cod">MG</span></td><td>8–30 m</td><td><span class="h-pos">Núcleo central</span></td></tr>
        </tbody>
      </table>
    </div>

    <div class="h-block">
      <div class="h-head">
        <div class="h-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="2" y="2" width="12" height="12" rx="2" stroke="#3B6D11" stroke-width="1.1"/><path d="M5 8h6M5 5.5h6M5 10.5h4" stroke="#3B6D11" stroke-width="1" stroke-linecap="round"/></svg></div>
        <span class="h-title">Catálogo de especies</span>
      </div>
      <p class="h-desc">39 especies autóctonas vasculares seleccionadas para las condiciones de Córdoba, con datos agronómicos y fichas fotográficas.</p>
      <div class="h-cat-grid">
        <div class="h-cat-card">
          <div class="h-cat-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="3" y="1" width="10" height="14" rx="1.5" stroke="#3B6D11" stroke-width="1.1"/><path d="M5 5h6M5 8h6M5 11h3" stroke="#3B6D11" stroke-width="1" stroke-linecap="round"/></svg></div>
          <div><div class="h-cat-label">Catálogo con fotos</div><div class="h-cat-sub">Descargar PDF</div></div>
        </div>
        <div class="h-cat-card">
          <div class="h-cat-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="1" y="3" width="14" height="10" rx="1.5" stroke="#3B6D11" stroke-width="1.1"/><path d="M1 6h14M5 6v7" stroke="#3B6D11" stroke-width="1" stroke-linecap="round"/></svg></div>
          <div><div class="h-cat-label">Catálogo de especies</div><div class="h-cat-sub">Descargar Excel</div></div>
        </div>
      </div>
    </div>

    <div class="h-block">
      <div class="h-head">
        <div class="h-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="#3B6D11" stroke-width="1.1"/><path d="M8 5v3l2 2" stroke="#3B6D11" stroke-width="1.1" stroke-linecap="round"/></svg></div>
        <span class="h-title">Cómo usar la aplicación</span>
      </div>
      <div class="h-step-grid">
        <div class="h-step"><div class="h-step-num">1</div><div class="h-step-body"><p>Configura la parcela</p><span>Introduce las dimensiones y el presupuesto máximo en el panel izquierdo. La superficie debe estar entre 100 y 300 m².</span></div></div>
        <div class="h-step"><div class="h-step-num">2</div><div class="h-step-body"><p>Personaliza el diseño <em style="font-weight:400;color:var(--color-text-secondary)">(opcional)</em></p><span>Activa "Personalizar el diseño" para ajustar estratos, excluir especies o indicar preferencias de plantación por zona.</span></div></div>
        <div class="h-step"><div class="h-step-num">3</div><div class="h-step-body"><p>Pulsa "Generar diseño"</p><span>El generador calcula la distribución de plantas, el coste estimado y genera todos los archivos de salida.</span></div></div>
        <div class="h-step"><div class="h-step-num">4</div><div class="h-step-body"><p>Consulta y descarga</p><span>Los resultados aparecen en tres pestañas: resumen económico, plano visual y descargas.</span></div></div>
      </div>
    </div>

    <div class="h-block">
      <div class="h-head">
        <div class="h-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M2 13l3-4 3 2 3-5 3 7" stroke="#3B6D11" stroke-width="1.1" stroke-linecap="round" stroke-linejoin="round"/></svg></div>
        <span class="h-title">Cómo genera el diseño el generador</span>
      </div>
      <p class="h-desc">El generador no coloca las plantas de forma fija: cada resultado es distinto. El sistema sigue dos reglas simultáneas:</p>
      <div class="h-step-grid" style="margin:0 0 0">
        <div class="h-step"><div class="h-step-num">1</div><div class="h-step-body"><p>Gradiente de alturas</p><span>La posición de cada celda determina qué estrato tiene más probabilidades de aparecer. Borde → plantas bajas (CA). Centro → árboles altos (MG). Las zonas intermedias mezclan estratos en proporciones variables, generando transiciones naturales.</span></div></div>
        <div class="h-step"><div class="h-step-num">2</div><div class="h-step-body"><p>Diversidad obligatoria</p><span>Dentro de cada celda de 1 m² nunca se repite la misma especie. La competencia interespecífica entre plantas de distinto porte es uno de los principios clave del método Miyawaki.</span></div></div>
      </div>
    </div>

    <div class="h-block">
      <div class="h-head">
        <div class="h-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="2" y="2" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/><rect x="9" y="2" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/><rect x="2" y="9" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/><rect x="9" y="9" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/></svg></div>
        <span class="h-title">Preferencias de diseño</span>
      </div>
      <p class="h-desc">Permiten modificar las probabilidades por defecto para que una especie tenga más o menos presencia en una zona concreta. Se añaden una a una con el botón ➕.</p>
      <div class="h-pref-grid">
        <div class="h-pref-box"><div class="h-pref-letter">A — Especie</div><span>La especie a la que quieres dar protagonismo en una zona concreta de la parcela.</span></div>
        <div class="h-pref-box"><div class="h-pref-letter">B — Anillo</div><span>La zona de la parcela. Anillo 1 = borde. Último anillo = centro.</span></div>
        <div class="h-pref-box"><div class="h-pref-letter">C — Probabilidad (%)</div><span>1–25: puede aparecer pero no domina. 75–100: muchas papeletas para ser seleccionada.</span></div>
      </div>
      <div class="h-note" style="margin-top:1rem"><span>Si una preferencia no es ecológicamente coherente, el generador muestra un aviso pero respeta la decisión sin bloquear la ejecución.</span></div>
    </div>

    <div class="h-block">
      <div class="h-head">
        <div class="h-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="2" y="2" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/><rect x="9" y="2" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/><rect x="2" y="9" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/><rect x="9" y="9" width="5" height="5" rx="1" stroke="#3B6D11" stroke-width="1.1"/></svg></div>
        <span class="h-title">Cómo leer el plano</span>
      </div>
      <div class="h-viz-grid">
        <div class="h-viz-box">
          <svg width="40" height="40" viewBox="0 0 40 40"><rect x="2" y="2" width="36" height="36" rx="4" fill="#9FCD88"/><circle cx="20" cy="20" r="10" fill="#D6EAD0" stroke="#1A3A1A" stroke-width="0.8"/><text x="20" y="24" text-anchor="middle" font-size="7" font-weight="bold" fill="#111">CA01</text></svg>
          <div><p>Cuadrado de un color</p><span>Todas las plantas de esa celda (1 m²) pertenecen al mismo estrato.</span></div>
        </div>
        <div class="h-viz-box">
          <svg width="40" height="40" viewBox="0 0 40 40"><rect x="2" y="2" width="36" height="36" rx="4" fill="#BBDCA8"/><circle cx="14" cy="14" r="8" fill="#D6EAD0" stroke="#1A3A1A" stroke-width="0.8"/><text x="14" y="18" text-anchor="middle" font-size="6" font-weight="bold" fill="#111">CA03</text><circle cx="28" cy="28" r="8" fill="#4E9940" stroke="#1A3A1A" stroke-width="0.8"/><text x="28" y="32" text-anchor="middle" font-size="6" font-weight="bold" fill="white">ME02</text></svg>
          <div><p>Cuadrado con círculos</p><span>Celda mixta con 2–3 plantas de estratos distintos. Cada círculo = una planta.</span></div>
        </div>
      </div>
      <div class="h-strat-grid">
        <div class="h-strat" style="background:#F0F8EC;border-color:#C8DDB8"><p style="color:#27500A">CA — Caméfitas</p><span style="color:#3B6D11">Borde exterior · &lt; 1 m</span></div>
        <div class="h-strat" style="background:#E4F2D8;border-color:#9FCD88"><p style="color:#27500A">MI — Microfaner.</p><span style="color:#3B6D11">Intermedios · 1–4 m</span></div>
        <div class="h-strat" style="background:#C8DFB0;border-color:#4E9940"><p style="color:#1A4D0A">ME — Mesofaner.</p><span style="color:#27500A">Interiores · 3–10 m</span></div>
        <div class="h-strat" style="background:#3A6B27;border-color:#27500A"><p style="color:#EAF3DE">MG — Megafaner.</p><span style="color:#C0DD97">Núcleo · 8–30 m</span></div>
      </div>
      <div class="h-note"><span>La vista previa es orientativa. Para ver los códigos de especie con detalle, usa los archivos de descarga.</span></div>
    </div>

    <div class="h-block">
      <div class="h-head">
        <div class="h-icon"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M8 2v8M5 7l3 3 3-3" stroke="#3B6D11" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/><path d="M3 13h10" stroke="#3B6D11" stroke-width="1.1" stroke-linecap="round"/></svg></div>
        <span class="h-title">Archivos disponibles en Descargas</span>
      </div>
      <div class="h-dl-grid">
        <div class="h-dl"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="1" y="3" width="14" height="10" rx="1.5" stroke="#3B6D11" stroke-width="1.1"/><path d="M1 6h14M5 6v7" stroke="#3B6D11" stroke-width="1" stroke-linecap="round"/></svg><div><p>Excel</p><span>Cuadrícula del diseño y desglose de costes</span></div></div>
        <div class="h-dl"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="3" y="1" width="10" height="14" rx="1.5" stroke="#3B6D11" stroke-width="1.1"/><path d="M5 5h6M5 8h6M5 11h3" stroke="#3B6D11" stroke-width="1" stroke-linecap="round"/></svg><div><p>PDF — DIN A1</p><span>Plano técnico con leyenda y cajetín</span></div></div>
        <div class="h-dl"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M2 14l4-4m0 0l2-6 2 6m-4 0h4m2 4l-4-4" stroke="#3B6D11" stroke-width="1.1" stroke-linecap="round" stroke-linejoin="round"/></svg><div><p>DXF — plano CAD</p><span>Compatible con AutoCAD, QGIS y LibreCAD</span></div></div>
        <div class="h-dl"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M8 2l4 2.5v5L8 12 4 9.5v-5L8 2z" stroke="#3B6D11" stroke-width="1.1" stroke-linejoin="round"/><path d="M8 2v10M4 4.5l4 3 4-3" stroke="#3B6D11" stroke-width="1" stroke-linecap="round"/></svg><div><p>GLB — modelo 3D</p><span>Visualizable en 3dviewer.net</span></div></div>
      </div>
    </div>

    </div>
    """)

    _CAT_DIR = os.path.join(_APP_DIR, "catálogos")
    _pdf_path = os.path.join(_CAT_DIR, "cátalogo_mycoroots.pdf")
    _xls_path = os.path.join(_CAT_DIR, "catálogo_especies_mycoroots.xlsx")
    cat_col1, cat_col2 = st.columns(2)
    if os.path.exists(_pdf_path):
        with open(_pdf_path, "rb") as f:
            cat_col1.download_button("Descargar catálogo PDF", f.read(), "catálogo_mycoroots.pdf", "application/pdf", use_container_width=True)
    if os.path.exists(_xls_path):
        with open(_xls_path, "rb") as f:
            cat_col2.download_button("Descargar catálogo Excel", f.read(), "catálogo_especies_mycoroots.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

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
