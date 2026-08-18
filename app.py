"""
Qualitech — Dashboard de Performance Operacional (app Streamlit)

Lê o Performance_<Mês>.xlsx diretamente da pasta sincronizada do
OneDrive/SharePoint informada na barra lateral e recalcula tudo sozinho
sempre que o arquivo é atualizado — não depende de voltar a pedir para o
Claude gerar um HTML novo.

Layout e padrão visual espelham o dashboard HTML de referência
(Qualitech_Dashboard_Performance_Ago2026.html): mesma paleta de marca,
mesmo cabeçalho ("brandbar"), mesmos cartões de KPI, painéis com título +
descrição. Três abas — "Performance Operacional" (diárias/WO), "Recursos
& Pessoas" (efetivo/RH) e "Escalas & Embarques" (dobras/AFI/ociosidade) —
para não exigir rolar a página inteira para alternar entre os assuntos.

Rodar localmente:
    pip install -r requirements.txt
    streamlit run app.py
"""
import base64
import html
import itertools
import os
import time

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from streamlit_autorefresh import st_autorefresh

from data import DashboardData, list_available_workbooks, load_workbook_data
from workforce import (
    WorkforceData,
    list_available_workforce_workbooks,
    load_workforce_data,
)
from embarques import (
    EmbarquesData,
    list_available_embarques_workbooks,
    load_embarques_data,
)

# --------------------------------------------------------------------------- #
# Marca — paleta oficial Qualitech IRM (igual ao dashboard HTML de referência)
# --------------------------------------------------------------------------- #

NAVY = "#212F56"      # Cloud Burst — títulos / cabeçalho
NAVY_2 = "#2C3E6E"
STEEL = "#4C5C6E"
BLUE = "#097FBB"       # Lochmara — azul primário / destaques
GREEN = "#2E7D32"      # Fixa / bom
AMBER = "#F9A825"      # Variável / atenção
RED = "#C62828"        # Spot / crítico
GREY = "#898781"       # texto apagado
BASE_GREY = "#C3C2B7"  # baseline / eixos
HAIR = "#E1E0D9"       # linhas finas / grid
TEXT_PRIMARY = "#0B0B0B"
TEXT_SECONDARY = "#52514E"
PAGE_PLANE = "#F1F3F6"
SURFACE_2 = "#FBFBFA"

SERIES = [BLUE, "#eb6834", "#1baf7a", "#eda100", "#e87ba4", GREEN, "#4a3aa7", RED]

DEFAULT_PATH = (
    r"C:\Users\Jorge Gonçalves\OneDrive - Qualitech Inspeção, Reparo e "
    r"Manutenção Ltda\Área de Trabalho\WIP\9. Performance\Ago-2026"
)

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qualitech_logo_white.png")


def _logo_b64() -> str | None:
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


LOGO_B64 = _logo_b64()

st.set_page_config(
    page_title="Qualitech — Performance Operacional",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------------------------------------------------------- #
# CSS — porta as classes do dashboard HTML de referência (brandbar, KPI
# cards, section-title, panels, tabela, footer) para dentro do Streamlit.
# --------------------------------------------------------------------------- #

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {{ font-family: 'Inter', system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; }}
    * {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }}

    .block-container {{ padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1400px; }}
    #MainMenu, footer, [data-testid="stDecoration"] {{ visibility: hidden; }}
    /* nunca esconder o botão de reabrir a barra lateral — ele mora dentro do stToolbar */
    [data-testid="stExpandSidebarButton"] {{ visibility: visible !important; opacity: 1 !important; }}
    html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
        background:
            radial-gradient(1300px 680px at 4% -12%, rgba(9,127,187,0.30), transparent 62%),
            radial-gradient(1100px 620px at 100% -6%, rgba(33,47,86,0.26), transparent 58%),
            radial-gradient(1400px 800px at 50% 118%, rgba(33,47,86,0.14), transparent 62%),
            linear-gradient(165deg, #DEE6F3 0%, #CBD8EC 45%, #D6E0F0 100%) !important;
        background-attachment: fixed;
    }}
    [data-testid="stHeader"] {{ background: transparent !important; }}
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #FFFFFF 0%, #FAFBFD 100%);
        border-right: 1px solid {HAIR};
    }}
    [data-testid="stSidebar"] .block-container {{ padding-top: 1.1rem; }}

    /* ---------- custom scrollbar — discreto, na paleta da marca ---------- */
    ::-webkit-scrollbar {{ width: 9px; height: 9px; }}
    ::-webkit-scrollbar-track {{ background: transparent; }}
    ::-webkit-scrollbar-thumb {{ background: rgba(33,47,86,0.22); border-radius: 999px; border: 2px solid transparent; background-clip: content-box; }}
    ::-webkit-scrollbar-thumb:hover {{ background: rgba(33,47,86,0.38); background-clip: content-box; }}

    /* ---------- geometria suave — cantos mais arredondados em toda a UI ---------- */
    .stButton > button, [data-testid="stBaseButton-secondary"] {{
        border-radius: 12px !important; font-weight: 600 !important; letter-spacing: 0.1px;
        border: 1px solid {HAIR} !important; transition: all 0.16s ease !important;
        box-shadow: 0 1px 2px rgba(15,23,42,0.03);
    }}
    .stButton > button:hover {{
        border-color: {BLUE} !important; color: {BLUE} !important;
        box-shadow: 0 4px 12px rgba(9,127,187,0.16); transform: translateY(-1px);
    }}
    .stButton > button:active {{ transform: translateY(0); }}
    [data-testid="stTextInput"] input, [data-baseweb="select"] > div,
    [data-testid="stFileUploaderDropzone"], [data-baseweb="base-input"] {{
        border-radius: 12px !important;
    }}
    [data-testid="stTextInput"] input:focus {{ border-color: {BLUE} !important; box-shadow: 0 0 0 3px rgba(9,127,187,0.12) !important; }}
    [data-testid="stMultiSelect"] [data-baseweb="tag"] {{ border-radius: 999px !important; background: {NAVY} !important; }}
    [data-testid="stExpander"] {{ border-radius: 14px !important; overflow: hidden; }}
    [data-testid="stToggle"] label div[data-checked="true"] {{ background: {BLUE} !important; }}
    hr, [data-testid="stSidebarUserContent"] hr {{ border-color: {HAIR} !important; margin: 14px 0 !important; }}

    /* ---------- tabs (Performance × Recursos) — segmented control moderno ---------- */
    [data-testid="stTabs"] [role="tablist"] {{
        gap: 4px; border-bottom: none; margin-bottom: 22px;
        background: rgba(33,47,86,0.05); padding: 5px; border-radius: 14px;
        display: inline-flex; width: auto;
    }}
    [data-testid="stTab"] {{
        font-size: 13.5px; font-weight: 600; color: {STEEL}; padding: 8px 18px !important; letter-spacing: 0.1px;
        border-bottom: none !important; border-radius: 10px !important; transition: all 0.18s ease;
    }}
    [data-testid="stTab"]:hover {{ color: {NAVY}; background: rgba(255,255,255,0.6); }}
    [data-testid="stTab"][aria-selected="true"] {{
        color: {NAVY} !important; background: #fff !important;
        box-shadow: 0 2px 8px rgba(33,47,86,0.14), 0 0 0 1px rgba(33,47,86,0.05);
    }}
    [data-testid="stTab"] p {{ font-size: 13.5px !important; font-weight: 600 !important; }}
    [data-testid="stTab"][aria-selected="true"] p {{ font-weight: 700 !important; }}

    /* ---------- brandbar (cabeçalho) ---------- */
    .qt-brandbar {{
        background:
            radial-gradient(560px 220px at 96% -30%, rgba(9,127,187,0.55), transparent 65%),
            linear-gradient(135deg, {NAVY} 0%, {NAVY_2} 55%, {BLUE} 130%);
        color: #fff; padding: 22px 28px 24px; border-radius: 22px; margin-bottom: 26px;
        display: flex; align-items: flex-end; justify-content: space-between; flex-wrap: wrap; gap: 16px;
        box-shadow: 0 12px 32px rgba(33,47,86,0.24), 0 1px 0 rgba(255,255,255,0.06) inset;
        position: relative; overflow: hidden;
    }}
    .qt-brand-mark {{ display: flex; align-items: center; gap: 12px; }}
    .qt-brand-mark img {{ height: 34px; width: auto; flex: none; }}
    .qt-brand-word {{ line-height: 1.15; border-left: 1px solid rgba(255,255,255,0.28); padding-left: 12px; }}
    .qt-brand-word span {{ font-size: 11px; letter-spacing: 1.6px; text-transform: uppercase; color: #c9d7e6; }}
    .qt-brandbar h1 {{ font-size: 22px; margin: 2px 0 4px; font-weight: 700; letter-spacing: -0.2px; color: #fff; }}
    .qt-brandbar .sub {{ font-size: 13px; color: #c9d7e6; }}
    .qt-brandbar-right {{ text-align: right; }}
    .qt-period-chip {{
        display: inline-flex; align-items: center; gap: 8px;
        background: rgba(255,255,255,0.10); border: 1px solid rgba(255,255,255,0.24);
        padding: 7px 14px; border-radius: 999px; font-size: 12.5px; font-weight: 600; letter-spacing: 0.2px;
        color: #fff;
    }}
    .qt-period-chip .dot {{
        width: 7px; height: 7px; border-radius: 50%; background: {GREEN};
        box-shadow: 0 0 0 3px rgba(46,125,50,0.30);
    }}
    .qt-cutoff-note {{ margin-top: 8px; font-size: 11.5px; color: #a9bccf; }}

    /* ---------- section titles ---------- */
    .qt-section-title {{ display: flex; align-items: center; gap: 10px; margin: 30px 2px 14px; }}
    .qt-section-title .chip {{
        width: 7px; height: 7px; border-radius: 2px; background: {BLUE}; flex: none;
        box-shadow: 0 0 0 3px rgba(9,127,187,0.14);
    }}
    .qt-section-title h2 {{
        font-size: 13.5px; margin: 0; font-weight: 700; letter-spacing: 0.6px; color: {NAVY};
        text-transform: uppercase; white-space: nowrap;
    }}
    .qt-section-title .rule {{ flex: 1; height: 1px; background: linear-gradient(90deg, {HAIR} 0%, rgba(225,224,217,0.15) 100%); }}
    .qt-section-title .tag {{
        font-size: 10.5px; color: {STEEL}; font-weight: 600; white-space: nowrap;
        background: rgba(76,92,110,0.08); padding: 3px 9px; border-radius: 999px;
    }}

    /* ---------- KPI cards — sem preenchimento, só contorno; o azul do
       painel principal aparece por trás de indicadores, painéis e tabelas ---------- */
    .qt-kpi-card {{
        background: transparent; border: 1px solid rgba(33,47,86,0.16); border-radius: 18px;
        padding: 16px 18px 15px; position: relative; overflow: hidden; height: 100%;
        box-shadow: 0 6px 16px rgba(33,47,86,0.05);
        transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
    }}
    .qt-kpi-card:hover {{
        transform: translateY(-2px); border-color: rgba(33,47,86,0.28);
        box-shadow: 0 10px 22px rgba(33,47,86,0.10);
    }}
    .qt-kpi-card::before {{
        content: ""; position: absolute; left: 0; top: 10px; bottom: 10px; width: 4px; border-radius: 3px;
        background: var(--kpi-accent, {BLUE});
    }}
    .qt-kpi-card::after {{
        content: ""; position: absolute; right: -30px; top: -30px; width: 90px; height: 90px; border-radius: 50%;
        background: var(--kpi-accent, {BLUE}); opacity: 0.07; pointer-events: none;
    }}
    .qt-kpi-label {{
        font-size: 10.5px; color: {NAVY}; opacity: 0.62; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.5px; margin-bottom: 7px;
    }}
    .qt-kpi-value {{ font-size: 26px; font-weight: 800; letter-spacing: -0.6px; color: {TEXT_PRIMARY}; line-height: 1; }}
    .qt-kpi-value .unit {{ font-size: 13px; font-weight: 600; color: {TEXT_SECONDARY}; margin-left: 2px; }}
    .qt-kpi-sub {{ margin-top: 8px; font-size: 11.5px; color: {TEXT_SECONDARY}; display: flex; align-items: center; gap: 5px; }}
    .qt-delta {{ font-weight: 600; padding: 0; border-radius: 0; font-size: 11.5px; background: transparent !important; }}
    .qt-delta.neutral {{ color: {TEXT_SECONDARY}; }}
    .qt-delta.good {{ color: {GREEN}; }}
    .qt-delta.bad {{ color: {RED}; }}
    .qt-kpi-bar-track {{ margin-top: 10px; height: 5px; border-radius: 999px; background: rgba(33,47,86,0.12); overflow: hidden; }}
    .qt-kpi-bar-fill {{ height: 100%; border-radius: 999px; background: var(--kpi-accent, {BLUE}); transition: width 0.4s ease; }}

    /* ---------- panels (bordered containers wrapping charts) ----------
       Streamlit's stable "stVerticalBlockBorderWrapper" testid is gone in
       current versions; panel() now assigns every container an explicit
       key, which Streamlit mirrors as a "st-key-qtpanel_*" class — target
       that instead. Transparent fill on purpose: the page's blue gradient
       shows through every panel, card and table — only a thin outline
       separates sections now. */
    div[data-testid="stVerticalBlock"][class*="st-key-qtpanel_"] {{
        background: transparent !important; border: 1px solid rgba(33,47,86,0.14) !important;
        border-radius: 20px !important;
        box-shadow: 0 6px 18px rgba(33,47,86,0.05);
        overflow: hidden; transition: box-shadow 0.2s ease, border-color 0.2s ease;
    }}
    div[data-testid="stVerticalBlock"][class*="st-key-qtpanel_"]:hover {{
        border-color: rgba(33,47,86,0.26);
        box-shadow: 0 10px 26px rgba(33,47,86,0.09);
    }}
    /* fallback for older Streamlit builds that still use the classic testid */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: transparent !important; border: 1px solid rgba(33,47,86,0.14) !important;
        border-radius: 20px !important;
        box-shadow: 0 6px 18px rgba(33,47,86,0.05);
        overflow: hidden;
    }}
    .qt-panel-title {{ font-size: 13.5px; margin: 0 0 2px; font-weight: 700; color: {TEXT_PRIMARY}; }}
    .qt-panel-desc {{ font-size: 11.5px; color: {STEEL}; margin-bottom: 6px; line-height: 1.45; }}

    /* ---------- tables — moldura translúcida, sem preenchimento; a grade
       nativa do Streamlit assume o tom azulado definido no theme (config.toml) ---------- */
    [data-testid="stDataFrame"] {{
        border: 1px solid rgba(33,47,86,0.14); border-radius: 14px; overflow: hidden; background: transparent;
    }}
    [data-testid="stElementContainer"]:has([data-testid="stDataFrame"]) {{ background: transparent; }}

    /* ---------- qt-table — tabela HTML própria, 100% transparente inclusive
       no cabeçalho (troca o st.dataframe nos 3 resumos: o canvas nativo do
       Streamlit ignora o tema nessa linha e sempre pinta um branco fixo) ---------- */
    .qt-table-scroll {{ width: 100%; overflow-x: auto; }}
    .qt-table {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
    .qt-table thead th {{
        text-align: left; font-weight: 700; color: {NAVY}; text-transform: uppercase;
        font-size: 10px; letter-spacing: 0.4px; padding: 10px 12px; white-space: nowrap;
        border-bottom: 1px solid rgba(33,47,86,0.22); background: transparent;
    }}
    .qt-table thead th.num {{ text-align: right; }}
    .qt-table tbody td {{
        padding: 9px 12px; color: {TEXT_PRIMARY}; border-bottom: 1px solid rgba(33,47,86,0.08);
        background: transparent; white-space: nowrap;
    }}
    .qt-table tbody td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .qt-table tbody tr:last-child td {{ border-bottom: none; }}
    .qt-table tbody tr:hover td {{ background: rgba(9,127,187,0.06); }}

    /* ---------- footer ---------- */
    .qt-disclaimer {{
        margin-top: 10px; padding: 12px 14px; background: #FFF7EE; border: 1px solid #F3DCC0;
        border-radius: 14px; font-size: 11.5px; color: #7a4d1e; line-height: 1.55;
    }}
    .qt-footer-note {{ font-size: 11px; color: {GREY}; margin-top: 12px; text-align: center; }}
</style>
""", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Data loading (cache invalidada automaticamente quando o arquivo muda)
# --------------------------------------------------------------------------- #

@st.cache_data(show_spinner="Lendo Performance_Ago.xlsx…")
def _load(path: str, mtime: float) -> DashboardData:
    return load_workbook_data(path)


def load_data(path: str) -> DashboardData | None:
    if not path or not os.path.isfile(path):
        return None
    mtime = os.path.getmtime(path)
    return _load(path, mtime)


@st.cache_data(show_spinner="Lendo base de Recursos (RH)…")
def _load_wf(path: str, mtime: float) -> WorkforceData:
    return load_workforce_data(path)


def load_workforce(path: str) -> WorkforceData | None:
    if not path or not os.path.isfile(path):
        return None
    mtime = os.path.getmtime(path)
    return _load_wf(path, mtime)


@st.cache_data(show_spinner="Lendo base de Escalas & Embarques…")
def _load_emb(path: str, mtime: float) -> EmbarquesData:
    return load_embarques_data(path)


def load_embarques(path: str) -> EmbarquesData | None:
    if not path or not os.path.isfile(path):
        return None
    mtime = os.path.getmtime(path)
    return _load_emb(path, mtime)


# --------------------------------------------------------------------------- #
# Sidebar — fonte de dados, filtros, atualização
# --------------------------------------------------------------------------- #

with st.sidebar:
    if LOGO_B64:
        st.markdown(
            f'<img src="data:image/png;base64,{LOGO_B64}" '
            f'style="height:30px;filter:invert(1) grayscale(1) brightness(0.3);margin-bottom:4px;"/>',
            unsafe_allow_html=True,
        )
    st.caption("Dashboard de Performance Operacional")
    st.divider()

    st.markdown("**Fonte de dados**")
    folder = st.text_input(
        "Pasta sincronizada (OneDrive/SharePoint)",
        value=DEFAULT_PATH,
        help="Pasta do mês (ex.: .../9. Performance/Ago-2026) ou a pasta pai "
             "'9. Performance' para escolher entre vários meses.",
    )

    workbooks = list_available_workbooks(folder)
    chosen_path = None
    if workbooks:
        labels = [os.path.relpath(w, folder) for w in workbooks]
        idx = st.selectbox("Arquivo", options=range(len(workbooks)),
                            format_func=lambda i: labels[i])
        chosen_path = workbooks[idx]
    else:
        uploaded = st.file_uploader("Ou envie o .xlsx manualmente", type=["xlsx"])
        if uploaded is not None:
            tmp_path = os.path.join("/tmp", uploaded.name)
            with open(tmp_path, "wb") as f:
                f.write(uploaded.getbuffer())
            chosen_path = tmp_path
        else:
            st.warning("Pasta não encontrada nesta máquina. Ajuste o caminho acima "
                        "ou envie o arquivo manualmente.")

    st.markdown("**Base de Recursos (RH)**")
    wf_workbooks = list_available_workforce_workbooks(folder)
    chosen_wf_path = None
    if wf_workbooks:
        wf_labels = [os.path.relpath(w, folder) for w in wf_workbooks]
        wf_idx = st.selectbox("Arquivo de RH", options=range(len(wf_workbooks)),
                               format_func=lambda i: wf_labels[i], key="wf_select")
        chosen_wf_path = wf_workbooks[wf_idx]
    else:
        wf_uploaded = st.file_uploader("Ou envie o Planejamento_Ativos_*.xlsx", type=["xlsx"], key="wf_uploader")
        if wf_uploaded is not None:
            tmp_wf_path = os.path.join("/tmp", wf_uploaded.name)
            with open(tmp_wf_path, "wb") as f:
                f.write(wf_uploaded.getbuffer())
            chosen_wf_path = tmp_wf_path
        else:
            st.caption("Opcional — habilita a seção *Recursos & Pessoas*. Procurado na mesma "
                       "pasta acima, em arquivos com 'ativos' ou 'planejamento' no nome.")

    st.markdown("**Base de Escalas & Embarques**")
    emb_workbooks = list_available_embarques_workbooks(folder)
    chosen_emb_path = None
    if emb_workbooks:
        emb_labels = [os.path.relpath(w, folder) for w in emb_workbooks]
        emb_idx = st.selectbox("Arquivo de Escalas", options=range(len(emb_workbooks)),
                                format_func=lambda i: emb_labels[i], key="emb_select")
        chosen_emb_path = emb_workbooks[emb_idx]
    else:
        emb_uploaded = st.file_uploader("Ou envie o Relatório Mensal de Eventos*.xlsx", type=["xlsx"], key="emb_uploader")
        if emb_uploaded is not None:
            tmp_emb_path = os.path.join("/tmp", emb_uploaded.name)
            with open(tmp_emb_path, "wb") as f:
                f.write(emb_uploaded.getbuffer())
            chosen_emb_path = tmp_emb_path
        else:
            st.caption("Opcional — habilita a seção *Escalas & Embarques*. Procurado na mesma "
                       "pasta acima, em arquivos com 'eventos', 'embarque' ou 'escala' no nome.")

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔄 Atualizar agora", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with c2:
        auto = st.toggle("Auto (60s)", value=False, help="Re-lê a planilha automaticamente a cada 60s")
    if auto:
        st_autorefresh(interval=60_000, key="auto_refresh")

    data = load_data(chosen_path) if chosen_path else None
    data_wf = load_workforce(chosen_wf_path) if chosen_wf_path else None
    data_emb = load_embarques(chosen_emb_path) if chosen_emb_path else None

    if data is not None:
        st.caption(f"📄 {os.path.basename(data.path)}")
        st.caption(f"🕒 Atualizado em {time.strftime('%d/%m/%Y %H:%M', time.localtime(data.mtime))}")
    if data_wf is not None:
        st.caption(f"👥 {os.path.basename(data_wf.path)}")
    if data_emb is not None:
        st.caption(f"🚢 {os.path.basename(data_emb.path)}")

    st.divider()
    st.markdown("**Filtros**")
    if data is not None and not data.wo.empty:
        coords = sorted(data.wo["coord"].dropna().unique().tolist())
        clientes = sorted(data.wo["cliente"].dropna().unique().tolist())
        tipos_wo = sorted(data.wo["tipo_wo"].dropna().unique().tolist())
        tipos = sorted(data.wo["tipo"].dropna().unique().tolist())
    else:
        coords = clientes = tipos_wo = tipos = []

    f_coord = st.multiselect("Coordenador", coords, default=[])
    f_cliente = st.multiselect("Cliente", clientes, default=[])
    f_tipo_wo = st.multiselect("Tipo de Contrato", tipos_wo, default=[])
    f_tipo = st.multiselect("Tipo de Serviço", tipos, default=[])


def logo_header_html(title: str, subtitle: str, period_label: str, cutoff_note: str, live: bool) -> str:
    logo_img = (f'<img src="data:image/png;base64,{LOGO_B64}"/>' if LOGO_B64
                else '<div style="font-weight:800;font-size:20px;color:#fff;">Q</div>')
    dot_color = GREEN if live else BASE_GREY
    status = "Dados ao vivo" if live else "Aguardando arquivo"
    return f"""
    <div class="qt-brandbar">
      <div class="qt-brand-mark">
        {logo_img}
        <div class="qt-brand-word"><span>Inspection · Repair · Maintenance</span></div>
      </div>
      <div style="flex:1;min-width:220px;">
        <h1>{title}</h1>
        <div class="sub">{subtitle}</div>
      </div>
      <div class="qt-brandbar-right">
        <span class="qt-period-chip"><span class="dot" style="background:{dot_color};"></span> {status}
        {(' · ' + period_label) if period_label else ''}</span>
        <div class="qt-cutoff-note">{cutoff_note}</div>
      </div>
    </div>
    """


if data is None:
    st.markdown(logo_header_html(
        "QUALITECH — Dashboard de Performance Operacional",
        "Aguardando arquivo de dados",
        "", "", live=False,
    ), unsafe_allow_html=True)
    st.info("Informe, na barra lateral, o caminho da pasta sincronizada do OneDrive/SharePoint que contém o "
            "Performance_*.xlsx — ou envie o arquivo manualmente para pré-visualizar.")
    st.stop()

wo = data.wo.copy()
if f_coord:
    wo = wo[wo["coord"].isin(f_coord)]
if f_cliente:
    wo = wo[wo["cliente"].isin(f_cliente)]
if f_tipo_wo:
    wo = wo[wo["tipo_wo"].isin(f_tipo_wo)]
if f_tipo:
    wo = wo[wo["tipo"].isin(f_tipo)]
filtered = bool(f_coord or f_cliente or f_tipo_wo or f_tipo)

# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #

st.markdown(logo_header_html(
    "Dashboard de Performance Operacional",
    f"Consolidado de indicadores — baseado nas tabelas dinâmicas de {os.path.basename(data.path)}",
    data.periodo_label,
    f"Atualizado em {time.strftime('%d/%m/%Y %H:%M', time.localtime(data.mtime))}"
    + (" · <b>filtros ativos</b>" if filtered else ""),
    live=True,
), unsafe_allow_html=True)

fmt0 = lambda v: f"{v:,.0f}".replace(",", ".")
fmt1p = lambda v: f"{v*100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
fmt0p = lambda v: f"{v*100:,.0f}%"

CHART_FONT = dict(family="system-ui, -apple-system, Segoe UI, Roboto, sans-serif", color=TEXT_SECONDARY, size=12)
HOVER_STYLE = dict(bgcolor=NAVY, font=dict(color="white", size=12, family=CHART_FONT["family"]),
                    bordercolor=NAVY)
# painéis não têm mais fundo branco por trás dos gráficos — o grid/eixo precisa
# de contraste próprio contra o azul da página, em vez do HAIR (quase branco).
GRID_ON_PAGE = "rgba(33,47,86,0.14)"
AXIS_ON_PAGE = "rgba(33,47,86,0.35)"


def base_layout(fig, height=280, legend=True):
    fig.update_layout(
        height=height, margin=dict(t=10, b=10, l=10, r=10), font=CHART_FONT,
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", hoverlabel=HOVER_STYLE,
        legend=dict(orientation="h", y=-0.14, x=0.5, xanchor="center") if legend else None,
        showlegend=legend,
    )
    fig.update_xaxes(gridcolor=GRID_ON_PAGE, zeroline=False, linecolor=AXIS_ON_PAGE)
    fig.update_yaxes(gridcolor=GRID_ON_PAGE, zeroline=False, linecolor=AXIS_ON_PAGE)
    return fig


def donut(labels, values, colors, center_title=None, center_sub=None, height=270, value_fmt=None):
    text = [value_fmt(v) for v in values] if value_fmt else [fmt0(v) for v in values]
    fig = go.Figure(data=[go.Pie(
        labels=labels, values=values, hole=0.62, marker=dict(colors=colors, line=dict(color="white", width=2)),
        text=text, textinfo="text", textfont=dict(size=13, color="white", family="Arial Black"),
        hoverlabel=HOVER_STYLE,
    )])
    fig.update_layout(
        showlegend=True, legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center"),
        margin=dict(t=10, b=10, l=10, r=10), height=height, font=CHART_FONT,
        paper_bgcolor="rgba(0,0,0,0)",
        annotations=[dict(text=f"<b>{center_title}</b><br><span style='font-size:11px;color:{GREY}'>{center_sub}</span>",
                           x=0.5, y=0.5, showarrow=False, font=dict(size=21, color=NAVY))] if center_title else [],
    )
    return fig


def section_title(title: str, tag: str = ""):
    tag_html = f'<div class="tag">{tag}</div>' if tag else ""
    st.markdown(f"""
    <div class="qt-section-title"><div class="chip"></div><h2>{title}</h2><div class="rule"></div>{tag_html}</div>
    """, unsafe_allow_html=True)


def kpi_card(col, label, value, unit="", delta_text="", accent=BLUE, bar_pct=None):
    bar_html = (f'<div class="qt-kpi-bar-track"><div class="qt-kpi-bar-fill" '
                f'style="width:{max(0, min(100, bar_pct*100)):.0f}%;--kpi-accent:{accent};"></div></div>'
                if bar_pct is not None else "")
    col.markdown(f"""
    <div class="qt-kpi-card" style="--kpi-accent:{accent};">
      <div class="qt-kpi-label">{label}</div>
      <div class="qt-kpi-value">{value}<span class="unit">{unit}</span></div>
      <div class="qt-kpi-sub"><span class="qt-delta neutral">{delta_text}</span></div>
      {bar_html}
    </div>
    """, unsafe_allow_html=True)


_panel_key_counter = itertools.count()


def panel(title: str = "", desc: str = ""):
    """Bordered container styled like the HTML `.panel` card; use as a `with` block.

    Streamlit's bordered `st.container(border=True)` no longer exposes a stable
    `stVerticalBlockBorderWrapper` testid in current versions — the border is
    applied via an emotion-cache class that changes between builds. To style it
    reliably we give every panel an explicit `key`, which Streamlit always
    mirrors onto the DOM as a stable `st-key-<key>` class (see CSS below).
    """
    key = f"qtpanel_{next(_panel_key_counter)}"
    c = st.container(border=True, key=key)
    if title or desc:
        with c:
            title_html = f'<div class="qt-panel-title">{title}</div>' if title else ""
            desc_html = f'<div class="qt-panel-desc">{desc}</div>' if desc else ""
            st.markdown(title_html + desc_html, unsafe_allow_html=True)
    return c


def _fmt_cell(v):
    """Formata uma célula de tabela seguindo o mesmo padrão numérico dos KPIs
    (milhar com ponto), preservando strings já formatadas (ex.: '62,7%')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return fmt0(v) if float(v).is_integer() else f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(v)


def html_table(df: pd.DataFrame, right_cols=None):
    """Tabela em HTML puro — substitui st.dataframe() nas 3 tabelas de resumo.

    st.dataframe() é desenhado num <canvas> (glide-data-grid) cujo cabeçalho
    lê uma cor quase-branca fixa, ignorando o tema/CSS da página — por isso,
    mesmo com o resto do dashboard transparente, a linha de título da tabela
    continuava branca. Uma tabela HTML comum dá controle total via CSS
    (classe .qt-table, abaixo) e mantém o mesmo fundo transparente do painel.
    """
    right_cols = set(right_cols or [])
    cols = list(df.columns)
    thead = "".join(
        f'<th class="{"num" if c in right_cols else ""}">{html.escape(str(c))}</th>' for c in cols
    )
    body_rows = []
    for _, row in df.iterrows():
        tds = "".join(
            f'<td class="{"num" if c in right_cols else ""}">{html.escape(_fmt_cell(row[c]))}</td>' for c in cols
        )
        body_rows.append(f"<tr>{tds}</tr>")
    st.markdown(
        f'<div class="qt-table-scroll"><table class="qt-table"><thead><tr>{thead}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody></table></div>',
        unsafe_allow_html=True,
    )

tab1, tab2, tab3 = st.tabs([
    "📊 Performance Operacional", "👥 Recursos & Pessoas", "🚢 Escalas & Embarques",
])

with tab1:
    # --------------------------------------------------------------------------- #
    # Indicadores-chave (KPIs)
    # --------------------------------------------------------------------------- #

    t = data.totais
    n_wo = len(wo) if filtered else data.n_wo
    util, ocio, last_reading = data.pct_utilizacao_ociosidade
    sispat = data.sispat_summary
    por_tipo_wo = data.por_tipo_wo if not filtered else (
        wo["tipo_wo"].value_counts().rename_axis("tipo_wo").reset_index(name="n")
        .assign(pct=lambda d: d["n"] / d["n"].sum() if len(wo) else 0)
    )
    fixo_n = int(por_tipo_wo.loc[por_tipo_wo["tipo_wo"] == "Fixo", "n"].sum()) if len(por_tipo_wo) else 0
    var_n = int(por_tipo_wo.loc[por_tipo_wo["tipo_wo"] == "Variável", "n"].sum()) if len(por_tipo_wo) else 0
    spot_n = int(por_tipo_wo.loc[por_tipo_wo["tipo_wo"] == "Spot", "n"].sum()) if len(por_tipo_wo) else 0
    tot_n = max(1, fixo_n + var_n + spot_n)

    section_title("Indicadores-chave (KPIs)", data.periodo_label)
    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, "Man-days Realizados (Acum.)", fmt0(t["real_acum"]), " md",
             f"{fmt1p(t['real_acum']/t['real_prog']) if t['real_prog'] else '—'} do previsto",
             accent=BLUE, bar_pct=(t["real_acum"]/t["real_prog"]) if t["real_prog"] else None)
    kpi_card(c2, "Utilização × Ociosidade", fmt1p(util) if util is not None else "—", "",
             f"Ociosidade: {fmt1p(ocio)}" if ocio is not None else "sem leitura",
             accent=BLUE, bar_pct=util)
    kpi_card(c3, "Backlog Pendente", fmt0(data.pendente_total), " md",
             f"Já lançado: {fmt0(data.lancado_total)} md", accent=AMBER)
    kpi_card(c4, "Conformidade SISPAT", fmt0p(sispat["sim"]/sispat["total"]) if sispat["total"] else "—", "",
             f"{sispat['sim']} de {sispat['total']} plataformas", accent=GREEN,
             bar_pct=(sispat["sim"]/sispat["total"]) if sispat["total"] else None)

    c5, c6, c7, c8 = st.columns(4)
    kpi_card(c5, "Diárias Fixas", str(fixo_n), " WOs", f"{fmt1p(fixo_n/tot_n)} do total", accent=GREEN, bar_pct=fixo_n/tot_n)
    kpi_card(c6, "Diárias Variáveis", str(var_n), " WOs", f"{fmt1p(var_n/tot_n)} do total", accent=AMBER, bar_pct=var_n/tot_n)
    kpi_card(c7, "Diárias Spot", str(spot_n), " WOs", f"{fmt1p(spot_n/tot_n)} do total", accent=RED, bar_pct=spot_n/tot_n)
    kpi_card(c8, "Ordens de Serviço", str(n_wo), " WOs",
             f"ILDD {fmt0p(data.ildd_medio)} · IAP {fmt0p(data.iap_medio)}" if data.ildd_medio else "—", accent=NAVY)

    # --------------------------------------------------------------------------- #
    # Mix de Contratos & Ocupação da Frota
    # --------------------------------------------------------------------------- #

    section_title("Mix de Contratos &amp; Utilização da Equipe", "Fixa · Variável · Spot · Utilização")
    g1, g2, g3 = st.columns(3)
    with g1:
        with panel("Diárias por Tipo de Contrato", "Ordens de serviço do mês — Fixa × Variável × Spot"):
            st.plotly_chart(donut(["Fixa", "Variável", "Spot"], [fixo_n, var_n, spot_n],
                                   [GREEN, AMBER, RED], center_title=tot_n, center_sub="OS"),
                             use_container_width=True, key="donut_mix")
    with g2:
        with panel("Inspeção × Manutenção por Tipo de Contrato", "Composição de cada tipo de contrato por natureza do serviço"):
            cross = data.cross_tipowo_tipo if not filtered else pd.crosstab(wo["tipo_wo"], wo["tipo"])
            if not cross.empty:
                fig = go.Figure()
                colors_map = {"Inspeção": BLUE, "Manutenção": NAVY, "Outros": BASE_GREY}
                for col in cross.columns:
                    fig.add_bar(y=cross.index, x=cross[col], name=col, orientation="h",
                                marker_color=colors_map.get(col, BASE_GREY))
                fig.update_layout(barmode="stack", barnorm="percent")
                base_layout(fig, height=270)
                st.plotly_chart(fig, use_container_width=True, key="stacked_cross")
            else:
                st.caption("Sem dados suficientes com os filtros atuais.")
    with g3:
        desc = "Última leitura do mês" if util is not None else "Sem leitura disponível"
        with panel("Utilização × Ociosidade da Equipe", desc):
            if util is not None:
                st.plotly_chart(donut(["Utilização", "Ociosidade"], [util, ocio], [BLUE, BASE_GREY],
                                       center_title=fmt0p(util), center_sub="utilização", value_fmt=fmt0p),
                                 use_container_width=True, key="donut_util")
                if last_reading:
                    st.caption(f"Última leitura: {pd.Timestamp(last_reading['date']).strftime('%d/%m')} · "
                               f"{int(last_reading['ativo'] or 0)} colaboradores no efetivo ativo")
            else:
                st.caption("Sem leitura de utilização disponível.")

    # --------------------------------------------------------------------------- #
    # Evolução da Operação
    # --------------------------------------------------------------------------- #

    section_title("Evolução da Operação", "Man-days / dia")
    ce1, ce2 = st.columns(2)
    with ce1:
        with panel("Evolução Diária — Programado × Realizado", f"Man-days programados vs. executados por dia ({data.periodo_label})"):
            if not data.evolucao_diaria.empty:
                fig = go.Figure()
                fig.add_scatter(x=data.evolucao_diaria["dia"], y=data.evolucao_diaria["programado"],
                                 name="Programado", line=dict(color=BLUE, dash="dot", width=2))
                fig.add_scatter(x=data.evolucao_diaria["dia"], y=data.evolucao_diaria["real_prog"],
                                 name="Realizado", line=dict(color=GREEN, width=2.5))
                base_layout(fig, height=300)
                st.plotly_chart(fig, use_container_width=True, key="evo_diaria")
            else:
                st.caption("Sem dados de evolução diária.")
    with ce2:
        with panel("Evolução Acumulada — Programado × Realizado", "Curva de avanço acumulado do mês, com meta programada como referência"):
            if not data.evolucao_acumulada.empty:
                fig = go.Figure()
                fig.add_scatter(x=data.evolucao_acumulada["dia"], y=data.evolucao_acumulada["programado_acum"],
                                 name="Programado Acum.", line=dict(color=BLUE, dash="dot", width=2))
                fig.add_scatter(x=data.evolucao_acumulada["dia"], y=data.evolucao_acumulada["real_acum"],
                                 name="Real Acum.", line=dict(color=GREEN, width=2.5), fill="tozeroy",
                                 fillcolor="rgba(46,125,50,0.08)")
                base_layout(fig, height=300)
                st.plotly_chart(fig, use_container_width=True, key="evo_acum")
            else:
                st.caption("Sem dados de evolução acumulada.")

    # --------------------------------------------------------------------------- #
    # Mix Operacional & Conformidade
    # --------------------------------------------------------------------------- #

    section_title("Mix Operacional &amp; Conformidade", "Tipo · SISPAT · Histórico")
    h1, h2, h3 = st.columns(3)
    with h1:
        with panel("Man-days Realizados por Tipo", "Inspeção vs. Manutenção (Real Acumulado)"):
            if not data.por_tipo.empty:
                fig = donut(data.por_tipo["tipo"].tolist(), data.por_tipo["real_acum"].tolist(),
                            SERIES[:len(data.por_tipo)], center_title=fmt0(data.por_tipo["real_acum"].sum()), center_sub="md")
                st.plotly_chart(fig, use_container_width=True, key="donut_tipo")
            else:
                st.caption("Sem dados por tipo.")
    with h2:
        with panel("Taxa de Utilização — Histórico Mensal", "Realizado ÷ Capacidade da equipe (efetivo disponível, fora do ciclo de folga)"):
            if not data.utilizacao_mensal.empty:
                fig = go.Figure()
                fig.add_bar(x=data.utilizacao_mensal["mes"], y=data.utilizacao_mensal["taxa"], marker_color=BLUE)
                fig.update_yaxes(tickformat=".0%")
                base_layout(fig, height=270, legend=False)
                st.plotly_chart(fig, use_container_width=True, key="util_mensal")
            else:
                st.caption("Sem histórico mensal de utilização.")
    with h3:
        with panel("Conformidade SISPAT", "Plataformas cadastradas no sistema regulatório"):
            if sispat["total"]:
                fig = donut(["Sim", "Não"], [sispat["sim"], sispat["nao"]], [GREEN, RED],
                            center_title=fmt0p(sispat["sim"]/sispat["total"]), center_sub="conforme")
                st.plotly_chart(fig, use_container_width=True, key="donut_sispat")
            else:
                st.caption("Sem dados de SISPAT.")

    # --------------------------------------------------------------------------- #
    # Produtividade por Coordenador
    # --------------------------------------------------------------------------- #

    section_title("Produtividade por Coordenador", "Man-days")
    pc1, pc2 = st.columns(2)
    with pc1:
        with panel("Real Acumulado por Coordenador", "Man-days efetivamente executados no mês, por coordenador responsável"):
            if not data.por_coordenador.empty:
                df = data.por_coordenador.sort_values("real_acum", ascending=True)
                fig = go.Figure(go.Bar(x=df["real_acum"], y=df["coord"], orientation="h", marker_color=BLUE))
                base_layout(fig, height=300, legend=False)
                st.plotly_chart(fig, use_container_width=True, key="coord_bar")
            else:
                st.caption("Sem dados por coordenador.")
    with pc2:
        with panel("Backlog de Logística por Coordenador", "Man-days programados ainda pendentes de lançamento (Ajuste_Mês)"):
            if not data.backlog_coordenador.empty:
                df = data.backlog_coordenador.sort_values("backlog_logistica", ascending=True)
                fig = go.Figure(go.Bar(x=df["backlog_logistica"], y=df["coord"], orientation="h", marker_color=RED))
                base_layout(fig, height=300, legend=False)
                st.plotly_chart(fig, use_container_width=True, key="backlog_bar")
            else:
                st.caption("Sem dados de backlog.")

    section_title("Resumo Consolidado por Coordenador", "Tabela")
    with panel("", ""):
        if not data.por_coordenador.empty:
            show = data.por_coordenador.copy()
            show["% Avanço"] = (show["real_acum"] / show["real_prog"]).map(
                lambda v: fmt1p(v) if pd.notna(v) and v != float("inf") else "—")
            show = show.rename(columns={"coord": "Coordenador", "planejado": "Planejado", "programado": "Programado",
                                         "real_prog": "Real+Prog", "real_acum": "Real Acum", "ajuste_mes": "Ajuste_Mês"})
            html_table(show[["Coordenador", "Planejado", "Programado", "Real+Prog", "Real Acum", "Ajuste_Mês", "% Avanço"]],
                       right_cols=["Planejado", "Programado", "Real+Prog", "Real Acum", "Ajuste_Mês", "% Avanço"])
        else:
            st.caption("Sem dados por coordenador.")

    # --------------------------------------------------------------------------- #
    # Carteira de Clientes
    # --------------------------------------------------------------------------- #

    section_title("Carteira de Clientes", "Top clientes · Movimentação de WOs")
    cl1, cl2 = st.columns(2)
    with cl1:
        with panel("Top Clientes — Man-days Realizados", 'Principais clientes por Real Acumulado no mês (demais agrupados em "Outros")'):
            if not data.por_cliente.empty:
                top = data.por_cliente.sort_values("real_acum", ascending=False)
                top8 = top.head(8)
                outros_val = top.iloc[8:]["real_acum"].sum() if len(top) > 8 else 0
                labels = top8["cliente"].tolist() + (["Outros"] if outros_val else [])
                values = top8["real_acum"].tolist() + ([outros_val] if outros_val else [])
                fig = go.Figure(go.Bar(x=values[::-1], y=labels[::-1], orientation="h", marker_color=BLUE))
                base_layout(fig, height=340, legend=False)
                st.plotly_chart(fig, use_container_width=True, key="cliente_bar")
            else:
                st.caption("Sem dados por cliente.")
    with cl2:
        with panel("Movimentação de WOs por Cliente", "Ordens de serviço que entraram (+) e saíram (−) da carteira no período"):
            if not data.movimentacao_clientes.empty:
                m = data.movimentacao_clientes
                m = m[(m["entraram"] != 0) | (m["sairam"] != 0)]
                if not m.empty:
                    fig = go.Figure()
                    fig.add_bar(y=m["cliente"], x=m["entraram"], name="Entraram", orientation="h", marker_color=GREEN)
                    fig.add_bar(y=m["cliente"], x=m["sairam"], name="Saíram", orientation="h", marker_color=RED)
                    fig.update_layout(barmode="relative")
                    base_layout(fig, height=340)
                    st.plotly_chart(fig, use_container_width=True, key="mov_bar")
                else:
                    st.caption("Sem movimentação no período.")
            else:
                st.caption("Sem dados de movimentação de clientes.")

    # --------------------------------------------------------------------------- #
    # Plano Estratégico
    # --------------------------------------------------------------------------- #

    section_title("Plano Estratégico — Meta × Comercial × Real", "Man-days")
    with panel("", ""):
        if not data.plano_estrategico.empty:
            pe = data.plano_estrategico
            fig = go.Figure()
            fig.add_bar(x=pe["mes"], y=pe["meta"], name="Meta", marker_color=BASE_GREY)
            fig.add_bar(x=pe["mes"], y=pe["comercial"], name="Comercial", marker_color=AMBER)
            fig.add_bar(x=pe["mes"], y=pe["real"], name="Real", marker_color=GREEN)
            fig.update_layout(barmode="group")
            base_layout(fig, height=300)
            st.plotly_chart(fig, use_container_width=True, key="plano_bar")
        else:
            st.caption("Sem dados de plano estratégico.")

    section_title("Resumo Consolidado por Cliente", "Tabela")
    with panel("", ""):
        if not data.por_cliente.empty:
            show = data.por_cliente.copy().sort_values("real_acum", ascending=False)
            show = show.rename(columns={"cliente": "Cliente", "planejado": "Planejado", "programado": "Programado",
                                         "real_prog": "Real+Prog", "real_acum": "Real Acum"})
            html_table(show[["Cliente", "Planejado", "Programado", "Real+Prog", "Real Acum"]],
                       right_cols=["Planejado", "Programado", "Real+Prog", "Real Acum"])
        else:
            st.caption("Sem dados por cliente.")

with tab2:
    # --------------------------------------------------------------------------- #
    # Recursos & Pessoas (RH) — efetivo, turnover, certificações
    # --------------------------------------------------------------------------- #

    if data_wf is not None:
        section_title("Recursos &amp; Pessoas", data_wf.periodo_label)
        st.caption("Base de efetivo (RH) — indicadores de headcount, complementares aos indicadores "
                   "de diárias/WO acima. Não é afetada pelos filtros da barra lateral.")

        tv_rate = data_wf.turnover_rate_last
        tv_last = data_wf.turnover_last
        vencendo = data_wf.certs_vencendo_90d_n
        vencidas = data_wf.certs_vencidas_n

        r1c1, r1c2, r1c3 = st.columns(3)
        kpi_card(r1c1, "Efetivo Ativo", fmt0(data_wf.headcount_ativo), " pessoas",
                 f"{data_wf.headcount_inativo} inativos · {data_wf.headcount_afastado} afastados",
                 accent=BLUE)
        kpi_card(r1c2, "Líderes Designados", fmt0(data_wf.lideres_n), " pessoas",
                 f"{fmt1p(data_wf.pct_lideres)} do efetivo ativo", accent=NAVY, bar_pct=data_wf.pct_lideres)
        kpi_card(r1c3, "Novos Admitidos (2026)", fmt0(data_wf.novos_admitidos_n), " pessoas",
                 "pipeline de contratações do ano", accent=GREEN)

        r2c1, r2c2, r2c3 = st.columns(3)
        kpi_card(r2c1, "Turnover (último mês)", fmt1p(tv_rate) if tv_rate is not None else "—", "",
                 (f"{tv_last['label']}: {fmt0(tv_last['admitidos'])} admitidos · "
                  f"{fmt0((tv_last['demitidos'] or 0) + (tv_last['pedido_demissao'] or 0))} saídas") if tv_last else "sem dados",
                 accent=AMBER)
        kpi_card(r2c2, "Colaboradores Bloqueados", fmt0(data_wf.bloqueios_n), " pessoas",
                 "sem acesso liberado em ao menos 1 cliente/sonda", accent=RED)
        kpi_card(r2c3, "Certificações a Vencer (90d)", fmt0(vencendo), "",
                 f"{fmt0(vencidas)} já vencidas — risco de compliance" if vencidas else "nenhuma vencida",
                 accent=RED if (vencendo or vencidas) else GREEN)

        w1, w2, w3 = st.columns(3)
        with w1:
            with panel("Efetivo por Natureza", "Headcount ativo — Inspeção × Manutenção"):
                if not data_wf.por_tipo_mo.empty:
                    dfm = data_wf.por_tipo_mo
                    fig = donut(dfm["tipo"].tolist(), dfm["n"].tolist(), SERIES[:len(dfm)],
                                center_title=fmt0(dfm["n"].sum()), center_sub="pessoas")
                    st.plotly_chart(fig, use_container_width=True, key="donut_wf_tipo")
                else:
                    st.caption("Sem dados de efetivo por natureza.")
        with w2:
            with panel("Efetivo por Coordenador", "Headcount ativo alocado por coordenador responsável"):
                if not data_wf.por_coordenador.empty:
                    dfc = data_wf.por_coordenador.sort_values("n", ascending=True)
                    fig = go.Figure(go.Bar(x=dfc["n"], y=dfc["coordenador"], orientation="h", marker_color=BLUE))
                    base_layout(fig, height=270, legend=False)
                    st.plotly_chart(fig, use_container_width=True, key="bar_wf_coord")
                else:
                    st.caption("Sem dados por coordenador.")
        with w3:
            with panel("Efetivo por Tipo de Contrato", "Headcount ativo — Fixa × Variável × Spot"):
                if not data_wf.por_contrato.empty:
                    dfk = data_wf.por_contrato
                    colors_k = {"Fixa": GREEN, "Variável": AMBER, "Spot": RED}
                    fig = donut(dfk["tipo"].tolist(), dfk["n"].tolist(),
                                [colors_k.get(t, BASE_GREY) for t in dfk["tipo"]],
                                center_title=fmt0(dfk["n"].sum()), center_sub="pessoas")
                    st.plotly_chart(fig, use_container_width=True, key="donut_wf_contrato")
                else:
                    st.caption("Sem dados de efetivo por tipo de contrato.")

        w4, w5 = st.columns(2)
        with w4:
            with panel("Admissões × Saídas por Mês", "Movimentação de pessoal — histórico do Turnover"):
                if not data_wf.turnover.empty:
                    tv = data_wf.turnover.copy()
                    tv["saidas"] = (tv["demitidos"].fillna(0) + tv["pedido_demissao"].fillna(0))
                    fig = go.Figure()
                    fig.add_bar(x=tv["label"], y=tv["admitidos"], name="Admitidos", marker_color=GREEN)
                    fig.add_bar(x=tv["label"], y=tv["saidas"], name="Saídas", marker_color=RED)
                    fig.update_layout(barmode="group")
                    base_layout(fig, height=300)
                    st.plotly_chart(fig, use_container_width=True, key="bar_wf_turnover")
                else:
                    st.caption("Sem histórico de turnover.")
        with w5:
            with panel("Efetivo por Cliente", 'Headcount ativo alocado por cliente (demais agrupados em "Outros")'):
                if not data_wf.por_cliente.empty:
                    dfcl = data_wf.por_cliente.sort_values("n", ascending=True)
                    fig = go.Figure(go.Bar(x=dfcl["n"], y=dfcl["cliente"], orientation="h", marker_color=NAVY))
                    base_layout(fig, height=300, legend=False)
                    st.plotly_chart(fig, use_container_width=True, key="bar_wf_cliente")
                else:
                    st.caption("Sem dados de efetivo por cliente.")

        section_title("Certificações — Risco de Vencimento", "Vencidas + próximos 90 dias")
        with panel("", ""):
            crit = data_wf.certs_criticas
            if not crit.empty:
                show = crit.copy()
                show["Situação"] = show["dias_restantes"].map(
                    lambda d: "🔴 Vencida" if d < 0 else ("🟠 Vence em breve" if d <= 30 else "🟡 Vencendo"))
                show["Validade"] = show["validade_dt"].dt.strftime("%d/%m/%Y")
                show = show.rename(columns={"nome": "Nome", "qualificacao": "Qualificação", "dias_restantes": "Dias"})
                html_table(show[["Nome", "Qualificação", "Validade", "Dias", "Situação"]].head(30),
                           right_cols=["Dias"])
            else:
                st.caption("Nenhuma certificação vencida ou vencendo nos próximos 90 dias. ✅")
    else:
        section_title("Recursos &amp; Pessoas", "")
        st.info("Aponte, na barra lateral, para o arquivo **Planejamento_Ativos_Status_*.xlsx** "
                "(base de RH) para habilitar os indicadores de efetivo, turnover e certificações.")


with tab3:
    # --------------------------------------------------------------------------- #
    # Escalas & Embarques — Dobras, AFI (folga) e Ociosidade/STB
    # --------------------------------------------------------------------------- #

    if data_emb is not None:
        latest = data_emb.latest
        section_title("Escalas &amp; Embarques", f"Ref.: {data_emb.mes_ref_label}")
        st.caption("Base de eventos de escala — dobras (embarque extra), AFI (folga indenizada) e "
                   "ociosidade/STB (aguardando embarque). Não é afetada pelos filtros da barra lateral.")

        if latest:
            embarcado = (latest["regulares"] or 0) + (latest["dobras"] or 0)
            e1, e2, e3 = st.columns(3)
            kpi_card(e1, "Efetivo no Mês", fmt0(latest["efetivo"]), " pessoas",
                     f"{data_emb.mes_ref_label} · base de referência da escala", accent=NAVY)
            kpi_card(e2, "Embarcado (Regular + Dobra)", fmt0(embarcado), " diárias",
                     f"{fmt1p(latest['pct_reg_cap'])} de utilização vs. capacidade", accent=BLUE,
                     bar_pct=latest["pct_reg_cap"])
            kpi_card(e3, "Ociosidade / Standby", fmt0(latest["ociosidade"]), " diárias",
                     f"{fmt1p(latest['pct_ociosidade'])} do total do mês — menor é melhor", accent=AMBER,
                     bar_pct=latest["pct_ociosidade"])

            e4, e5, e6 = st.columns(3)
            kpi_card(e4, "Dobras (Embarque Extra)", fmt0(latest["dobras"]), " diárias",
                     f"{fmt1p(latest['pct_dobras_hd'])} das diárias-homem do mês", accent=RED,
                     bar_pct=min(1.0, latest["pct_dobras_hd"] * 5))
            kpi_card(e5, "AFI (Folga Indenizada)", fmt0(latest["afi"]), " diárias",
                     f"{fmt1p(latest['pct_afi_hd'])} das diárias-homem do mês", accent=GREEN,
                     bar_pct=min(1.0, latest["pct_afi_hd"] * 5))
            kpi_card(e6, "Funcionários com STB", fmt0(int(data_emb.pct_com_stb * data_emb.total_funcionarios)),
                     " pessoas", f"{fmt1p(data_emb.pct_com_stb)} da base no período", accent=AMBER,
                     bar_pct=data_emb.pct_com_stb)

        m1, m2, m3 = st.columns(3)
        with m1:
            with panel("Evolução Mensal — Diárias × Ociosidade", "Total embarcado (regular + dobra) vs. standby, mês a mês"):
                if not data_emb.monthly.empty:
                    mo = data_emb.monthly
                    fig = go.Figure()
                    fig.add_scatter(x=mo["mes"], y=mo["total_diarias"], name="Diárias (total)",
                                     mode="lines+markers", line=dict(color=BLUE, width=3))
                    fig.add_scatter(x=mo["mes"], y=mo["ociosidade"], name="Ociosidade/STB",
                                     mode="lines+markers", line=dict(color=AMBER, width=3))
                    base_layout(fig, height=280)
                    st.plotly_chart(fig, use_container_width=True, key="line_emb_mensal")
                else:
                    st.caption("Sem série mensal disponível.")
        with m2:
            with panel("Dobras × AFI por Mês", "Volume de embarque extra vs. folga indenizada, mês a mês"):
                if not data_emb.monthly.empty:
                    mo = data_emb.monthly
                    fig = go.Figure()
                    fig.add_bar(x=mo["mes"], y=mo["dobras"], name="Dobras", marker_color=RED)
                    fig.add_bar(x=mo["mes"], y=mo["afi"], name="AFI", marker_color=GREEN)
                    fig.update_layout(barmode="group")
                    base_layout(fig, height=280)
                    st.plotly_chart(fig, use_container_width=True, key="bar_emb_dobras_afi")
                else:
                    st.caption("Sem série mensal disponível.")
        with m3:
            with panel("Composição de Eventos", f"Mix do mês de referência ({data_emb.mes_ref_label})"):
                if latest:
                    labels = ["Regulares", "Dobras", "Ociosidade/STB", "AFI"]
                    values = [latest["regulares"], latest["dobras"], latest["ociosidade"], latest["afi"]]
                    fig = donut(labels, values, [BLUE, RED, AMBER, GREEN],
                                center_title=fmt0(sum(values)), center_sub="diárias")
                    st.plotly_chart(fig, use_container_width=True, key="donut_emb_mix")
                else:
                    st.caption("Sem dados do mês de referência.")

        section_title("Concentração por Família de Função", "Dobras + AFI + Ociosidade acumulados no período")
        with panel("", ""):
            if not data_emb.por_familia.empty:
                pf = data_emb.por_familia.sort_values("total", ascending=True)
                fig = go.Figure()
                fig.add_bar(y=pf["familia"], x=pf["dobras"], name="Dobras", orientation="h", marker_color=RED)
                fig.add_bar(y=pf["familia"], x=pf["afi"], name="AFI", orientation="h", marker_color=GREEN)
                fig.add_bar(y=pf["familia"], x=pf["ociosidade"], name="Ociosidade/STB", orientation="h", marker_color=AMBER)
                fig.update_layout(barmode="stack")
                base_layout(fig, height=320)
                st.plotly_chart(fig, use_container_width=True, key="bar_emb_familia")
            else:
                st.caption("Sem dados de composição por família de função.")

        t1, t2 = st.columns(2)
        with t1:
            with panel("Ranking — Maior Ociosidade/STB", "Top 15 funcionários, acumulado no período"):
                top_o = data_emb.top_ociosidade
                if not top_o.empty:
                    show = top_o.rename(columns={"nome": "Nome", "funcao": "Função", "ociosidade": "STB (dias)",
                                                   "pct_ociosidade": "% do HD"})
                    html_table(show[["Nome", "Função", "STB (dias)", "% do HD"]], right_cols=["STB (dias)", "% do HD"])
                else:
                    st.caption("Sem funcionários com ociosidade registrada.")
        with t2:
            with panel("Ranking — Mais Dobras", "Top 15 funcionários com mais embarque extra, acumulado no período"):
                top_d = data_emb.top_dobras
                if not top_d.empty:
                    show = top_d.rename(columns={"nome": "Nome", "funcao": "Função", "dobras": "Dobras",
                                                   "pct_dobras_hd": "% do HD"})
                    html_table(show[["Nome", "Função", "Dobras", "% do HD"]], right_cols=["Dobras", "% do HD"])
                else:
                    st.caption("Sem funcionários com dobra registrada.")

        t3, t4 = st.columns(2)
        with t3:
            with panel("Ranking — Mais AFI (Folga Indenizada)", "Top 15 funcionários, acumulado no período"):
                top_a = data_emb.top_afi
                if not top_a.empty:
                    show = top_a.rename(columns={"nome": "Nome", "funcao": "Função", "afi": "AFI",
                                                   "pct_afi_hd": "% do HD"})
                    html_table(show[["Nome", "Função", "AFI", "% do HD"]], right_cols=["AFI", "% do HD"])
                else:
                    st.caption("Sem funcionários com AFI registrada.")
        with t4:
            with panel("Observações da Liderança", "Justificativas registradas para dobras/AFI fora do padrão"):
                obs = data_emb.observacoes
                if not obs.empty:
                    show = obs.rename(columns={"nome": "Nome", "evento": "Evento", "comentario": "Comentário"})
                    html_table(show[["Nome", "Evento", "Comentário"]].head(12))
                else:
                    st.caption("Sem observações registradas nesta base.")

        section_title("Base Completa — Escala por Funcionário", "Dobras, AFI e Ociosidade/STB acumulados no período")
        with panel("", ""):
            if not data_emb.workers.empty:
                show = data_emb.workers.sort_values("ociosidade", ascending=False).head(50).rename(columns={
                    "nome": "Nome", "funcao": "Função", "familia": "Família",
                    "dobras": "Dobras", "afi": "AFI", "ociosidade": "Ociosidade/STB",
                })
                html_table(show[["Nome", "Função", "Família", "Dobras", "AFI", "Ociosidade/STB"]],
                           right_cols=["Dobras", "AFI", "Ociosidade/STB"])
                st.caption(f"Mostrando 50 de {data_emb.total_funcionarios} funcionários da base, "
                           f"ordenados por Ociosidade/STB.")
            else:
                st.caption("Sem base de funcionários disponível.")
    else:
        section_title("Escalas &amp; Embarques", "")
        st.info("Aponte, na barra lateral, para o arquivo **Relatório Mensal de Eventos*.xlsx** "
                "(base de escalas) para habilitar os indicadores de dobras, AFI/folga e ociosidade/STB.")


# --------------------------------------------------------------------------- #
# Footer
# --------------------------------------------------------------------------- #

st.markdown(f"""
<div class="qt-disclaimer">
  <b>Nota:</b> dashboard consolidado a partir das tabelas dinâmicas do arquivo <b>{os.path.basename(data.path)}</b>
  ({data.periodo_label}). Identidade visual alinhada à paleta oficial de marca da Qualitech IRM.
  Valores em man-days (diárias), salvo indicação em contrário.
</div>
<div class="qt-footer-note">
  Qualitech IRM · Dashboard gerado automaticamente a partir do Performance_*.xlsx —
  sem necessidade de regenerar manualmente a cada atualização.
</div>
""", unsafe_allow_html=True)
