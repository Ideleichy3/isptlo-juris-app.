"""
ILR Academic Solutions — Plataforma de Gestão de Júris Académicos
Versão 4.0 | Sprint 2 | Comercial & Escalável
Autora: Ph.D. Ideleichy Lombillo Rivero

Instalação:
    pip install streamlit pandas openpyxl plotly fpdf2 requests

Execução:
    streamlit run isptlo_juris_app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import datetime
import io
import json
import base64
import requests
from pdf_block import render_docente_pdf_block, render_tesouraria_pdf_block

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Plataforma de Gestão de Júris Académicos",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

IRT_RATE   = 0.065
TODAY      = datetime.date.today().strftime("%d/%m/%Y")
SHEET_ID   = "1zo5dwPSMZRhh7YeW7GGT9KWtZgNUQGj9MgtTQ3z_Sq4"
APP_VER    = "4.0"

# ── LOGO (embutido) ───────────────────────────────────────────────────────────
try:
    import os
    _logo_path = os.path.join(os.path.dirname(__file__), "Image_1v_APP_juris_pagamento.png")
    if os.path.isfile(_logo_path):
        with open(_logo_path, "rb") as _f:
            LOGO_SRC = "data:image/png;base64," + base64.b64encode(_f.read()).decode()
    else:
        LOGO_SRC = ""
except Exception:
    LOGO_SRC = ""

# ── SUBSÍDIOS (fallback — sobrescritos se GSheets disponível) ────────────────
SUBSIDIOS_DEFAULT = {
    ("Presidente de Júri",        "Individual"): 20_000,
    ("Presidente de Júri",        "Dupla"):      27_500,
    ("1º Vogal Arguente",         "Individual"): 20_000,
    ("1º Vogal Arguente",         "Dupla"):      27_500,
    ("2º Vogal Tutor/Orientador", "Individual"): 80_000,
    ("2º Vogal Tutor/Orientador", "Dupla"):     104_000,
    ("Co-tutor",                  "Individual"): 40_000,
    ("Co-tutor",                  "Dupla"):      72_000,
    ("Secretário",                "Individual"): 10_000,
    ("Secretário",                "Dupla"):      15_000,
    # Estágios
    ("Presidente de Júri",        "Estágio"):    15_000,
    ("Oponente/Arguente",         "Estágio"):    15_000,
    ("Tutor de Estágio",          "Estágio"):    60_000,
    ("Secretário",                "Estágio"):     8_000,
}

MEMBERS_DEFAULT = [
    {"Nome": "Manuel Octávio Isaac Spínola",  "Grau": "PhD",    "Cargo": "Presidente do ISPTLO",
     "Departamento": "Agronomia / Engenharias e Tecnologias"},
    {"Nome": "Ideleichy Lombillo Rivero",     "Grau": "PhD",    "Cargo": "Vice-Presidente Científica",
     "Departamento": "Todos os Cursos"},
    {"Nome": "Elizabeth González",            "Grau": "MSc",    "Cargo": "Vice-Presidente Académica",
     "Departamento": "Todos os Cursos"},
    {"Nome": "Domingos Lunga",                "Grau": "MSc",    "Cargo": "Vice-Presidente Admin.",
     "Departamento": "Todos os Cursos"},
    {"Nome": "José Fernando Manuel",          "Grau": "MSc",    "Cargo": "Docente",
     "Departamento": "Contabilidade e Gestão / CSH"},
    {"Nome": "Walquiria Chissimo",            "Grau": "MSc",    "Cargo": "Docente",
     "Departamento": "Ensino Primário / CSH"},
    {"Nome": "Félix Palau",                   "Grau": "MSc",    "Cargo": "Docente",
     "Departamento": "Agronomia / Engenharias e Tecnologias"},
    {"Nome": "Marilda Augusto",               "Grau": "Lic",    "Cargo": "Docente",
     "Departamento": "Enfermagem / Ciências da Saúde"},
    {"Nome": "José Monteiro",                 "Grau": "Lic",    "Cargo": "Docente",
     "Departamento": "Contabilidade e Gestão / CSH"},
    {"Nome": "Domingos Ngando",               "Grau": "Lic",    "Cargo": "Docente",
     "Departamento": "Ensino Primário / CSH"},
]

CURSOS_DEFAULT = [
    "Contabilidade e Gestão / CSH",
    "Ensino Primário / CSH",
    "Agronomia / Engenharias e Tecnologias",
    "Enfermagem / Ciências da Saúde",
    "Todos os Cursos",
]

FUNCOES_TFC = [
    "Presidente de Júri",
    "1º Vogal Arguente",
    "2º Vogal Tutor/Orientador",
    "Co-tutor",
    "Secretário",
]
FUNCOES_ESTAGIO = [
    "Presidente de Júri",
    "Oponente/Arguente",
    "Tutor de Estágio",
    "Secretário",
]
MESES_REF = [
    "Jan/2026","Fev/2026","Mar/2026","Abr/2026","Mai/2026","Jun/2026",
    "Jul/2026","Ago/2026","Set/2026","Out/2026","Nov/2026","Dez/2026",
]
CANAIS    = ["Transferência Bancária","Numerário","Cheque","Multicaixa"]
MESES_PT  = {"Jan":"Jan","Feb":"Fev","Mar":"Mar","Apr":"Abr","May":"Mai",
             "Jun":"Jun","Jul":"Jul","Aug":"Ago","Sep":"Set","Oct":"Out",
             "Nov":"Nov","Dec":"Dez"}

# ── GOOGLE SHEETS — CARREGAMENTO DINÂMICO ─────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_sheet(sheet_name: str) -> pd.DataFrame | None:
    """
    Carrega uma aba do Google Sheet via CSV público.
    Retorna None se não acessível — app usa defaults.
    NOTA: O Google Sheet deve estar partilhado como 'Qualquer pessoa com o link pode ver'.
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    )
    try:
        resp = requests.get(url, timeout=8)
        if resp.status_code == 200 and len(resp.content) > 50:
            df = pd.read_csv(io.StringIO(resp.text))
            df = df.dropna(how="all")
            return df
    except Exception:
        pass
    return None


@st.cache_data(ttl=300, show_spinner=False)
def load_members() -> list[dict]:
    """Carrega membros da aba MEMBROS_JÚRI. Fallback para lista interna."""
    df = load_sheet("MEMBROS_JURI")
    if df is not None and not df.empty:
        cols_needed = ["Nome", "Grau", "Cargo", "Departamento"]
        available   = [c for c in cols_needed if c in df.columns]
        if "Nome" in available and "Grau" in available:
            return df[available].fillna("").to_dict("records")
    return MEMBERS_DEFAULT


@st.cache_data(ttl=300, show_spinner=False)
def load_subsidios() -> dict:
    """Carrega tabela de subsídios. Fallback para SUBSIDIOS_DEFAULT."""
    df = load_sheet("TABELA_SUBSIDIOS")
    if df is not None and "Funcao" in df.columns and "Tipo" in df.columns and "Valor" in df.columns:
        result = {}
        for _, row in df.iterrows():
            try:
                result[(str(row["Funcao"]).strip(), str(row["Tipo"]).strip())] = float(row["Valor"])
            except (ValueError, TypeError):
                pass
        if result:
            return result
    return SUBSIDIOS_DEFAULT


# ── FUNÇÕES DE HIERARQUIA (baseadas em REGRAS, não em nomes fixos) ────────────
def get_nivel(grau: str) -> int:
    """PhD=3, MSc=2, Lic=1, Externo=0"""
    g = grau.upper().strip()
    if "PHD" in g or "DOUTORAD" in g or "DR." in g:
        return 3
    if "MSC" in g or "MESTRE" in g or "MESTRAD" in g:
        return 2
    if "LIC" in g or "LICENCIAD" in g:
        return 1
    return 0


def is_vice_president(cargo: str) -> bool:
    return "vice" in cargo.lower() or "vp" in cargo.lower()


def pode_presidir_base(grau: str) -> bool:
    return get_nivel(grau) >= 3  # só PhD por defeito


def pode_presidir_estagio_base(grau: str) -> bool:
    return get_nivel(grau) >= 2  # MSc+ podem presidir estágios


def pode_arguir(grau: str) -> bool:
    return get_nivel(grau) >= 2  # MSc+ por defeito, Lic com excepção


def validar_funcao(grau: str, cargo: str, funcao: str, modalidade: str) -> dict:
    """
    Valida se o docente pode exercer a função pedida.
    Retorna: {ok: bool, aviso: str|None, requer_confirmacao: bool, tipo_excecao: str|None}
    """
    nivel = get_nivel(grau)
    vp    = is_vice_president(cargo)

    if modalidade in ("Defesa TFC", "Pré-Defesa TFC"):
        if funcao == "Presidente de Júri":
            if nivel >= 3:
                return {"ok": True, "aviso": None, "requer_confirmacao": False, "tipo": None}
            if nivel == 2:
                if vp:
                    return {"ok": True,
                            "aviso": "Vice-Presidente: excepção hierárquica aprovada automaticamente.",
                            "requer_confirmacao": False, "tipo": "Excepção: VP"}
                return {"ok": False,
                        "aviso": "Este docente (MSc) não tem grau de PhD. "
                                 "Exerce cargo de Vice-Presidente ou possui Notório Saber?",
                        "requer_confirmacao": True, "tipo": "Excepção: Notório Saber/VP"}
            return {"ok": False,
                    "aviso": "Licenciados não podem presidir júris de TFC.",
                    "requer_confirmacao": False, "tipo": None}

        if funcao in ("1º Vogal Arguente", "Oponente/Arguente"):
            if nivel >= 2:
                return {"ok": True, "aviso": None, "requer_confirmacao": False, "tipo": None}
            return {"ok": False,
                    "aviso": "Este docente é Licenciado. "
                             "Possui Notório Saber ou é Especialista reconhecido?",
                    "requer_confirmacao": True, "tipo": "Excepção: Notório Saber"}

        # Tutor, Co-tutor, Secretário — sem restrição rígida de grau
        return {"ok": True, "aviso": None, "requer_confirmacao": False, "tipo": None}

    # ── Modalidade Estágio ────────────────────────────────────────────────────
    if modalidade == "Relatório de Estágio":
        if funcao == "Presidente de Júri":
            if nivel >= 3:
                return {"ok": True, "aviso": None, "requer_confirmacao": False, "tipo": None}
            if nivel == 2:
                return {"ok": False,
                        "aviso": "Em Estágios, Mestres podem presidir com Notório Saber "
                                 "ou cargo de chefia. Confirmar?",
                        "requer_confirmacao": True, "tipo": "Excepção: Notório Saber (Estágio)"}
            # Lic — excepção especialista
            return {"ok": False,
                    "aviso": "Licenciados só presidem estágios se forem Especialistas reconhecidos. Confirmar?",
                    "requer_confirmacao": True, "tipo": "Excepção: Especialista (Estágio)"}
        if funcao in ("Oponente/Arguente", "Tutor de Estágio"):
            # Em Estágios, Lic especialistas podem ser oponentes/tutores
            if nivel >= 2:
                return {"ok": True, "aviso": None, "requer_confirmacao": False, "tipo": None}
            return {"ok": False,
                    "aviso": "Licenciado como Oponente/Tutor em Estágio: "
                             "confirmar que é Especialista na área?",
                    "requer_confirmacao": True, "tipo": "Excepção: Especialista"}
        return {"ok": True, "aviso": None, "requer_confirmacao": False, "tipo": None}

    return {"ok": True, "aviso": None, "requer_confirmacao": False, "tipo": None}


# ── UTILITÁRIOS ───────────────────────────────────────────────────────────────
def kz(val):
    if val is None or val == 0:
        return "— Kz"
    return f"{val:,.2f} Kz".replace(",","X").replace(".",",").replace("X",".")


def get_bruto(funcao: str, tipo: str, subsidios: dict) -> float:
    return subsidios.get((funcao, tipo), 0.0)


def calcular_resumo_irt(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    grp = (df.groupby(["docente","mes_ref"])["bruto"]
           .sum().reset_index()
           .rename(columns={"bruto":"total_bruto"}))
    grp["irt_6_5"]  = (grp["total_bruto"] * IRT_RATE).round(2)
    grp["liquido"]  = grp["total_bruto"] - grp["irt_6_5"]
    return grp


def get_lancamentos_df() -> pd.DataFrame:
    if not st.session_state.lancamentos:
        return pd.DataFrame(columns=[
            "despacho","docente","grau","cargo","funcao","tipo","modalidade",
            "curso","data","mes_ref","bruto","excecao",
            "estudante","titulo_trabalho","avaliacao",
        ])
    return pd.DataFrame(st.session_state.lancamentos)


def get_resumo() -> pd.DataFrame:
    return calcular_resumo_irt(get_lancamentos_df())


def mes_from_date(d: datetime.date) -> str:
    return MESES_PT.get(d.strftime("%b"), d.strftime("%b")) + "/" + str(d.year)


def check_duplicate(despacho: str, docente: str) -> bool:
    df = get_lancamentos_df()
    if df.empty:
        return False
    return not df[(df["despacho"] == despacho) & (df["docente"] == docente)].empty


# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background: #0D1B3E; }
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stRadio label { color: white !important; }
[data-testid="stSidebar"] .stSelectbox label { color: #FFD700 !important; font-weight:bold; }
[data-testid="stSidebar"] .stButton button {
    background:#1F3864 !important; color:#FFD700 !important;
    border:1px solid #FFD700 !important; border-radius:6px !important;
    font-size:12px !important; padding:4px 8px !important;
}
.kpi-card {
    background:#1F3864; border-radius:12px; padding:1rem 1.2rem;
    color:white; text-align:center; margin-bottom:0.5rem;
}
.kpi-card .kpi-label { font-size:11px; opacity:0.85; margin-bottom:4px; }
.kpi-card .kpi-value { font-size:24px; font-weight:700; line-height:1.1; }
.kpi-card.red   { background:#C00000; }
.kpi-card.green { background:#1E7145; }
.kpi-card.amber { background:#854F0B; }
.kpi-card.teal  { background:#005A5B; }
.alert-legal {
    background:#FFF2CC; border-left:5px solid #854F0B;
    padding:.6rem 1rem; border-radius:6px; font-size:13px; color:#5C3D00; margin:.4rem 0;
}
.alert-error {
    background:#FFCCCC; border-left:5px solid #C00000;
    padding:.6rem 1rem; border-radius:6px; font-size:13px; color:#5C0000; margin:.4rem 0;
}
.alert-ok {
    background:#C6EFCE; border-left:5px solid #1E7145;
    padding:.6rem 1rem; border-radius:6px; font-size:13px; color:#0A3D1F; margin:.4rem 0;
}
.alert-info {
    background:#DDEEFF; border-left:5px solid #1F3864;
    padding:.6rem 1rem; border-radius:6px; font-size:13px; color:#1F3864; margin:.4rem 0;
}
.alert-excecao {
    background:#EAD1DC; border-left:5px solid #4A235A;
    padding:.6rem 1rem; border-radius:6px; font-size:13px; color:#4A235A; margin:.4rem 0;
}
.section-title {
    font-size:16px; font-weight:700; color:#1F3864;
    border-bottom:3px solid #1F3864; padding-bottom:4px; margin:1.2rem 0 0.8rem 0;
}
.main-header {
    background:linear-gradient(135deg,#0D1B3E 0%,#1F3864 100%);
    color:white; padding:1rem 1.4rem; border-radius:12px; margin-bottom:1rem;
    display:flex; align-items:center; gap:1rem;
    border-bottom:3px solid #FFD700;
}
.main-header img { width:62px; height:62px; border-radius:50%; border:2px solid #FFD700; flex-shrink:0; }
.main-header h1  { color:#FFD700; margin:0; font-size:20px; }
.main-header p   { color:#CCDDFF; margin:3px 0 0 0; font-size:12px; }
.step-indicator { display:flex; justify-content:center; gap:5px; margin:.3rem 0 .6rem 0; }
.step-dot { width:8px; height:8px; border-radius:50%; background:#3A5A8A; display:inline-block; }
.step-dot.active { background:#FFD700; }
.step-dot.done   { background:#1E7145; }
.ilr-brand {
    background:linear-gradient(135deg,#060E1F,#1F3864);
    border-top:2px solid #FFD700; border-radius:0 0 8px 8px;
    padding:.7rem .6rem; text-align:center;
}
.ilr-brand .ilr-name   { color:#FFD700; font-size:11px; font-weight:700; line-height:1.5; }
.ilr-brand .ilr-slogan { color:#AABBDD; font-size:9px; font-style:italic; }
.ilr-brand .ilr-ver    { color:#7799CC; font-size:9px; margin-top:2px; }
.badge-excecao {
    display:inline-block; background:#4A235A; color:white;
    font-size:9px; font-weight:700; padding:1px 6px; border-radius:8px;
}
.badge-admin {
    display:inline-block; background:#C00000; color:white;
    font-size:9px; font-weight:700; padding:1px 6px; border-radius:8px;
}
.stDownloadButton button {
    background:#1F3864 !important; color:#FFD700 !important;
    border:2px solid #FFD700 !important; border-radius:8px !important; font-weight:600 !important;
}
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────────────────────
_state_defaults = {
    "lancamentos":  [],
    "pagamentos":   {},
    "modo":         "Tesouraria",
    "pagina_idx":   0,
    "admin_mode":   False,
    "gs_status":    None,          # "ok" | "fallback" | None
    "edit_idx":     None,
    "confirmacao_pendente": None,  # dict com dados pendentes de excepção
}
for k, v in _state_defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── CARREGAR DADOS DINÂMICOS ──────────────────────────────────────────────────
MEMBERS  = load_members()
SUBSIDIOS = load_subsidios()
CURSOS   = sorted(set(m.get("Departamento","") for m in MEMBERS if m.get("Departamento","")))
if not CURSOS:
    CURSOS = CURSOS_DEFAULT

MEMBER_NAMES = [m["Nome"] for m in MEMBERS]

def get_member(nome: str) -> dict:
    for m in MEMBERS:
        if m["Nome"] == nome:
            return m
    return {"Nome": nome, "Grau": "MSc", "Cargo": "Docente", "Departamento": ""}

GS_OK = load_sheet("MEMBROS_JURI") is not None

# ── PÁGINA MENU (controlado APENAS por índice — fix dos botões) ───────────────
PAGINAS_TES = [
    "➕ 1. Registar Lançamento",
    "📊 2. Dashboard Executivo",
    "🧾 3. Resumo IRT por Docente",
    "💳 4. Registar Pagamento",
    "📋 5. Todos os Lançamentos",
    "📥 6. Exportar Relatórios",
]

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    # Logo + nome
    logo_html = (f'<img src="{LOGO_SRC}" style="width:52px;height:52px;'
                 f'border-radius:50%;border:2px solid #FFD700;margin-right:10px"/>'
                 if LOGO_SRC else "")
    st.markdown(
        f'<div style="display:flex;align-items:center;padding:.4rem 0">'
        f'{logo_html}'
        f'<div><div style="color:#FFD700;font-weight:700;font-size:14px">'
        f'ISPTLO-JURIS</div>'
        f'<div style="color:#AABBDD;font-size:10px">Gestão de Júris v{APP_VER}</div></div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div style="font-size:10px;color:{"#1E7145" if GS_OK else "#854F0B"};'
        f'margin-bottom:.3rem">'
        f'{"🟢 Google Sheets ligado" if GS_OK else "🟡 Modo offline (dados locais)"}'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")

    modo = st.radio("Modo", ["👤 Docente","🏦 Tesouraria"],
                    index=0 if st.session_state.modo=="Docente" else 1)
    st.session_state.modo = "Docente" if "Docente" in modo else "Tesouraria"
    st.markdown("---")

    if st.session_state.modo == "Docente":
        # Docente — menu simples
        pagina = "💰 Calcular os meus haveres"
        st.markdown("**💰 Calculadora de Haveres**")
    else:
        # ── NAVEGAÇÃO CORRIGIDA ─────────────────────────────────────────────
        # O selectbox segue o índice — NÃO define o índice.
        # Os botões modificam o índice e fazem rerun — selectbox acompanha.

        n = len(PAGINAS_TES)
        idx = st.session_state.pagina_idx

        # Sincronizar selectbox (apenas visualização + mudança manual)
        def _on_pagina_change():
            sel = st.session_state["_sel_pag"]
            st.session_state.pagina_idx = PAGINAS_TES.index(sel)

        st.selectbox(
            "Secção",
            PAGINAS_TES,
            index=idx,
            key="_sel_pag",
            on_change=_on_pagina_change,
        )

        # Step dots
        dots = "".join(
            f'<span class="step-dot {"active" if i==idx else ("done" if i<idx else "")}">'
            f'</span>'
            for i in range(n)
        )
        st.markdown(f'<div class="step-indicator">{dots}</div>', unsafe_allow_html=True)

        # Nav buttons — modificam o índice ANTES do selectbox renderizar
        c1, c2 = st.columns(2)
        with c1:
            if st.button("◀ Ant.", disabled=(idx==0), use_container_width=True, key="btn_prev"):
                st.session_state.pagina_idx = max(0, idx-1)
                st.rerun()
        with c2:
            if st.button("Próx. ▶", disabled=(idx==n-1), use_container_width=True, key="btn_next"):
                st.session_state.pagina_idx = min(n-1, idx+1)
                st.rerun()
        if st.button("🏠 Início", use_container_width=True, key="btn_home"):
            st.session_state.pagina_idx = 0
            st.rerun()

        pagina = PAGINAS_TES[st.session_state.pagina_idx]

    st.markdown("---")

    # Admin
    with st.expander("🔐 Acesso Admin", expanded=False):
        pwd = st.text_input("", type="password", key="admin_pwd",
                            placeholder="Password administrador",
                            label_visibility="collapsed")
        if st.button("Entrar", key="btn_admin"):
            st.session_state.admin_mode = (pwd == "ILR2026@ISPTLO")
            st.rerun()
        if st.session_state.admin_mode:
            st.success("Admin activo")
            if st.button("Sair", key="btn_logout_admin"):
                st.session_state.admin_mode = False
                st.rerun()

    st.markdown("---")
    st.markdown(
        "<small style='color:#5577AA'>Base legal: D.P. nº 191/18<br>"
        "IRT: C.I.R.T. Lei nº 19/14 Art.º 67<br>"
        "Ofício 009/25 + Emenda DASG 023/2025</small>",
        unsafe_allow_html=True,
    )

    # ILR Brand
    st.markdown(
        '<div class="ilr-brand">'
        '<div class="ilr-name">Ph.D. Ideleichy Lombillo Rivero<br>'
        'ILR — Academic Solutions</div>'
        '<div class="ilr-slogan">Gestão Inteligente para Instituições de Excelência</div>'
        f'<div class="ilr-ver">Versão {APP_VER} | ISPTLO — TFC/Pagamentos</div>'
        '</div>',
        unsafe_allow_html=True,
    )

# ── HEADER PRINCIPAL ──────────────────────────────────────────────────────────
logo_img = f'<img src="{LOGO_SRC}"/>' if LOGO_SRC else "🎓"
st.markdown(
    f'<div class="main-header">'
    f'{logo_img}'
    f'<div>'
    f'<h1>Plataforma de Gestão de Júris Académicos</h1>'
    f'<p>ILR Academic Solutions | Ph.D. Ideleichy Lombillo Rivero | '
    f'v{APP_VER} | IRT calculado sobre total mensal por docente</p>'
    f'</div></div>',
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════════
# MODO DOCENTE
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.modo == "Docente":
    st.markdown('<div class="section-title">💰 Calculadora de Haveres — Modo Docente</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<div class="alert-legal">ℹ O IRT de 6,5% incide sobre o <strong>total mensal '
        'acumulado</strong> — não por sessão. Introduza todas as participações do mês.</div>',
        unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        docente_sel = st.selectbox("O meu nome", MEMBER_NAMES)
        mes_sel     = st.selectbox("Mês de referência", MESES_REF, index=3)
    with c2:
        mem      = get_member(docente_sel)
        grau_d   = mem.get("Grau","MSc")
        cargo_d  = mem.get("Cargo","Docente")
        nivel_d  = get_nivel(grau_d)
        st.info(f"**Nível:** {grau_d} | **Cargo:** {cargo_d}")

    modalidade_d = st.selectbox("Modalidade", ["Defesa TFC","Pré-Defesa TFC","Relatório de Estágio"])
    funcoes_base = FUNCOES_ESTAGIO if modalidade_d=="Relatório de Estágio" else FUNCOES_TFC

    n_part  = st.number_input("Número de participações no mês", 1, 20, 1)
    total_b = 0
    parts   = []
    for i in range(int(n_part)):
        with st.expander(f"Participação {i+1}", expanded=True):
            c1, c2 = st.columns(2)
            with c1:
                func_i = st.selectbox("Função", funcoes_base, key=f"dfunc_{i}")
            with c2:
                tipo_i_opts = ["Estágio"] if modalidade_d=="Relatório de Estágio" else ["Individual","Dupla"]
                tipo_i = st.selectbox("Tipo", tipo_i_opts, key=f"dtype_{i}")
            bruto_i = get_bruto(func_i, tipo_i, SUBSIDIOS)
            st.markdown(f"**Valor bruto:** `{kz(bruto_i)}`")
            total_b += bruto_i
            parts.append({"funcao":func_i,"tipo":tipo_i,"bruto":bruto_i})

    st.markdown("---")
    irt     = round(total_b * IRT_RATE, 2)
    liquido = total_b - irt

    c1,c2,c3 = st.columns(3)
    c1.markdown(f'<div class="kpi-card"><div class="kpi-label">TOTAL BRUTO MENSAL</div>'
                f'<div class="kpi-value">{kz(total_b)}</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="kpi-card red"><div class="kpi-label">IRT 6,5% → AGT</div>'
                f'<div class="kpi-value">{kz(irt)}</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="kpi-card green"><div class="kpi-label">LÍQUIDO A RECEBER</div>'
                f'<div class="kpi-value">{kz(liquido)}</div></div>', unsafe_allow_html=True)

    if parts:
        st.dataframe(pd.DataFrame(parts).assign(bruto_fmt=lambda d: d["bruto"].apply(kz))
                     [["funcao","tipo","bruto_fmt"]].rename(columns={"funcao":"Função","tipo":"Tipo","bruto_fmt":"Valor Bruto"}),
                     use_container_width=True, hide_index=True)

    render_docente_pdf_block(
        docente_name=docente_sel, mes_ref=mes_sel, participacoes=parts,
        total_bruto=total_b, total_irt=irt, total_liquido=liquido)

# ══════════════════════════════════════════════════════════════════════════════
# MODO TESOURARIA
# ══════════════════════════════════════════════════════════════════════════════
else:

    # ────────────────────────────────────────────────────────────────────────
    # 1. REGISTAR LANÇAMENTO
    # ────────────────────────────────────────────────────────────────────────
    if pagina == PAGINAS_TES[0]:
        st.markdown('<div class="section-title">➕ Registar Lançamento</div>',
                    unsafe_allow_html=True)
        if not GS_OK:
            st.markdown('<div class="alert-legal">🟡 Modo offline — dados do Google Sheets '
                        'não disponíveis. Partilhe o Sheet publicamente para ligar.</div>',
                        unsafe_allow_html=True)

        # Verificar confirmação de excepção pendente
        if st.session_state.confirmacao_pendente:
            pend = st.session_state.confirmacao_pendente
            st.markdown(f'<div class="alert-excecao">⚠️ <strong>Excepção Hierárquica Requerida</strong><br>'
                        f'{pend["aviso"]}</div>', unsafe_allow_html=True)
            col_y, col_n = st.columns(2)
            with col_y:
                if st.button("✅ Confirmar Excepção e Gravar", type="primary",
                             use_container_width=True):
                    reg = pend["reg"]
                    reg["excecao"] = pend["tipo"]
                    st.session_state.lancamentos.append(reg)
                    st.session_state.confirmacao_pendente = None
                    st.success(f"✅ Registo realizado com sucesso e base de dados actualizada. "
                               f"[{pend['tipo']}]")
                    st.rerun()
            with col_n:
                if st.button("❌ Cancelar", use_container_width=True):
                    st.session_state.confirmacao_pendente = None
                    st.rerun()
            st.stop()

        with st.form("form_lancamento", clear_on_submit=False):
            modalidade = st.selectbox(
                "Modalidade *",
                ["Defesa TFC","Pré-Defesa TFC","Relatório de Estágio"],
                help="Escolha a modalidade antes de seleccionar as funções")

            st.markdown("---")
            c1, c2 = st.columns(2)
            with c1:
                despacho  = st.text_input("Nº Despacho de Nomeação *",
                                          placeholder="Ex.: DESP-TFC-004/2026")
                acto      = st.selectbox("Tipo de Acto *", ["Defesa","Pré-Defesa","Estágio"])
                data_acto = st.date_input("Data do Acto *", value=datetime.date.today())
                curso     = st.selectbox("Curso / Departamento *", CURSOS)
            with c2:
                docente   = st.selectbox("Docente *", MEMBER_NAMES + ["[Externo]"])
                if docente == "[Externo]":
                    docente   = st.text_input("Nome completo do docente externo *")
                    grau_ext  = st.selectbox("Grau académico", ["PhD","MSc","Lic"])
                    cargo_ext = st.text_input("Cargo / Instituição")
                    mem_sel   = {"Nome":docente,"Grau":grau_ext,
                                 "Cargo":cargo_ext,"Departamento":"Externo"}
                else:
                    mem_sel = get_member(docente)
                grau_sel  = mem_sel.get("Grau","MSc")
                cargo_sel = mem_sel.get("Cargo","Docente")
                st.info(f"Nível: **{grau_sel}** | Cargo: **{cargo_sel}**")
                funcoes_disp = (FUNCOES_ESTAGIO if modalidade=="Relatório de Estágio"
                                else FUNCOES_TFC)
                funcao    = st.selectbox("Função no Júri *", funcoes_disp)
                tipo_opts = (["Estágio"] if modalidade=="Relatório de Estágio"
                             else ["Individual","Dupla"])
                tipo      = st.selectbox("Tipo de Defesa *", tipo_opts)

            # Campos do estudante
            st.markdown("**Dados do Estudante / Trabalho**")
            cs1, cs2, cs3 = st.columns([2,3,1])
            with cs1:
                estudante = st.text_input("Nome do Estudante(s) *",
                                          placeholder="Nome completo ou dupla")
            with cs2:
                titulo    = st.text_input("Título do Trabalho / Relatório *",
                                          placeholder="Título completo")
            with cs3:
                avaliacao = st.selectbox("Avaliação (0-20) *",
                                         ["—"] + ["Reprovado (0-9)"] +
                                         [str(i) for i in range(10,21)])

            bruto_calc = get_bruto(funcao, tipo, SUBSIDIOS)
            mes_ref    = mes_from_date(data_acto)
            st.markdown(
                f'<div class="alert-ok">✅ Valor bruto (VLOOKUP): <strong>{kz(bruto_calc)}</strong>'
                f' | Mês: <strong>{mes_ref}</strong></div>',
                unsafe_allow_html=True)

            submitted = st.form_submit_button("💾 Registar", type="primary",
                                              use_container_width=True)

        if submitted:
            errors = []
            if not despacho.strip():   errors.append("Nº de Despacho obrigatório")
            if not docente.strip():    errors.append("Nome do docente obrigatório")
            if not estudante.strip():  errors.append("Nome do estudante obrigatório")
            if not titulo.strip():     errors.append("Título do trabalho obrigatório")
            if avaliacao == "—":       errors.append("Avaliação obrigatória")
            if bruto_calc == 0:        errors.append("Combinação Função+Tipo sem valor")
            if check_duplicate(despacho.strip(), docente.strip()):
                errors.append(f"Erro: Registo duplicado — {docente} já consta no Despacho {despacho}")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                validacao = validar_funcao(grau_sel, cargo_sel, funcao, modalidade)
                reg = {
                    "despacho": despacho.strip(), "docente": docente.strip(),
                    "grau": grau_sel, "cargo": cargo_sel,
                    "funcao": funcao, "tipo": tipo, "modalidade": modalidade,
                    "curso": curso, "acto": acto,
                    "data": data_acto.strftime("%d/%m/%Y"), "mes_ref": mes_ref,
                    "bruto": bruto_calc, "excecao": None,
                    "estudante": estudante.strip(), "titulo_trabalho": titulo.strip(),
                    "avaliacao": avaliacao,
                }
                if validacao["ok"]:
                    st.session_state.lancamentos.append(reg)
                    st.success(
                        f"✅ Registo realizado com sucesso e base de dados actualizada. "
                        f"{docente} | {funcao} | {kz(bruto_calc)} | {mes_ref}")
                    if validacao["aviso"]:
                        st.markdown(
                            f'<div class="alert-excecao">{validacao["aviso"]}</div>',
                            unsafe_allow_html=True)
                    st.rerun()
                elif validacao["requer_confirmacao"]:
                    st.session_state.confirmacao_pendente = {
                        "reg": reg, "aviso": validacao["aviso"],
                        "tipo": validacao["tipo"],
                    }
                    st.rerun()
                else:
                    st.error(f"❌ {validacao['aviso']}")

        # Botões de navegação no fundo
        st.markdown("---")
        _, cn = st.columns([3,1])
        with cn:
            if st.button("Próximo: Dashboard ▶", use_container_width=True):
                st.session_state.pagina_idx = 1
                st.rerun()

    # ────────────────────────────────────────────────────────────────────────
    # 2. DASHBOARD EXECUTIVO
    # ────────────────────────────────────────────────────────────────────────
    elif pagina == PAGINAS_TES[1]:
        st.markdown('<div class="alert-info">ℹ️ Esta secção é apenas de consulta — '
                    'não requer preenchimento.</div>', unsafe_allow_html=True)
        df     = get_lancamentos_df()
        resumo = get_resumo()

        total_b = df["bruto"].sum()   if not df.empty else 0
        total_i = round(total_b * IRT_RATE, 2)
        total_l = total_b - total_i
        n_est   = df["estudante"].nunique() if not df.empty and "estudante" in df.columns else 0

        c1,c2,c3,c4 = st.columns(4)
        c1.markdown(f'<div class="kpi-card"><div class="kpi-label">TOTAL BRUTO</div>'
                    f'<div class="kpi-value">{kz(total_b)}</div></div>',unsafe_allow_html=True)
        c2.markdown(f'<div class="kpi-card red"><div class="kpi-label">IRT → AGT</div>'
                    f'<div class="kpi-value">{kz(total_i)}</div></div>',unsafe_allow_html=True)
        c3.markdown(f'<div class="kpi-card green"><div class="kpi-label">LÍQUIDO A PAGAR</div>'
                    f'<div class="kpi-value">{kz(total_l)}</div></div>',unsafe_allow_html=True)
        c4.markdown(f'<div class="kpi-card teal"><div class="kpi-label">ESTUDANTES AVALIADOS</div>'
                    f'<div class="kpi-value">{n_est}</div></div>',unsafe_allow_html=True)

        if not resumo.empty:
            col1,col2 = st.columns(2)
            with col1:
                st.markdown('<div class="section-title">Bruto / IRT / Líquido por Docente</div>',
                            unsafe_allow_html=True)
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Bruto",   x=resumo["docente"],
                                     y=resumo["total_bruto"], marker_color="#2E5088"))
                fig.add_trace(go.Bar(name="IRT",     x=resumo["docente"],
                                     y=resumo["irt_6_5"],    marker_color="#C00000"))
                fig.add_trace(go.Bar(name="Líquido", x=resumo["docente"],
                                     y=resumo["liquido"],    marker_color="#1E7145"))
                fig.update_layout(barmode="group",height=340,
                                  margin=dict(l=10,r=10,t=10,b=60),
                                  xaxis_tickangle=-30,
                                  legend=dict(orientation="h",y=1.1))
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                st.markdown('<div class="section-title">Total Bruto por Mês</div>',
                            unsafe_allow_html=True)
                mes_grp = resumo.groupby("mes_ref")["total_bruto"].sum().reset_index()
                fig2 = px.bar(mes_grp,x="mes_ref",y="total_bruto",
                              color_discrete_sequence=["#1F3864"],
                              labels={"mes_ref":"Mês","total_bruto":"Total Bruto (Kz)"})
                fig2.update_layout(height=340,margin=dict(l=10,r=10,t=10,b=40))
                st.plotly_chart(fig2, use_container_width=True)

            # Distribuição por modalidade
            if "modalidade" in df.columns and not df.empty:
                st.markdown('<div class="section-title">Distribuição por Modalidade</div>',
                            unsafe_allow_html=True)
                mod_grp = df.groupby("modalidade")["bruto"].agg(["count","sum"]).reset_index()
                mod_grp.columns = ["Modalidade","Nº Participações","Total Bruto (Kz)"]
                c1,c2 = st.columns(2)
                with c1:
                    fig3 = px.pie(mod_grp,names="Modalidade",values="Nº Participações",
                                  color_discrete_sequence=["#1F3864","#C00000","#1E7145"])
                    fig3.update_layout(height=280,margin=dict(l=5,r=5,t=5,b=5))
                    st.plotly_chart(fig3, use_container_width=True)
                with c2:
                    st.dataframe(mod_grp, use_container_width=True, hide_index=True)

            # Aproveitamento académico
            if "avaliacao" in df.columns and not df.empty:
                st.markdown('<div class="section-title">Aproveitamento Académico</div>',
                            unsafe_allow_html=True)
                av_df = df[df["avaliacao"].notna() & (df["avaliacao"]!="—")].copy()
                if not av_df.empty:
                    av_disp = av_df[["estudante","titulo_trabalho","avaliacao",
                                     "curso","mes_ref"]].drop_duplicates(subset=["estudante","titulo_trabalho"])
                    av_disp.columns = ["Estudante","Título","Avaliação","Curso","Mês"]
                    st.dataframe(av_disp, use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="alert-legal">ℹ Sem lançamentos. '
                        'Use "Registar Lançamento" para começar.</div>',unsafe_allow_html=True)

        st.markdown("---")
        cp, cn = st.columns(2)
        with cp:
            if st.button("◀ Registar Lançamento", use_container_width=True):
                st.session_state.pagina_idx = 0; st.rerun()
        with cn:
            if st.button("Resumo IRT ▶", use_container_width=True):
                st.session_state.pagina_idx = 2; st.rerun()

    # ────────────────────────────────────────────────────────────────────────
    # 3. RESUMO IRT
    # ────────────────────────────────────────────────────────────────────────
    elif pagina == PAGINAS_TES[2]:
        st.markdown('<div class="alert-info">ℹ️ Secção de consulta — IRT calculado sobre '
                    'o total mensal por docente.</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">🧾 Resumo IRT — Cálculo Mensal por Docente</div>',
                    unsafe_allow_html=True)
        st.markdown(
            '<div class="alert-legal">⭐ <strong>Regra fiscal (C.I.R.T. Art.º 67):</strong> '
            'IRT de 6,5% sobre o <strong>total bruto mensal acumulado</strong> por docente — '
            'não por sessão individual.</div>', unsafe_allow_html=True)

        resumo = get_resumo()
        if resumo.empty:
            st.info("Sem lançamentos ainda.")
        else:
            df_show = resumo.copy()
            for col in ["total_bruto","irt_6_5","liquido"]:
                df_show[col] = df_show[col].apply(kz)
            df_show.columns = ["Docente","Mês Ref.","Total Bruto Mensal",
                                "IRT 6,5% (sobre total)","Líquido a Pagar"]
            st.dataframe(df_show, use_container_width=True, hide_index=True)

            df_num  = get_resumo()
            irt_agt = df_num["irt_6_5"].sum()
            st.markdown(
                f'<div class="alert-error">📋 <strong>TOTAL IRT A ENTREGAR À AGT: '
                f'{kz(irt_agt)}</strong> | Prazo: dia 20 do mês seguinte.</div>',
                unsafe_allow_html=True)

        st.markdown("---")
        cp, cn = st.columns(2)
        with cp:
            if st.button("◀ Dashboard", use_container_width=True):
                st.session_state.pagina_idx = 1; st.rerun()
        with cn:
            if st.button("Registar Pagamento ▶", use_container_width=True):
                st.session_state.pagina_idx = 3; st.rerun()

    # ────────────────────────────────────────────────────────────────────────
    # 4. REGISTAR PAGAMENTO
    # ────────────────────────────────────────────────────────────────────────
    elif pagina == PAGINAS_TES[3]:
        st.markdown('<div class="section-title">💳 Registar Pagamento Efectuado</div>',
                    unsafe_allow_html=True)
        resumo = get_resumo()
        if resumo.empty:
            st.info("Sem lançamentos para pagar. Registe lançamentos primeiro.")
        else:
            with st.form("form_pagamento"):
                docente_p = st.selectbox("Docente", resumo["docente"].unique().tolist())
                meses_d   = resumo[resumo["docente"]==docente_p]["mes_ref"].tolist()
                mes_p     = st.selectbox("Mês de Referência", meses_d)
                row_sel   = resumo[(resumo["docente"]==docente_p)&(resumo["mes_ref"]==mes_p)]
                if not row_sel.empty:
                    liq_due = row_sel["liquido"].values[0]
                    irt_val = row_sel["irt_6_5"].values[0]
                    st.markdown(
                        f'<div class="alert-ok">Líquido a pagar: <strong>{kz(liq_due)}</strong> | '
                        f'IRT retido: <strong>{kz(irt_val)}</strong></div>',
                        unsafe_allow_html=True)
                    valor_pago   = st.number_input("Valor Pago (Kz)",min_value=0.0,
                                                   value=float(liq_due),step=100.0,format="%.2f")
                    canal_p      = st.selectbox("Canal", CANAIS)
                    data_pag     = st.date_input("Data do Pagamento")
                    comprovativo = st.text_input("Referência do Comprovativo")
                    if st.form_submit_button("✅ Confirmar Pagamento", type="primary"):
                        key_p = f"{docente_p}|{mes_p}"
                        st.session_state.pagamentos[key_p] = {
                            "valor_pago":valor_pago,"canal":canal_p,
                            "data":data_pag.strftime("%d/%m/%Y"),"comprovativo":comprovativo
                        }
                        saldo  = liq_due - valor_pago
                        estado = "PAGO" if saldo<=0 else ("PARCIAL" if valor_pago>0 else "EM ATRASO")
                        st.success(f"✅ Pagamento registado | Estado: {estado} | Saldo: {kz(max(saldo,0))}")
                        st.rerun()

        st.markdown("---")
        cp, cn = st.columns(2)
        with cp:
            if st.button("◀ Resumo IRT", use_container_width=True):
                st.session_state.pagina_idx = 2; st.rerun()
        with cn:
            if st.button("Lançamentos ▶", use_container_width=True):
                st.session_state.pagina_idx = 4; st.rerun()

    # ────────────────────────────────────────────────────────────────────────
    # 5. TODOS OS LANÇAMENTOS (com edição/eliminação)
    # ────────────────────────────────────────────────────────────────────────
    elif pagina == PAGINAS_TES[4]:
        st.markdown('<div class="section-title">📋 Todos os Lançamentos Registados</div>',
                    unsafe_allow_html=True)
        df = get_lancamentos_df()
        if df.empty:
            st.info("Sem lançamentos.")
        else:
            c1,c2,c3 = st.columns(3)
            with c1: f_doc = st.selectbox("Docente",   ["Todos"]+MEMBER_NAMES)
            with c2: f_mes = st.selectbox("Mês",       ["Todos"]+MESES_REF)
            with c3: f_mod = st.selectbox("Modalidade",["Todas","Defesa TFC",
                                                         "Pré-Defesa TFC","Relatório de Estágio"])
            df_f = df.copy()
            if f_doc!="Todos": df_f = df_f[df_f["docente"]==f_doc]
            if f_mes!="Todos": df_f = df_f[df_f["mes_ref"]==f_mes]
            if f_mod!="Todas" and "modalidade" in df_f.columns:
                df_f = df_f[df_f["modalidade"]==f_mod]

            st.markdown(f"**{len(df_f)} registos | Total: {kz(df_f['bruto'].sum())}**")

            # CRUD — eliminar individual
            for i_row, (orig_idx, row) in enumerate(df_f.iterrows()):
                excecao_badge = (f' <span class="badge-excecao">{row.get("excecao","")}</span>'
                                 if row.get("excecao") else "")
                with st.expander(
                    f"#{orig_idx+1} | {row['docente']} | {row.get('funcao','')} | "
                    f"{row['mes_ref']} | {kz(row['bruto'])}",
                    expanded=False):

                    cols = ["despacho","docente","grau","funcao","tipo","modalidade",
                            "curso","data","mes_ref","bruto","excecao",
                            "estudante","titulo_trabalho","avaliacao"]
                    display_cols = [c for c in cols if c in row.index]
                    st.json({c: str(row[c]) for c in display_cols})

                    if st.button(f"🗑 Eliminar registo #{orig_idx+1}",
                                 key=f"del_{orig_idx}", type="secondary"):
                        st.session_state.lancamentos.pop(orig_idx)
                        st.success(f"Registo #{orig_idx+1} eliminado.")
                        st.rerun()

            if st.session_state.admin_mode:
                if st.button("🗑 Limpar TODOS os lançamentos",type="secondary"):
                    st.session_state.lancamentos = []
                    st.rerun()

        st.markdown("---")
        cp, cn = st.columns(2)
        with cp:
            if st.button("◀ Pagamentos", use_container_width=True):
                st.session_state.pagina_idx = 3; st.rerun()
        with cn:
            if st.button("Exportar ▶", use_container_width=True):
                st.session_state.pagina_idx = 5; st.rerun()

    # ────────────────────────────────────────────────────────────────────────
    # 6. EXPORTAR
    # ────────────────────────────────────────────────────────────────────────
    elif pagina == PAGINAS_TES[5]:
        st.markdown('<div class="section-title">📥 Exportar Relatórios</div>',
                    unsafe_allow_html=True)
        df_lan = get_lancamentos_df()
        df_res = get_resumo()

        if df_lan.empty:
            st.info("Sem dados para exportar.")
        else:
            # Excel
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_lan.to_excel(writer, sheet_name="Lançamentos", index=False)
                if not df_res.empty:
                    df_res.to_excel(writer, sheet_name="Resumo_IRT", index=False)
                # Aproveitamento académico
                if "estudante" in df_lan.columns:
                    av_df = df_lan[["estudante","titulo_trabalho","avaliacao",
                                    "curso","docente","mes_ref"]].drop_duplicates(
                                    subset=["estudante","titulo_trabalho"])
                    av_df.to_excel(writer, sheet_name="Aproveitamento", index=False)
            buf.seek(0)
            st.download_button(
                label="📊 Descarregar Excel completo",
                data=buf,
                file_name=f"ISPTLO_Juris_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # PDF Tesouraria
            t_bruto = df_res["total_bruto"].sum() if not df_res.empty else 0
            t_irt   = df_res["irt_6_5"].sum()     if not df_res.empty else 0
            t_liq   = df_res["liquido"].sum()      if not df_res.empty else 0
            render_tesouraria_pdf_block(
                mes_filtro="Todos", df_resumo=df_res,
                total_bruto=t_bruto, total_irt=t_irt, total_liquido=t_liq)

            # JSON — apenas Admin
            if st.session_state.admin_mode:
                st.markdown('<div class="alert-error">🔒 Ficheiro Admin — uso restrito.</div>',
                            unsafe_allow_html=True)
                json_data = json.dumps({
                    "lancamentos":  st.session_state.lancamentos,
                    "pagamentos":   st.session_state.pagamentos,
                    "exportado_em": TODAY,
                }, ensure_ascii=False, indent=2)
                st.download_button(
                    label="💾 Backup JSON (Admin)",
                    data=json_data.encode("utf-8"),
                    file_name=f"ISPTLO_backup_{datetime.date.today().strftime('%Y%m%d')}.json",
                    mime="application/json")

            st.markdown(
                '<div class="alert-legal">📋 O Excel inclui: Lançamentos, Resumo IRT e '
                'Aproveitamento Académico por estudante.</div>',unsafe_allow_html=True)

        st.markdown("---")
        if st.button("◀ Todos os Lançamentos", use_container_width=True):
            st.session_state.pagina_idx = 4; st.rerun()

# ── FOOTER ILR ────────────────────────────────────────────────────────────────
st.markdown("---")
admin_badge = ('<span class="badge-admin">ADMIN</span>'
               if st.session_state.get("admin_mode") else "")
logo_img_f  = (f'<img src="{LOGO_SRC}" style="width:36px;height:36px;border-radius:50%;'
               f'border:2px solid #FFD700;flex-shrink:0"/>' if LOGO_SRC else "")
st.markdown(
    f'<div style="display:flex;align-items:center;gap:12px;'
    f'background:linear-gradient(135deg,#060E1F,#1F3864);'
    f'border-radius:10px;padding:.8rem 1.2rem;border-top:2px solid #FFD700;">'
    f'{logo_img_f}'
    f'<div style="flex:1">'
    f'<div style="color:#FFD700;font-size:13px;font-weight:700">'
    f'Ph.D. Ideleichy Lombillo Rivero &nbsp;|&nbsp; ILR — Academic Solutions'
    f'&nbsp;&nbsp;{admin_badge}</div>'
    f'<div style="color:#AABBDD;font-size:10px;font-style:italic;margin-top:2px">'
    f'Gestão Inteligente para Instituições de Excelência</div>'
    f'<div style="color:#7799CC;font-size:9px;margin-top:3px">'
    f'Plataforma de Gestão de Júris Académicos v{APP_VER} &nbsp;|&nbsp; '
    f'D.P. nº 191/18 (ECDES) &nbsp;|&nbsp; C.I.R.T. Lei nº 19/14 &nbsp;|&nbsp; '
    f'Ofício 009/25 + Emenda DASG 023/2025</div>'
    f'</div></div>',
    unsafe_allow_html=True,
)
