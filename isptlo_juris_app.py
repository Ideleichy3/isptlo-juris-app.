"""
ISPTLO — Aplicação de Gestão de Pagamentos de Júri de TFC
Versão 3.0 | Python + Streamlit | Multiplataforma (Windows + Mobile)

Instalação:
    pip install streamlit pandas openpyxl plotly fpdf2

Execução:
    streamlit run isptlo_juris_app.py

Acesso mobile: abrir o URL no browser do telemóvel (mesma rede Wi-Fi)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime
import io
import json
from pathlib import Path

# ── CONFIGURAÇÃO DA APP ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="ISPTLO — Gestão de Júris de TFC",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CONSTANTES ─────────────────────────────────────────────────────────────────
IRT_RATE = 0.065

MEMBERS = [
    "Manuel Octávio Isaac Spínola",
    "Ideleichy Lombillo Rivero",
    "Elizabeth González",
    "Domingos Lunga",
    "José Fernando Manuel",
    "Walquiria Chissimo",
    "Félix Palau",
    "Marilda Augusto",
    "José Monteiro",
    "Domingos Ngando",
]

MEMBER_LEVEL = {
    "Manuel Octávio Isaac Spínola": "PhD",
    "Ideleichy Lombillo Rivero": "PhD",
    "Elizabeth González": "MSc",
    "Domingos Lunga": "MSc",
    "José Fernando Manuel": "MSc",
    "Walquiria Chissimo": "MSc",
    "Félix Palau": "MSc",
    "Marilda Augusto": "Lic",
    "José Monteiro": "Lic",
    "Domingos Ngando": "Lic",
}

CURSOS = [
    "Contabilidade e Gestão / CSH",
    "Ensino Primário / CSH",
    "Agronomia / Engenharias e Tecnologias",
    "Enfermagem / Ciências da Saúde",
    "Todos os Cursos (Especial Ph.D.)",
]

FUNCOES_PERMITIDAS = {
    "PhD":  ["Presidente de Júri", "1º Vogal Arguente", "2º Vogal Tutor/Orientador"],
    "MSc":  ["1º Vogal Arguente", "2º Vogal Tutor/Orientador", "Co-tutor"],
    "Lic":  ["Secretário", "Co-tutor"],
}

# Tabela de subsídios brutos: (Função, Tipo) → valor Kz
SUBSIDIOS = {
    ("Presidente de Júri",         "Individual"): 20_000,
    ("Presidente de Júri",         "Dupla"):      27_500,
    ("1º Vogal Arguente",          "Individual"): 20_000,
    ("1º Vogal Arguente",          "Dupla"):      27_500,
    ("2º Vogal Tutor/Orientador",  "Individual"): 80_000,
    ("2º Vogal Tutor/Orientador",  "Dupla"):     104_000,
    ("Co-tutor",                   "Individual"): 40_000,
    ("Co-tutor",                   "Dupla"):      72_000,
    ("Secretário",                 "Individual"): 10_000,
    ("Secretário",                 "Dupla"):      15_000,
}

MESES_REF = [
    "Jan/2026","Fev/2026","Mar/2026","Abr/2026","Mai/2026","Jun/2026",
    "Jul/2026","Ago/2026","Set/2026","Out/2026","Nov/2026","Dez/2026",
]

CANAIS = ["Transferência Bancária", "Numerário", "Cheque", "Multicaixa"]

RESULTADOS_DEFESA = (
    ["Reprovado (0-9)"] + [str(i) for i in range(10, 21)]
)
RESULTADOS_PRE = ["Aprovado", "Reprovado", "Pendente", "Condicionado"]

# ── CSS PERSONALIZADO ─────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Sidebar */
    [data-testid="stSidebar"] { background: #1F3864; }
    [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stRadio label { color: white !important; }
    [data-testid="stSidebar"] .stSelectbox label { color: #FFD700 !important; font-weight: bold; }

    /* Header cards */
    .kpi-card {
        background: #1F3864; border-radius: 12px; padding: 1rem 1.2rem;
        color: white; text-align: center; margin-bottom: 0.5rem;
    }
    .kpi-card .kpi-label { font-size: 12px; opacity: 0.85; margin-bottom: 4px; }
    .kpi-card .kpi-value { font-size: 26px; font-weight: 700; line-height: 1.1; }
    .kpi-card.red  { background: #C00000; }
    .kpi-card.green { background: #1E7145; }
    .kpi-card.amber { background: #854F0B; }

    /* Alert boxes */
    .alert-legal {
        background: #FFF2CC; border-left: 5px solid #854F0B;
        padding: 0.75rem 1rem; border-radius: 6px; font-size: 13px;
        color: #5C3D00; margin: 0.5rem 0;
    }
    .alert-error {
        background: #FFCCCC; border-left: 5px solid #C00000;
        padding: 0.75rem 1rem; border-radius: 6px; font-size: 13px;
        color: #5C0000; margin: 0.5rem 0;
    }
    .alert-ok {
        background: #C6EFCE; border-left: 5px solid #1E7145;
        padding: 0.75rem 1rem; border-radius: 6px; font-size: 13px;
        color: #0A3D1F; margin: 0.5rem 0;
    }

    /* Section titles */
    .section-title {
        font-size: 16px; font-weight: 700; color: #1F3864;
        border-bottom: 3px solid #1F3864; padding-bottom: 4px;
        margin: 1.5rem 0 0.8rem 0;
    }

    /* Main header */
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E5088 100%);
        color: white; padding: 1.2rem 1.5rem; border-radius: 12px;
        margin-bottom: 1.2rem;
    }
    .main-header h1 { color: #FFD700; margin: 0; font-size: 22px; }
    .main-header p  { color: #CCDDFF; margin: 4px 0 0 0; font-size: 13px; }

    /* Dataframes */
    .stDataFrame { border: 1px solid #BDD7EE; border-radius: 8px; }

    /* Inputs */
    .stSelectbox > div > div { border: 1.5px solid #BDD7EE !important; }
    .stNumberInput > div > div { border: 1.5px solid #FFD700 !important; }
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────────────────────
if "lancamentos" not in st.session_state:
    st.session_state.lancamentos = []
if "pagamentos" not in st.session_state:
    st.session_state.pagamentos = {}  # key: (docente, mes) → valor_pago
if "modo" not in st.session_state:
    st.session_state.modo = "Tesouraria"

# ── FUNÇÕES UTILITÁRIAS ───────────────────────────────────────────────────────
def kz(val):
    """Formata valor em Kz com separador de milhares."""
    if val is None or val == 0:
        return "— Kz"
    return f"{val:,.2f} Kz".replace(",", "X").replace(".", ",").replace("X", ".")

def get_bruto(funcao, tipo):
    return SUBSIDIOS.get((funcao, tipo), 0)

def calcular_resumo_irt(lancamentos_df):
    """
    LÓGICA CONTABILÍSTICA CORRECTA:
    Agrupa por (docente, mês) → soma total bruto mensal → aplica IRT 6,5% sobre o total.
    """
    if lancamentos_df.empty:
        return pd.DataFrame()
    grp = (lancamentos_df
           .groupby(["docente", "mes_ref"])["bruto"]
           .sum()
           .reset_index()
           .rename(columns={"bruto": "total_bruto"}))
    grp["irt_6_5"] = (grp["total_bruto"] * IRT_RATE).round(2)
    grp["liquido"]  = grp["total_bruto"] - grp["irt_6_5"]
    return grp

def get_lancamentos_df():
    if not st.session_state.lancamentos:
        return pd.DataFrame(columns=["despacho","docente","nivel","funcao","tipo","data","mes_ref","bruto"])
    return pd.DataFrame(st.session_state.lancamentos)

def get_resumo():
    return calcular_resumo_irt(get_lancamentos_df())

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🎓 ISPTLO")
    st.markdown("**Gestão de Júri de TFC**")
    st.markdown("---")

    modo = st.radio("Modo de Utilização", ["👤 Docente", "🏦 Tesouraria"],
                    index=0 if st.session_state.modo == "Docente" else 1)
    st.session_state.modo = "Docente" if "Docente" in modo else "Tesouraria"

    st.markdown("---")
    st.markdown("**Navegação**")
    if st.session_state.modo == "Docente":
        pagina = st.selectbox("Secção", [
            "💰 Calcular os meus haveres",
        ])
    else:
        pagina = st.selectbox("Secção", [
            "📊 Dashboard Executivo",
            "➕ Registar Lançamento",
            "🧾 Resumo IRT por Docente",
            "💳 Registar Pagamento",
            "📋 Todos os Lançamentos",
            "📥 Exportar Excel / PDF",
        ])

    st.markdown("---")
    st.markdown(
        "<small style='color:#AABBDD'>Base legal: D.P. nº 191/18 (ECDES)<br>"
        "IRT: Lei nº 19/14 (C.I.R.T.) Art.º 67<br>"
        "Ofício 009/25 + Emenda DASG 023/2025</small>",
        unsafe_allow_html=True
    )

# ── HEADER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🎓 ISPTLO — Mapa de Controlo de Júri de TFC</h1>
  <p>Instrumento oficial de gestão financeira | Versão 3.0 | IRT calculado sobre total mensal por docente</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MODO DOCENTE — Calculadora individual
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.modo == "Docente":
    st.markdown('<div class="section-title">💰 Calculadora de Haveres — Modo Docente</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<div class="alert-legal">ℹ O IRT de 6,5% é retido sobre o <strong>total mensal acumulado</strong> '
        'das suas participações, não por sessão individual. Introduza abaixo todas as participações '
        'do mês para obter o valor líquido correcto.</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        docente_sel = st.selectbox("O meu nome", MEMBERS)
        mes_sel     = st.selectbox("Mês de referência", MESES_REF, index=3)

    with col2:
        nivel = MEMBER_LEVEL.get(docente_sel, "MSc")
        funcoes_disp = FUNCOES_PERMITIDAS.get(nivel, [])
        st.info(f"**Nível:** {nivel} | **Funções permitidas:** {', '.join(funcoes_disp)}")

    st.markdown("**Adicionar participações do mês:**")
    n_part = st.number_input("Número de participações neste mês", 1, 20, 1)

    total_bruto = 0
    participacoes = []
    for i in range(int(n_part)):
        with st.expander(f"Participação {i+1}", expanded=True):
            c1, c2 = st.columns(2)
            with c1:
                func_i = st.selectbox(f"Função", funcoes_disp,
                                      key=f"func_{i}", help="Seleccione a função que desempenhou nesta banca")
            with c2:
                tipo_i = st.selectbox(f"Tipo de defesa", ["Individual","Dupla"],
                                      key=f"tipo_{i}", help="Individual = 1 estudante | Dupla = 2 estudantes")
            bruto_i = get_bruto(func_i, tipo_i)
            st.markdown(f"**Valor bruto desta participação:** `{kz(bruto_i)}`")
            total_bruto += bruto_i
            participacoes.append({"funcao": func_i, "tipo": tipo_i, "bruto": bruto_i})

    st.markdown("---")
    irt    = round(total_bruto * IRT_RATE, 2)
    liquido= total_bruto - irt

    c1, c2, c3 = st.columns(3)
    c1.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">TOTAL BRUTO MENSAL</div>
        <div class="kpi-value">{kz(total_bruto)}</div></div>""", unsafe_allow_html=True)
    c2.markdown(f"""<div class="kpi-card red">
        <div class="kpi-label">RETENÇÃO IRT (6,5%) → AGT</div>
        <div class="kpi-value">{kz(irt)}</div></div>""", unsafe_allow_html=True)
    c3.markdown(f"""<div class="kpi-card green">
        <div class="kpi-label">VALOR LÍQUIDO A RECEBER</div>
        <div class="kpi-value">{kz(liquido)}</div></div>""", unsafe_allow_html=True)

    st.markdown(
        f'<div class="alert-legal">📋 Cálculo: {kz(total_bruto)} × 6,5% = {kz(irt)} (IRT retido) '
        f'| Líquido = {kz(total_bruto)} − {kz(irt)} = {kz(liquido)}</div>',
        unsafe_allow_html=True
    )

    # Detalhe das participações
    if participacoes:
        df_det = pd.DataFrame(participacoes)
        df_det["bruto_fmt"] = df_det["bruto"].apply(kz)
        st.dataframe(df_det[["funcao","tipo","bruto_fmt"]].rename(columns={
            "funcao":"Função","tipo":"Tipo","bruto_fmt":"Valor Bruto"
        }), use_container_width=True, hide_index=True)

    # =========================================================
    # EXPORTAÇÃO PROFISSIONAL DE RECIBO PDF (Versão ISPTLO v3.0)
    # =========================================================
    if st.button("📥 Gerar e Descarregar Recibo Oficial"):
        from fpdf import FPDF
        
        # 1. Função auxiliar para formatar Kwanza
        def fmt_kz(valor):
            return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " Kz"

        # 2. Verificar se a tabela tem dados e calcular totais internamente
        # Usamos 'df_filtrado' que é a base da sua tabela no ecrã
        if 'df_filtrado' in locals() and not df_filtrado.empty:
            res_bruto = df_filtrado['Valor Bruto'].sum()
            res_irt = res_bruto * 0.065
            res_liquido = res_bruto - res_irt
            
            # Criar o PDF profissional
            pdf = FPDF()
            pdf.add_page()
            
            # Cabeçalho Institucional (Sem acentos para evitar erros de fonte)
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, txt="INSTITUTO SUPERIOR POLITECNICO DO LIBOLO", ln=True, align='C')
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, txt="DIRECCAO DE INVESTIGACAO CIENTIFICA E POS-GRADUACAO", ln=True, align='C')
            pdf.ln(5)
            pdf.line(10, 32, 200, 32)
            pdf.ln(5)

            # Identificação do Docente
            nome_doc = docente_selecionado if 'docente_selecionado' in locals() else "Docente ISPTLO"
            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 10, txt=f"BENEFICIARIO: {nome_doc.upper()}", ln=True)
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 7, txt="ASSUNTO: Recibo de Honorarios por Participacao em Juris de TFC", ln=True)
            pdf.ln(5)

            # Tabela de Detalhe (Captura o que está na imagem image_17.png)
            pdf.set_fill_color(31, 56, 100) # Azul ISPTLO
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", "B", 10)
            
            pdf.cell(90, 8, "Funcao no Juri", border=1, fill=True)
            pdf.cell(40, 8, "Tipo", border=1, fill=True)
            pdf.cell(50, 8, "Valor Bruto", border=1, fill=True, align='R')
            pdf.ln()
            
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", "", 9)
            
            for index, row in df_filtrado.iterrows():
                pdf.cell(90, 8, str(row['Função']), border=1)
                pdf.cell(40, 8, str(row['Tipo']), border=1)
                pdf.cell(50, 8, fmt_kz(row['Valor Bruto']), border=1, align='R')
                pdf.ln()

            # Resumo Financeiro Final
            pdf.ln(10)
            pdf.set_font("Arial", "B", 11)
            pdf.cell(130, 8, txt="TOTAL BRUTO MENSAL:", border=0)
            pdf.cell(50, 8, txt=fmt_kz(res_bruto), border=0, align='R', ln=True)
            
            pdf.set_font("Arial", "", 11)
            pdf.cell(130, 8, txt="RETENCAO DE IRT (6,5%):", border=0)
            pdf.cell(50, 8, txt=f"- {fmt_kz(res_irt)}", border=0, align='R', ln=True)
            
            pdf.line(140, pdf.get_y()+2, 190, pdf.get_y()+2)
            pdf.ln(4)
            
            pdf.set_font("Arial", "B", 12)
            pdf.cell(130, 10, txt="VALOR LIQUIDO A RECEBER:", border=0)
            pdf.cell(50, 10, txt=fmt_kz(res_liquido), border=0, align='R', ln=True)

            # Rodapé Legal
            pdf.ln(20)
            pdf.set_font("Arial", "I", 8)
            pdf.multi_cell(0, 5, txt="Base Legal: Decreto Presidencial n 191/18 (ECDES) e Lei n 19/14 (C.I.R.T.). Documento processado pelo Sistema de Gestao ISPTLO-JURIS v3.0.")
            
            # Geração do Ficheiro para Descarga
            pdf_bytes = pdf.output()
            st.download_button(
                label="✅ Descarregar Recibo PDF Final",
                data=bytes(pdf_bytes),
                file_name=f"Recibo_Juri_{nome_doc.replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
            st.success("Recibo gerado com sucesso!")
        else:
            st.error("Erro: Não foi possível capturar os dados da tabela para o PDF.")


# ══════════════════════════════════════════════════════════════════════════════
# MODO TESOURARIA
# ══════════════════════════════════════════════════════════════════════════════
else:

    # ── Dashboard Executivo ────────────────────────────────────────────────────
    if pagina == "📊 Dashboard Executivo":
        df = get_lancamentos_df()
        resumo = get_resumo()

        total_bruto = df["bruto"].sum() if not df.empty else 0
        total_irt   = round(total_bruto * IRT_RATE, 2)
        total_liq   = total_bruto - total_irt

        # KPIs
        c1,c2,c3,c4 = st.columns(4)
        c1.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">TOTAL BRUTO LANÇADO</div>
            <div class="kpi-value">{kz(total_bruto)}</div></div>""",unsafe_allow_html=True)
        c2.markdown(f"""<div class="kpi-card red">
            <div class="kpi-label">TOTAL IRT → AGT</div>
            <div class="kpi-value">{kz(total_irt)}</div></div>""",unsafe_allow_html=True)
        c3.markdown(f"""<div class="kpi-card green">
            <div class="kpi-label">TOTAL LÍQUIDO A PAGAR</div>
            <div class="kpi-value">{kz(total_liq)}</div></div>""",unsafe_allow_html=True)
        n_doc = len(resumo["docente"].unique()) if not resumo.empty else 0
        c4.markdown(f"""<div class="kpi-card amber">
            <div class="kpi-label">DOCENTES COM LANÇAMENTOS</div>
            <div class="kpi-value">{n_doc}</div></div>""",unsafe_allow_html=True)

        st.markdown(
            '<div class="alert-legal">⚠ <strong>Nota AGT:</strong> O valor de IRT indicado deve ser declarado e '
            'entregue à Administração Geral Tributária até ao dia 20 do mês seguinte (C.I.R.T. Art.º 67).</div>',
            unsafe_allow_html=True
        )

        if not resumo.empty:
            col1, col2 = st.columns(2)

            with col1:
                st.markdown('<div class="section-title">IRT vs Líquido por Docente (mês actual)</div>',
                            unsafe_allow_html=True)
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Bruto",   x=resumo["docente"], y=resumo["total_bruto"],
                                     marker_color="#2E5088"))
                fig.add_trace(go.Bar(name="IRT",     x=resumo["docente"], y=resumo["irt_6_5"],
                                     marker_color="#C00000"))
                fig.add_trace(go.Bar(name="Líquido", x=resumo["docente"], y=resumo["liquido"],
                                     marker_color="#1E7145"))
                fig.update_layout(barmode="group", height=360, margin=dict(l=20,r=20,t=20,b=60),
                                   xaxis_tickangle=-30, legend=dict(orientation="h",y=1.1))
                st.plotly_chart(fig, use_container_width=True)

            with col2:
                st.markdown('<div class="section-title">Total Bruto por Mês de Referência</div>',
                            unsafe_allow_html=True)
                mes_grp = resumo.groupby("mes_ref")["total_bruto"].sum().reset_index()
                fig2 = px.bar(mes_grp, x="mes_ref", y="total_bruto",
                              color_discrete_sequence=["#1F3864"],
                              labels={"mes_ref":"Mês","total_bruto":"Total Bruto (Kz)"})
                fig2.update_layout(height=360, margin=dict(l=20,r=20,t=20,b=60))
                st.plotly_chart(fig2, use_container_width=True)

            # Participações por função
            st.markdown('<div class="section-title">Distribuição por Função no Júri</div>',
                        unsafe_allow_html=True)
            if not df.empty:
                func_grp = df.groupby("funcao")["bruto"].agg(["count","sum"]).reset_index()
                func_grp.columns = ["Função","Nº Participações","Total Bruto (Kz)"]
                c1, c2 = st.columns(2)
                with c1:
                    fig3 = px.pie(func_grp, names="Função", values="Nº Participações",
                                  color_discrete_sequence=px.colors.sequential.Blues_r)
                    fig3.update_layout(height=320,margin=dict(l=10,r=10,t=10,b=10))
                    st.plotly_chart(fig3, use_container_width=True)
                with c2:
                    st.dataframe(func_grp, use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="alert-legal">ℹ Sem lançamentos ainda. '
                        'Use "Registar Lançamento" para começar.</div>', unsafe_allow_html=True)

    # ── Registar Lançamento ───────────────────────────────────────────────────
    elif pagina == "➕ Registar Lançamento":
        st.markdown('<div class="section-title">➕ Registar Novo Lançamento (por sessão)</div>',
                    unsafe_allow_html=True)
        st.markdown(
            '<div class="alert-legal">ℹ Registe cada participação individual. '
            'O sistema agrupará automaticamente por docente/mês para calcular o IRT correcto.</div>',
            unsafe_allow_html=True
        )

        with st.form("form_lancamento"):
            c1, c2 = st.columns(2)
            with c1:
                despacho = st.text_input("Nº do Despacho de Nomeação",
                                         placeholder="Ex.: DESP-TFC-004/2026",
                                         help="Número do despacho oficial de nomeação do júri")
                acto     = st.selectbox("Tipo de Acto", ["Defesa", "Pré-Defesa"])
                data_acto= st.date_input("Data do Acto", value=datetime.date.today())
                curso    = st.selectbox("Curso / Departamento", CURSOS)
            with c2:
                docente  = st.selectbox("Docente", MEMBERS)
                nivel_d  = MEMBER_LEVEL.get(docente, "MSc")
                st.info(f"Nível: **{nivel_d}**")
                funcoes_d= FUNCOES_PERMITIDAS.get(nivel_d, [])
                funcao   = st.selectbox("Função no Júri", funcoes_d)
                tipo     = st.selectbox("Tipo de Defesa", ["Individual", "Dupla"])

            bruto_calc = get_bruto(funcao, tipo)
            mes_auto   = data_acto.strftime("%b/%Y").capitalize()
            # Fix Portuguese month abbreviation
            meses_pt = {"Jan":"Jan","Feb":"Fev","Mar":"Mar","Apr":"Abr","May":"Mai",
                        "Jun":"Jun","Jul":"Jul","Aug":"Ago","Sep":"Set","Oct":"Out",
                        "Nov":"Nov","Dec":"Dez"}
            mes_ref = meses_pt.get(data_acto.strftime("%b"), data_acto.strftime("%b")) + "/" + str(data_acto.year)

            st.markdown(f"""
            <div class="alert-ok">
            ✅ <strong>Valor bruto calculado automaticamente (VLOOKUP):</strong>
            {kz(bruto_calc)} | Mês de referência: <strong>{mes_ref}</strong>
            </div>""", unsafe_allow_html=True)

            if bruto_calc == 0:
                st.markdown('<div class="alert-error">⚠ Combinação Função+Tipo sem valor definido na tabela.</div>',
                            unsafe_allow_html=True)

            submitted = st.form_submit_button("💾 Registar Lançamento", type="primary")
            if submitted:
                if not despacho.strip():
                    st.error("O Nº do Despacho é obrigatório.")
                elif bruto_calc == 0:
                    st.error("Valor bruto = 0. Verifique a combinação Função/Tipo.")
                else:
                    reg = {
                        "despacho": despacho.strip(),
                        "docente":  docente,
                        "nivel":    nivel_d,
                        "funcao":   funcao,
                        "tipo":     tipo,
                        "acto":     acto,
                        "curso":    curso,
                        "data":     data_acto.strftime("%d/%m/%Y"),
                        "mes_ref":  mes_ref,
                        "bruto":    bruto_calc,
                    }
                    st.session_state.lancamentos.append(reg)
                    st.success(f"✅ Lançamento registado: {docente} | {funcao} | {kz(bruto_calc)} | {mes_ref}")
                    st.rerun()

    # ── Resumo IRT ────────────────────────────────────────────────────────────
    elif pagina == "🧾 Resumo IRT por Docente":
        st.markdown('<div class="section-title">🧾 Resumo IRT — Cálculo Mensal por Docente</div>',
                    unsafe_allow_html=True)
        st.markdown("""
        <div class="alert-legal">
        ⭐ <strong>Regra fiscal imperativa (C.I.R.T. Art.º 67):</strong> O IRT de 6,5% é calculado sobre
        o <strong>total bruto mensal acumulado</strong> por docente — soma de TODAS as participações no mês —
        NÃO por sessão individual.
        <br><br>
        <strong>Exemplo:</strong> Ideleichy Lombillo — Mar/2026:
        20.000 (presidente indiv.) + 27.500 (pres. dupla) + 104.000 (tutora dupla)
        = <strong>151.500 Kz</strong> → IRT = 151.500 × 6,5% = <strong>9.847,50 Kz</strong>
        → Líquido = <strong>141.652,50 Kz</strong>
        </div>""", unsafe_allow_html=True)

        resumo = get_resumo()
        if resumo.empty:
            st.info("Sem lançamentos. Registe participações primeiro.")
        else:
            # Format display
            df_show = resumo.copy()
            df_show["irt_6_5"]     = df_show["irt_6_5"].apply(kz)
            df_show["total_bruto"] = df_show["total_bruto"].apply(kz)
            df_show["liquido"]     = df_show["liquido"].apply(kz)
            df_show = df_show.rename(columns={
                "docente":"Docente","mes_ref":"Mês Ref.",
                "total_bruto":"Total Bruto Mensal","irt_6_5":"IRT 6,5% (sobre total)","liquido":"Líquido a Pagar"
            })
            st.dataframe(df_show, use_container_width=True, hide_index=True)

            # Detail per docente
            st.markdown('<div class="section-title">Detalhe por Docente (lançamentos individuais)</div>',
                        unsafe_allow_html=True)
            df_lan = get_lancamentos_df()
            docente_sel2 = st.selectbox("Filtrar por docente", ["Todos"] + MEMBERS)
            df_filtered = df_lan if docente_sel2 == "Todos" else df_lan[df_lan["docente"]==docente_sel2]
            if not df_filtered.empty:
                df_disp = df_filtered.copy()
                df_disp["bruto"] = df_disp["bruto"].apply(kz)
                st.dataframe(df_disp[["despacho","docente","funcao","tipo","data","mes_ref","bruto"]]
                             .rename(columns={"despacho":"Despacho","docente":"Docente","funcao":"Função",
                                              "tipo":"Tipo","data":"Data","mes_ref":"Mês","bruto":"Bruto"}),
                             use_container_width=True, hide_index=True)

            # Total IRT para AGT
            df_num = get_resumo()
            total_irt_agt = df_num["irt_6_5"].sum() if not df_num.empty else 0
            st.markdown(f"""
            <div class="alert-error">
            📋 <strong>TOTAL IRT A ENTREGAR À AGT: {kz(total_irt_agt)}</strong><br>
            Prazo: até ao dia 20 do mês seguinte ao do pagamento (C.I.R.T. Art.º 67).
            </div>""", unsafe_allow_html=True)

    # ── Registar Pagamento ────────────────────────────────────────────────────
    elif pagina == "💳 Registar Pagamento":
        st.markdown('<div class="section-title">💳 Registar Pagamento Efectuado</div>', unsafe_allow_html=True)
        resumo = get_resumo()
        if resumo.empty:
            st.info("Sem registos para pagar. Registe lançamentos primeiro.")
        else:
            with st.form("form_pagamento"):
                docente_p = st.selectbox("Docente", resumo["docente"].unique().tolist())
                mes_p     = st.selectbox("Mês de Referência",
                                         resumo[resumo["docente"]==docente_p]["mes_ref"].tolist()
                                         if docente_p else [])
                row_sel = resumo[(resumo["docente"]==docente_p) & (resumo["mes_ref"]==mes_p)]
                if not row_sel.empty:
                    liq_due = row_sel["liquido"].values[0]
                    irt_val = row_sel["irt_6_5"].values[0]
                    st.markdown(f"""
                    <div class="alert-ok">
                    Líquido a pagar: <strong>{kz(liq_due)}</strong> |
                    IRT retido: <strong>{kz(irt_val)}</strong>
                    </div>""", unsafe_allow_html=True)
                    valor_pago  = st.number_input("Valor Efectivamente Pago (Kz)", min_value=0.0,
                                                   value=float(liq_due), step=100.0,
                                                   format="%.2f")
                    canal_p     = st.selectbox("Canal de Pagamento", CANAIS)
                    data_pag    = st.date_input("Data do Pagamento")
                    comprovativo= st.text_input("Referência do Comprovativo")

                    pag_sub = st.form_submit_button("✅ Confirmar Pagamento", type="primary")
                    if pag_sub:
                        key_p = f"{docente_p}|{mes_p}"
                        st.session_state.pagamentos[key_p] = {
                            "valor_pago": valor_pago, "canal": canal_p,
                            "data": data_pag.strftime("%d/%m/%Y"), "comprovativo": comprovativo
                        }
                        saldo = liq_due - valor_pago
                        estado = "PAGO" if saldo <= 0 else ("PARCIAL" if valor_pago > 0 else "EM ATRASO")
                        st.success(f"✅ Pagamento registado | Estado: {estado} | Saldo: {kz(max(saldo,0))}")
                        st.rerun()

    # ── Todos os Lançamentos ──────────────────────────────────────────────────
    elif pagina == "📋 Todos os Lançamentos":
        st.markdown('<div class="section-title">📋 Todos os Lançamentos Registados</div>', unsafe_allow_html=True)
        df = get_lancamentos_df()
        if df.empty:
            st.info("Sem lançamentos ainda.")
        else:
            # Filters
            col1, col2, col3 = st.columns(3)
            with col1: f_doc = st.selectbox("Docente", ["Todos"] + MEMBERS)
            with col2: f_mes = st.selectbox("Mês", ["Todos"] + MESES_REF)
            with col3: f_func = st.selectbox("Função", ["Todas"] + list(SUBSIDIOS.keys().__class__()))

            df_f = df.copy()
            if f_doc != "Todos": df_f = df_f[df_f["docente"]==f_doc]
            if f_mes != "Todos": df_f = df_f[df_f["mes_ref"]==f_mes]

            total_f = df_f["bruto"].sum()
            st.markdown(f"**{len(df_f)} registos | Total bruto filtrado: {kz(total_f)}**")

            df_disp = df_f.copy()
            df_disp["bruto"] = df_disp["bruto"].apply(kz)
            st.dataframe(df_disp.rename(columns={
                "despacho":"Despacho","docente":"Docente","nivel":"Nível",
                "funcao":"Função","tipo":"Tipo","acto":"Acto","curso":"Curso",
                "data":"Data","mes_ref":"Mês Ref.","bruto":"Bruto"
            }), use_container_width=True, hide_index=True)

            if st.button("🗑 Limpar todos os lançamentos", type="secondary"):
                st.session_state.lancamentos = []
                st.rerun()

    # ── Exportar ──────────────────────────────────────────────────────────────
    elif pagina == "📥 Exportar Excel / PDF":
        st.markdown('<div class="section-title">📥 Exportar Relatórios</div>', unsafe_allow_html=True)

        df_lan = get_lancamentos_df()
        df_res = get_resumo()

        if df_lan.empty:
            st.info("Sem dados para exportar.")
        else:
            # Excel export
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_lan.to_excel(writer, sheet_name="Lançamentos", index=False)
                if not df_res.empty:
                    df_res.to_excel(writer, sheet_name="Resumo_IRT", index=False)

                # IRT summary sheet
                irt_export = df_res.copy() if not df_res.empty else pd.DataFrame()
                if not irt_export.empty:
                    irt_export.columns = ["Docente","Mês Ref.","Total Bruto (Kz)","IRT 6,5% (Kz)","Líquido (Kz)"]
                    irt_export.to_excel(writer, sheet_name="Resumo_IRT_formatado", index=False)

            buf.seek(0)
            st.download_button(
                label="📊 Descarregar Excel (Lançamentos + Resumo IRT)",
                data=buf,
                file_name=f"ISPTLO_Juris_Export_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # JSON backup
            json_data = json.dumps({
                "lancamentos": st.session_state.lancamentos,
                "pagamentos": st.session_state.pagamentos,
                "exportado_em": TODAY
            }, ensure_ascii=False, indent=2)
            st.download_button(
                label="💾 Descarregar Backup JSON",
                data=json_data.encode("utf-8"),
                file_name=f"ISPTLO_backup_{datetime.date.today().strftime('%Y%m%d')}.json",
                mime="application/json"
            )

            st.markdown("""
            <div class="alert-legal">
            📋 <strong>Instrução de uso do Excel exportado:</strong><br>
            • A folha <em>Lançamentos</em> contém o detalhe por sessão (sem IRT).<br>
            • A folha <em>Resumo_IRT</em> contém o IRT calculado sobre o total mensal por docente.<br>
            • Importe este ficheiro no Mapa de Controlo v3.0 para manter o histórico completo.
            </div>""", unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<small style='color:#888'>ISPTLO — Mapa de Controlo de Júri de TFC v3.0 | "
    "D.P. nº 191/18 (ECDES) | C.I.R.T. Lei nº 19/14 | Ofício 009/25 + Emenda DASG 023/2025</small>",
    unsafe_allow_html=True
)
