"""
═══════════════════════════════════════════════════════════════════════════════
ISPTLO - JURIS APP | BLOCO DE GERAÇÃO DE PDF (VERSÃO CORRIGIDA)
Ficheiro: pdf_block.py

PROBLEMA RESOLVIDO:
    O st.download_button() dispara um rerun() do Streamlit antes de executar
    o callback. Isso apaga todas as variáveis locais calculadas (total_bruto,
    df_filtrado, etc.), resultando em PDFs em branco ou NameError.

SOLUÇÃO APLICADA:
    Padrão "Calcular -> Persistir -> Descarregar" em dois momentos:
      1. MOMENTO DO CÁLCULO: os dados são guardados em st.session_state.
      2. MOMENTO DO DOWNLOAD: o PDF é gerado lendo de st.session_state
         (os dados estão sempre disponíveis, mesmo após o rerun).

INSTALAÇÃO:
    pip install fpdf2 pandas streamlit

UTILIZAÇÃO:
    Substitui o bloco antigo de PDF no módulo Docente E no módulo Tesouraria.
    Chame as funções abaixo nos pontos indicados no código original.
═══════════════════════════════════════════════════════════════════════════════
"""

# ─── IMPORTS NECESSÁRIOS (adicionar ao topo do isptlo_juris_app.py) ───────────
import streamlit as st
import pandas as pd
from fpdf import FPDF
import datetime
import io


# ═══════════════════════════════════════════════════════════════════════════════
# 1. FUNÇÃO DE FORMATAÇÃO (padrão contabilístico angolano)
# ═══════════════════════════════════════════════════════════════════════════════

def format_kz(value: float) -> str:
    """
    Formata um valor numérico para o padrão contabilístico angolano.
    Exemplo: 167500.50 -> '167.500,50 Kz'
    """
    if value is None:
        return "0,00 Kz"
    # Format with 2 decimal places, then swap separators
    formatted = f"{value:,.2f}"                  # '167,500.50'
    formatted = formatted.replace(",", "X")       # '167X500.50'
    formatted = formatted.replace(".", ",")       # '167X500,50'
    formatted = formatted.replace("X", ".")       # '167.500,50'
    return f"{formatted} Kz"


# ═══════════════════════════════════════════════════════════════════════════════
# 2. CLASSE PDF (fpdf2)
# ═══════════════════════════════════════════════════════════════════════════════

class ISPTLOReportPDF(FPDF):
    """
    Classe de geração de PDF para recibos e relatórios do ISPTLO.
    Herda de FPDF para permitir cabeçalho e rodapé automáticos.
    """

    @staticmethod
    def _resolve_font():
        """
        Resolve o caminho da fonte DejaVu de forma portátil.
        Tenta: (1) pasta do repo, (2) /tmp/ com download, (3) None (fallback Helvetica).
        """
        import os
        candidates = [
            os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf"),
            os.path.join(os.path.dirname(__file__), "fonts", "DejaVuSans.ttf"),
            "/tmp/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
        for path in candidates:
            if os.path.isfile(path):
                return path
        # Último recurso: tentar download (funciona no Streamlit Cloud)
        try:
            import urllib.request
            dest = "/tmp/DejaVuSans.ttf"
            url  = "https://github.com/dejavu-fonts/dejavu-fonts/releases/download/version_2_37/dejavu-fonts-ttf-2.37.zip"
            # Alternativa directa via sourceforge
            url2 = "https://sourceforge.net/projects/dejavu/files/dejavu/2.37/dejavu-fonts-ttf-2.37.tar.bz2"
            # Usar fonts.gstatic.com (Google) como CDN fiável
            url3 = "https://fonts.gstatic.com/s/dejavusans/v1/TEX-0-DejaVu.ttf"
            urllib.request.urlretrieve(url3, dest)
            if os.path.isfile(dest) and os.path.getsize(dest) > 10000:
                return dest
        except Exception:
            pass
        return None  # Cair para Helvetica

    def _get_font_name(self):
        return getattr(self, "_font_family", "Helvetica")

    def _safe(self, text):
        """Converte texto para latin-1 seguro (usado quando não há fonte Unicode)."""
        result = ""
        for ch in str(text):
            try:
                ch.encode("latin-1")
                result += ch
            except UnicodeEncodeError:
                # Mapeamento manual para caracteres portugueses comuns
                mapping = {
                    "ã": "a", "õ": "o", "â": "a", "ê": "e", "ô": "o",
                    "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
                    "à": "a", "è": "e", "ì": "i", "ò": "o", "ù": "u",
                    "ç": "c", "ñ": "n", "ü": "u",
                    "Ã": "A", "Õ": "O", "Â": "A", "Ê": "E", "Ô": "O",
                    "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
                    "À": "A", "È": "E", "Ì": "I", "Ò": "O", "Ù": "U",
                    "Ç": "C", "Ñ": "N", "Ü": "U",
                    "—": "-", "–": "-", "’": "'",
                    "“": '"', "”": '"',
                }
                result += mapping.get(ch, "?")
        return result

    def _c(self, txt):
        """Shorthand: sanitize text before placing in a cell (Helvetica-safe)."""
        if self._get_font_name() == "Helvetica":
            return self._safe(str(txt))
        return str(txt)

    def header(self):
        # Registar fonte Unicode na primeira chamada
        if not hasattr(self, "_fonts_registered"):
            font_path = self._resolve_font()
            if font_path:
                try:
                    bold_path = font_path.replace("DejaVuSans.ttf", "DejaVuSans-Bold.ttf")
                    obli_path = font_path.replace("DejaVuSans.ttf", "DejaVuSans-Oblique.ttf")
                    self.add_font("DejaVu", "",   font_path)
                    if os.path.isfile(bold_path):
                        self.add_font("DejaVu", "B", bold_path)
                    else:
                        self.add_font("DejaVu", "B", font_path)
                    if os.path.isfile(obli_path):
                        self.add_font("DejaVu", "I", obli_path)
                    else:
                        self.add_font("DejaVu", "I", font_path)
                    self._font_family = "DejaVu"
                except Exception:
                    self._font_family = "Helvetica"
            else:
                self._font_family = "Helvetica"
            self._fonts_registered = True
        font = self._get_font_name()
        # Faixa de cabeçalho azul escuro
        self.set_fill_color(31, 56, 100)        # #1F3864
        self.rect(0, 0, 210, 28, "F")
        # Título principal
        self.set_font(font, "B", 13)
        self.set_text_color(255, 215, 0)         # Gold
        self.set_xy(10, 6)
        self.cell(190, 7, "REPÚBLICA DE ANGOLA - ISPTLO", align="C")
        # Subtítulo
        self.set_font(self._get_font_name(), "", 9)
        self.set_text_color(204, 221, 255)       # Light blue
        self.set_xy(10, 14)
        self.cell(190, 6, "Instituto Superior Politécnico do Libolo | Mapa de Controlo de Júri de TFC", align="C")
        self.set_xy(10, 20)
        self.cell(190, 5, f"Gerado em: {datetime.date.today().strftime('%d/%m/%Y')} | v3.0", align="C")
        self.ln(10)

    def footer(self):
        self.set_y(-14)
        self.set_draw_color(31, 56, 100)
        self.set_line_width(0.4)
        self.line(10, self.get_y(), 200, self.get_y())
        self.set_font(self._get_font_name(), "I", 7)
        self.set_text_color(120, 120, 120)
        self.cell(
            0, 8,
            "D.P. nº 191/18 (ECDES) | C.I.R.T. Lei nº 19/14 Art.º 67 | "
            f"Página {self.page_no()}/{{nb}}",
            align="C"
        )

    def section_title(self, title: str):
        """Faixa de secção com título."""
        self.set_fill_color(46, 80, 136)         # #2E5088
        self.set_text_color(255, 255, 255)
        self.set_font(self._get_font_name(), "B", 10)
        self.set_x(10)
        self.cell(190, 8, self._c(f"  {title}"), fill=True, ln=True)
        self.ln(2)

    def meta_row(self, label: str, value: str, highlight: bool = False):
        """Par chave-valor para metadados do documento."""
        self.set_font(self._get_font_name(), "B", 9)
        self.set_text_color(46, 80, 136)
        self.set_x(10)
        self.cell(55, 6, self._c(label))
        self.set_font(self._get_font_name(), "", 9)
        self.set_text_color(0, 0, 0)
        if highlight:
            self.set_fill_color(198, 239, 206)   # Green fill
            self.cell(135, 6, self._c(value), fill=True, ln=True)
        else:
            self.cell(135, 6, self._c(value), ln=True)

    def kpi_block(
        self,
        total_bruto: float,
        total_irt: float,
        total_liquido: float
    ):
        """Bloco de 3 KPIs lado a lado."""
        self.ln(3)
        labels  = ["TOTAL BRUTO MENSAL", "RETENCAO IRT (6,5%) - AGT", "VALOR LIQUIDO A PAGAR"]
        values  = [total_bruto,              total_irt,                    total_liquido]
        r_fills = [(31, 56, 100),            (192, 0, 0),                  (30, 113, 69)]
        r_texts = [(255, 215, 0),            (255, 255, 255),              (255, 255, 255)]
        box_w   = 60
        x_start = 10
        for i, (lbl, val, fill, text) in enumerate(zip(labels, values, r_fills, r_texts)):
            x = x_start + i * (box_w + 5)
            # Label bar
            self.set_fill_color(*fill)
            self.set_text_color(*text)
            self.set_font(self._get_font_name(), "B", 7)
            self.set_xy(x, self.get_y())
            self.cell(box_w, 6, self._c(lbl), fill=True, align="C")
            # Value bar
            self.set_fill_color(240, 240, 240)
            self.set_text_color(0, 0, 0)
            self.set_font(self._get_font_name(), "B", 10)
            self.set_xy(x, self.get_y() + 6)
            self.cell(box_w, 9, self._c(format_kz(val)), fill=True, align="C")
        self.ln(20)

    def data_table(self, df: pd.DataFrame, col_widths: list, col_headers: list):
        """
        Tabela de dados genérica com cabeçalho escuro e linhas zebradas.
        col_widths: lista de larguras em mm para cada coluna.
        col_headers: lista de strings para o cabeçalho.
        """
        # Header row
        self.set_fill_color(31, 56, 100)
        self.set_text_color(255, 255, 255)
        self.set_font(self._get_font_name(), "B", 8)
        self.set_x(10)
        for header, width in zip(col_headers, col_widths):
            self.cell(width, 7, self._c(header), border=0, fill=True, align="C")
        self.ln()

        # Data rows
        self.set_text_color(0, 0, 0)
        self.set_font(self._get_font_name(), "", 8)
        for row_idx, (_, row) in enumerate(df.iterrows()):
            # Zebra fill
            if row_idx % 2 == 0:
                self.set_fill_color(235, 240, 250)
            else:
                self.set_fill_color(255, 255, 255)

            # Page break check
            if self.get_y() > 265:
                self.add_page()

            self.set_x(10)
            for col_idx, (col_name, width) in enumerate(zip(df.columns, col_widths)):
                cell_val = str(row[col_name]) if row[col_name] is not None else "-"
                # Right-align monetary columns (last 3 by convention)
                align = "R" if col_idx >= len(df.columns) - 3 else "L"
                self.cell(width, 6, self._c(cell_val), border=0, fill=True, align=align)
            self.ln()

        # Total bar
        self.set_fill_color(31, 56, 100)
        self.set_text_color(255, 255, 255)
        self.set_font(self._get_font_name(), "B", 8)
        self.set_x(10)
        self.ln(1)

    def legal_note(self, text: str, color: tuple = (133, 79, 11)):
        """Caixa de nota legal com borda lateral."""
        self.set_fill_color(255, 242, 204)
        self.set_draw_color(*color)
        self.set_line_width(0.8)
        x = self.get_x()
        y = self.get_y()
        self.line(10, y, 10, y + 14)
        self.set_fill_color(255, 242, 204)
        self.rect(10, y, 190, 14, "F")
        self.set_font(self._get_font_name(), "B", 7)
        self.set_text_color(*color)
        self.set_xy(13, y + 2)
        self.multi_cell(185, 4.5, self._c(text))
        self.ln(3)


# ═══════════════════════════════════════════════════════════════════════════════
# 3. FUNÇÕES DE GERAÇÃO DE PDF
# ═══════════════════════════════════════════════════════════════════════════════

def generate_docente_pdf(
    docente_name: str,
    mes_ref: str,
    participacoes: list,
    total_bruto: float,
    total_irt: float,
    total_liquido: float
) -> bytes:
    """
    Gera o recibo individual de um docente.
    Retorna os bytes do PDF - nunca faz rerun do Streamlit.
    """
    pdf = ISPTLOReportPDF(orientation="P", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=16)

    # ── Metadados do docente ─────────────────────────────────────────────────
    pdf.section_title("DADOS DO DOCENTE E PERÍODO")
    pdf.meta_row("Docente:", docente_name)
    pdf.meta_row("Período de Referência:", mes_ref)
    pdf.meta_row("Data de Emissão:", datetime.date.today().strftime("%d/%m/%Y"))
    pdf.meta_row("Base Legal:", "D.P. nº 191/18 (ECDES) | C.I.R.T. Lei nº 19/14 Art.º 67")
    pdf.ln(4)

    # ── KPIs ─────────────────────────────────────────────────────────────────
    pdf.section_title("RESUMO FINANCEIRO DO PERÍODO")
    pdf.kpi_block(total_bruto, total_irt, total_liquido)

    # ── Tabela de participações ───────────────────────────────────────────────
    pdf.section_title("DETALHE DAS PARTICIPAÇÕES NO MÊS")

    if participacoes:
        df_part = pd.DataFrame(participacoes)
        # Format bruto column
        if "bruto" in df_part.columns:
            df_part["bruto_fmt"] = df_part["bruto"].apply(format_kz)
        # Select display columns
        display_cols  = []
        display_hdrs  = []
        col_widths_map = {}

        col_candidates = [
            ("funcao",    "Função",        50),
            ("tipo",      "Tipo",          22),
            ("data",      "Data",          24),
            ("despacho",  "Despacho",      40),
            ("bruto_fmt", "Valor Bruto",   40),
        ]
        for col, hdr_lbl, width in col_candidates:
            if col in df_part.columns:
                display_cols.append(col)
                display_hdrs.append(hdr_lbl)
                col_widths_map[col] = width

        if display_cols:
            pdf.data_table(
                df_part[display_cols],
                [col_widths_map[c] for c in display_cols],
                display_hdrs
            )

    # ── Cálculo detalhado do IRT ──────────────────────────────────────────────
    pdf.section_title("CÁLCULO DE RETENÇÃO NA FONTE (IRT)")
    pdf.set_font(pdf._get_font_name(), "", 9)
    pdf.set_text_color(0, 0, 0)
    pdf.set_x(10)
    pdf.multi_cell(
        190, 5.5,
        f"Total Bruto Mensal Acumulado: {format_kz(total_bruto)}\n"
        f"Retencao IRT (6.5%): {format_kz(total_irt)}\n"
        f"Valor Liquido a Transferir: {format_kz(total_liquido)}"
    )
    pdf.ln(2)

    # ── Nota legal ────────────────────────────────────────────────────────────
    pdf.legal_note(
        "NOTA FISCAL: A retenção de IRT de 6,5% incide sobre o TOTAL MENSAL acumulado por docente, "
        "não por sessão individual. Base legal: C.I.R.T. (Lei nº 19/14) Art.º 67. "
        "O ISPTLO é responsável tributário pela retenção e entrega à AGT até ao dia 20 do mês seguinte."
    )

    # ── Assinaturas ───────────────────────────────────────────────────────────
    pdf.ln(8)
    pdf.set_font(pdf._get_font_name(), "", 9)
    pdf.set_text_color(0, 0, 0)
    sig_y = pdf.get_y()
    # Left: docente
    pdf.set_xy(15, sig_y)
    pdf.cell(80, 5, "_" * 35, align="C")
    pdf.set_xy(15, sig_y + 6)
    pdf.cell(80, 5, docente_name, align="C")
    pdf.set_xy(15, sig_y + 11)
    pdf.cell(80, 5, "Docente", align="C")
    # Right: tesouraria
    pdf.set_xy(115, sig_y)
    pdf.cell(80, 5, "_" * 35, align="C")
    pdf.set_xy(115, sig_y + 6)
    pdf.cell(80, 5, "[Responsável da Tesouraria]", align="C")
    pdf.set_xy(115, sig_y + 11)
    pdf.cell(80, 5, "Departamento de Finanças - ISPTLO", align="C")

    return bytes(pdf.output())


def generate_tesouraria_pdf(
    mes_filtro: str,
    df_resumo: pd.DataFrame,
    total_bruto: float,
    total_irt: float,
    total_liquido: float
) -> bytes:
    """
    Gera o relatório executivo da Tesouraria com todos os docentes do período.
    """
    pdf = ISPTLOReportPDF(orientation="P", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=16)

    # ── Cabeçalho do relatório ───────────────────────────────────────────────
    pdf.section_title("RELATÓRIO EXECUTIVO - PAGAMENTOS DE JÚRI DE TFC")
    pdf.meta_row("Período de Referência:", mes_filtro if mes_filtro != "Todos" else "Todos os Meses")
    pdf.meta_row("Data de Emissão:", datetime.date.today().strftime("%d/%m/%Y"))
    pdf.meta_row("Classificação:", "CONFIDENCIAL - Uso Restrito - Presidência e Conselho de Direcção")
    pdf.ln(4)

    # ── KPIs globais ─────────────────────────────────────────────────────────
    pdf.section_title("TOTAIS CONSOLIDADOS")
    pdf.kpi_block(total_bruto, total_irt, total_liquido)

    # ── Nota AGT ─────────────────────────────────────────────────────────────
    pdf.legal_note(
        f"TOTAL IRT A ENTREGAR À AGT: {format_kz(total_irt)}  |  "
        "Prazo: até ao dia 20 do mês seguinte (C.I.R.T. Art.º 67).  |  "
        "Regra fiscal: IRT calculado sobre o TOTAL MENSAL por docente, não por sessão."
    )

    # ── Tabela por docente ────────────────────────────────────────────────────
    if not df_resumo.empty:
        pdf.section_title("DETALHE POR DOCENTE E MÊS - RESUMO IRT")
        df_pdf = df_resumo.copy()
        # Format monetary columns
        for col in ["total_bruto", "irt_6_5", "liquido"]:
            if col in df_pdf.columns:
                df_pdf[col] = df_pdf[col].apply(format_kz)

        col_map = {
            "docente":     ("Docente",               62),
            "mes_ref":     ("Mês Ref.",               22),
            "total_bruto": ("Total Bruto (Kz)",       40),
            "irt_6_5":     ("IRT 6,5% (Kz)",          34),
            "liquido":     ("Líquido a Pagar (Kz)",   32),
        }
        display_cols  = [c for c in col_map if c in df_pdf.columns]
        display_hdrs  = [col_map[c][0] for c in display_cols]
        display_widths= [col_map[c][1] for c in display_cols]

        pdf.data_table(df_pdf[display_cols], display_widths, display_hdrs)

    # ── Bloco de controlo documental ─────────────────────────────────────────
    pdf.ln(4)
    pdf.section_title("BLOCO DE CONTROLO DOCUMENTAL")
    pdf.meta_row("Elaborado por:",  "[Nome e Cargo - Tesouraria]")
    pdf.meta_row("Aprovado por:",   "[Nome, Grau e Cargo - Presidência]")
    pdf.meta_row("Acordo de Ref.:", "Acordo 1.3 - Acta da Comissão de Júris de 28/04/2026")
    pdf.meta_row("Ref. do Mapa:",   "ISPTLO_Mapa_Controlo_v3.0")

    return bytes(pdf.output())


# ═══════════════════════════════════════════════════════════════════════════════
# 4. BLOCO A INSERIR NO MÓDULO DOCENTE
#    Localização no código original:
#    -> logo após o cálculo de `total_bruto`, `irt`, `liquido` e a exibição das KPIs
# ═══════════════════════════════════════════════════════════════════════════════

def render_docente_pdf_block(
    docente_name: str,
    mes_ref: str,
    participacoes: list,
    total_bruto: float,
    total_irt: float,
    total_liquido: float
):
    """
    PADRÃO CORRECTO para Streamlit:
      MOMENTO 1: Persiste os dados no session_state quando o cálculo é feito.
      MOMENTO 2: Lê do session_state para gerar o PDF (após o rerun do download_button).

    Como usar:
        Chame esta função logo após calcular total_bruto, irt, liquido no módulo Docente.
        Passe as mesmas variáveis locais.
    """

    # ── MOMENTO 1: Persistir dados no session_state ──────────────────────────
    # Sempre que há novos dados calculados, actualiza o estado.
    # Isto garante que os dados existem DEPOIS do rerun disparado pelo download_button.
    st.session_state["pdf_docente_data"] = {
        "docente_name":  docente_name,
        "mes_ref":       mes_ref,
        "participacoes": participacoes,
        "total_bruto":   total_bruto,
        "total_irt":     total_irt,
        "total_liquido": total_liquido,
        "timestamp":     datetime.datetime.now().isoformat(),
    }

    # ── MOMENTO 2: Ler do estado e gerar PDF ─────────────────────────────────
    # O botão só aparece se houver dados persistidos.
    if "pdf_docente_data" in st.session_state and st.session_state["pdf_docente_data"]:
        cached = st.session_state["pdf_docente_data"]

        # Gerar PDF em memória - NÃO usa variáveis locais, usa o session_state
        pdf_bytes = generate_docente_pdf(
            docente_name  = cached["docente_name"],
            mes_ref       = cached["mes_ref"],
            participacoes = cached["participacoes"],
            total_bruto   = cached["total_bruto"],
            total_irt     = cached["total_irt"],
            total_liquido = cached["total_liquido"],
        )

        # Nome do ficheiro dinâmico com o nome do docente
        safe_name = cached["docente_name"].replace(" ", "_")
        filename  = f"Recibo_{safe_name}_{cached['mes_ref'].replace('/','-')}.pdf"

        st.download_button(
            label    = f"📄 Descarregar Recibo PDF - {cached['docente_name']}",
            data     = pdf_bytes,
            file_name= filename,
            mime     = "application/pdf",
            key      = "btn_pdf_docente",
            help     = "O recibo inclui o detalhe das participações e o cálculo correcto do IRT mensal."
        )


# ═══════════════════════════════════════════════════════════════════════════════
# 5. BLOCO A INSERIR NO MÓDULO TESOURARIA (Dashboard / Exportar)
#    Localização no código original:
#    -> na secção "📥 Exportar Excel / PDF", logo após a geração do Excel
# ═══════════════════════════════════════════════════════════════════════════════

def render_tesouraria_pdf_block(
    mes_filtro: str,
    df_resumo: pd.DataFrame,
    total_bruto: float,
    total_irt: float,
    total_liquido: float
):
    """
    Mesmo padrão persist-then-read para o módulo de Tesouraria.
    Gera um relatório executivo com todos os docentes do período.
    """

    # ── MOMENTO 1: Persistir ─────────────────────────────────────────────────
    # Converte o DataFrame para dict para poder serializar no session_state
    st.session_state["pdf_tesouraria_data"] = {
        "mes_filtro":    mes_filtro,
        "df_resumo":     df_resumo.to_dict("records") if not df_resumo.empty else [],
        "total_bruto":   total_bruto,
        "total_irt":     total_irt,
        "total_liquido": total_liquido,
        "timestamp":     datetime.datetime.now().isoformat(),
    }

    # ── MOMENTO 2: Ler do estado e gerar PDF ─────────────────────────────────
    if "pdf_tesouraria_data" in st.session_state and st.session_state["pdf_tesouraria_data"]:
        cached = st.session_state["pdf_tesouraria_data"]

        # Reconstituir DataFrame do session_state
        df_cached = (
            pd.DataFrame(cached["df_resumo"])
            if cached["df_resumo"]
            else pd.DataFrame()
        )

        pdf_bytes = generate_tesouraria_pdf(
            mes_filtro    = cached["mes_filtro"],
            df_resumo     = df_cached,
            total_bruto   = cached["total_bruto"],
            total_irt     = cached["total_irt"],
            total_liquido = cached["total_liquido"],
        )

        mes_safe = cached["mes_filtro"].replace("/", "-").replace(" ", "_")
        filename = f"ISPTLO_Relatorio_Juris_{mes_safe}_{datetime.date.today().strftime('%Y%m%d')}.pdf"

        st.download_button(
            label    = " Descarregar Relatório Executivo PDF",
            data     = pdf_bytes,
            file_name= filename,
            mime     = "application/pdf",
            key      = "btn_pdf_tesouraria",
            help     = "Relatório confidencial - Presidência e Conselho de Direcção."
        )

        st.markdown("""
        <div style='background:#C6EFCE;border-left:4px solid #1E7145;
                    padding:0.6rem 1rem;border-radius:6px;font-size:13px;color:#0A3D1F;margin-top:6px'>
         <strong>PDF gerado correctamente.</strong>
        O IRT foi calculado sobre o total mensal por docente (não por sessão)
        conforme exige o <strong>C.I.R.T. Art.º 67</strong>.
        </div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# 6. EXEMPLO DE INTEGRAÇÃO NO isptlo_juris_app.py
# ═══════════════════════════════════════════════════════════════════════════════
#
# ── NO MÓDULO DOCENTE (após calcular os valores e mostrar as KPIs) ─────────────
#
#   # ... cálculo existente ...
#   total_bruto = sum(p["bruto"] for p in participacoes)
#   irt         = round(total_bruto * IRT_RATE, 2)
#   liquido     = total_bruto - irt
#
#   # ... mostrar as KPIs com st.markdown(...) ...
#
#   #  SUBSTITUIR o bloco antigo de PDF por esta chamada:
#   from pdf_block import render_docente_pdf_block
#   render_docente_pdf_block(
#       docente_name  = docente_sel,        # variável local existente
#       mes_ref       = mes_sel,            # variável local existente
#       participacoes = participacoes,      # lista de dicts existente
#       total_bruto   = total_bruto,
#       total_irt     = irt,
#       total_liquido = liquido,
#   )
#
#
# ── NO MÓDULO TESOURARIA (secção "📥 Exportar Excel / PDF") ────────────────────
#
#   df_resumo   = get_resumo()
#   total_bruto = df_resumo["total_bruto"].sum() if not df_resumo.empty else 0
#   total_irt   = df_resumo["irt_6_5"].sum()    if not df_resumo.empty else 0
#   total_liq   = df_resumo["liquido"].sum()    if not df_resumo.empty else 0
#
#   #  SUBSTITUIR o bloco antigo de PDF por esta chamada:
#   from pdf_block import render_tesouraria_pdf_block
#   render_tesouraria_pdf_block(
#       mes_filtro    = mes_filtro,         # variável de filtro existente
#       df_resumo     = df_resumo,
#       total_bruto   = total_bruto,
#       total_irt     = total_irt,
#       total_liquido = total_liq,
#   )
#
# ══════════════════════════════════════════════════════════════════════════════
