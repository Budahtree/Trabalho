"""Gera o Modelo de Avaliação de Investimentos (Excel).

Uso:  python3 ferramentas/gerar_modelo.py [caminho_saida.xlsx]

O ficheiro gerado é 100% baseado em fórmulas: todos os resultados recalculam
quando se alteram os pressupostos. Os valores por defeito são os do projeto
Cocoon (caso de teste).
"""
import sys
from copy import copy
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment
from openpyxl.chart import LineChart, BarChart, PieChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.formatting.rule import DataBarRule
from openpyxl.workbook.properties import CalcProperties

OUT = sys.argv[1] if len(sys.argv) > 1 else "Modelo_Avaliacao_Investimentos.xlsx"

# ---------------------------------------------------------------- estilos
FONT = "Arial"
NAVY = "1B2A41"          # cor principal
NAVY2 = "2E4A72"         # banner secundário
BLUE = "2E75B6"          # acento
LIGHT = "EAF1FB"         # fundo de secção
GOLD = "C8A951"
GREEN = "2E7D32"
RED = "C62828"
GREY = "7F7F7F"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_TITLE = Font(name=FONT, size=16, bold=True, color="FFFFFF")
F_SUB = Font(name=FONT, size=9, italic=True, color=GREY)
F_HDR = Font(name=FONT, size=9, bold=True, color="FFFFFF")
F_SEC = Font(name=FONT, size=10, bold=True, color=NAVY)
F_NAV = Font(name=FONT, size=8, color=BLUE, underline="single")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor=LIGHT)
FILL_IN = PatternFill("solid", fgColor="FFF8E1")
FILL_KEY = PatternFill("solid", fgColor="FFE699")
FILL_TOT = PatternFill("solid", fgColor="F3F6FA")
FILL_OK = PatternFill("solid", fgColor="D8F0DA")
FILL_BAD = PatternFill("solid", fgColor="F9D6D5")
FILL_WARN = PatternFill("solid", fgColor="FFF0C2")
FILL_CARD = PatternFill("solid", fgColor="F3F6FA")
FILL_NAV = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="D0D7E1")
DOT = Side(style="hair", color="A6A6A6")
B_TOP = Border(top=Side(style="thin", color=NAVY))
B_TOT = Border(top=Side(style="thin", color=NAVY), bottom=Side(style="double", color=NAVY))
B_BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
B_IN = Border(left=DOT, right=DOT, top=DOT, bottom=DOT)
B_ROW = Border(bottom=Side(style="hair", color="D0D7E1"))

EUR = '#,##0;[Red](#,##0);"–"'
EUR2 = '#,##0.00;[Red](#,##0.00);"–"'
MEUR = '#,##0.00,," M€";[Red](#,##0.00,," M€");"–"'
PCT = '0.0%;[Red](0.0%);"–"'
MULT = '0.00"x";[Red](0.00"x");"–"'
INT = '0;[Red](0);"–"'
ANOS = '[=1]0" ano";0" anos"'
YEAR = '0'

wb = Workbook()
names = {}
NAV = ["Início", "Dashboard", "Estrutura", "Pressupostos", "Capex", "Pessoal", "FSE", "Relatório", "Checks"]


def name(nm, ws, ref):
    """Cria um nome definido ao nível do livro."""
    names[nm] = f"'{ws.title}'!{ref}"
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws.title}'!{ref}")


def style_range(ws, rng, font=None, fill=None, fmt=None, align=None, border=None):
    for row in ws[rng]:
        for c in row:
            if font: c.font = font
            if fill: c.fill = fill
            if fmt: c.number_format = fmt
            if align: c.alignment = align
            if border: c.border = border


def header_row(ws, row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 26


def section(ws, row, text, ncols, start_col=1):
    for col in range(start_col, start_col + ncols):
        c = ws.cell(row=row, column=col)
        c.fill = FILL_SEC
        c.border = Border(bottom=Side(style="medium", color=BLUE))
    c = ws.cell(row=row, column=start_col, value=text)
    c.font = F_SEC
    ws.row_dimensions[row].height = 20


def set_input(c, key=False):
    c.font = F_INPUT
    c.fill = FILL_KEY if key else FILL_IN
    c.border = B_IN


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


def merge(ws, r1, c1, r2, c2, value=None, font=None, fill=None, fmt=None, align=None, border=None):
    """Junta células e aplica estilo a todo o intervalo (bordas incluídas)."""
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    for r_ in range(r1, r2 + 1):
        for c_ in range(c1, c2 + 1):
            cell = ws.cell(row=r_, column=c_)
            if fill: cell.fill = fill
            if border: cell.border = border
    c = ws.cell(row=r1, column=c1)
    if value is not None: c.value = value
    if font: c.font = font
    if fmt: c.number_format = fmt
    if align: c.alignment = align
    return c


def outline_box(ws, r1, c1, r2, c2, color, style="medium"):
    sd = Side(style=style, color=color)
    for r_ in range(r1, r2 + 1):
        for c_ in range(c1, c2 + 1):
            cell = ws.cell(row=r_, column=c_)
            b = cell.border
            cell.border = Border(left=sd if c_ == c1 else b.left, right=sd if c_ == c2 else b.right,
                                 top=sd if r_ == r1 else b.top, bottom=sd if r_ == r2 else b.bottom)


def banner(ws, title, subtitle, ncols, nav=True, nav_span=1, start_col=1):
    """Banner superior (título + subtítulo) e barra de navegação."""
    fill1 = PatternFill("solid", fgColor=NAVY)
    fill2 = PatternFill("solid", fgColor=NAVY2)
    for col in range(1, start_col + ncols):
        ws.cell(row=1, column=col).fill = fill1
        ws.cell(row=2, column=col).fill = fill2
    c = ws.cell(row=1, column=start_col, value=title)
    c.font = F_TITLE
    c.alignment = Alignment(vertical="center")
    c = ws.cell(row=2, column=start_col, value=subtitle)
    c.font = Font(name=FONT, size=9, italic=True, color="DCE6F2")
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 18
    if not nav:
        return
    for col in range(1, start_col + ncols):
        ws.cell(row=3, column=col).fill = FILL_NAV
    col = start_col
    for sh in NAV:
        if sh == ws.title:
            c = merge(ws, 3, col, 3, col + nav_span - 1, value=sh,
                      font=Font(name=FONT, size=8, bold=True, color=NAVY), align=Alignment(horizontal="center"))
        else:
            c = merge(ws, 3, col, 3, col + nav_span - 1, value=sh, font=F_NAV, align=Alignment(horizontal="center"))
            c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{sh}'!A1", display=sh)
        col += nav_span
    ws.row_dimensions[3].height = 16


def style_chart(ch, colors, kind="line", yfmt='#,##0.0,,"M"'):
    for s_, col in zip(ch.series, colors):
        if kind == "line":
            s_.graphicalProperties.line.solidFill = col
            s_.graphicalProperties.line.width = 28575
            s_.smooth = False
        else:
            s_.graphicalProperties.solidFill = col
            s_.graphicalProperties.line.solidFill = col
    ch.legend.position = "b"
    ch.y_axis.numFmt = yfmt
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E7EE"))
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.tickLblPos = "low"


# ================================================================ INÍCIO
ws0 = wb.active
ws0.title = "Início"
ws_dash = wb.create_sheet("Dashboard")
ws_est = wb.create_sheet("Estrutura")
ws_in = wb.create_sheet("Pressupostos")
ws_cx = wb.create_sheet("Capex")
ws_p = wb.create_sheet("Pessoal")
ws_f = wb.create_sheet("FSE")
ws_rep = wb.create_sheet("Relatório")
ws_chk = wb.create_sheet("Checks")
SCEN = ["Base", "Pessimista", "Otimista"]
ws_eng = [wb.create_sheet(f"M_{s}") for s in SCEN]

# ================================================================ PRESSUPOSTOS
ws = ws_in
banner(ws, "Pressupostos do projeto",
       "Preencha apenas as células a azul (fundo amarelo). Amarelo forte = pressuposto-chave a validar. "
       "Valores por defeito: projeto Cocoon (caso de teste).", 7)
widths(ws, {"A": 46, "B": 3, "C": 16, "D": 16, "E": 16, "F": 16, "G": 70})

r = 4


def inp(label, nm, value, fmt=EUR, unit="", note="", key=False, valid=None):
    """Input de valor único na coluna C."""
    global r
    ws.cell(row=r, column=1, value=label).font = F_BASE
    c = ws.cell(row=r, column=3, value=value)
    set_input(c, key)
    c.number_format = fmt
    if fmt in ("@",):
        c.alignment = Alignment(horizontal="left")
    ws.cell(row=r, column=6, value=unit).font = F_SUB
    ws.cell(row=r, column=7, value=note).font = F_SUB
    if valid is not None:
        valid.add(c)
    name(nm, ws, f"$C${r}")
    r += 1


def inp_table(label_fmt, nm, values, fmt=EUR, unit="", note=""):
    """Tabela vertical (ex.: ADR por ano) com nome para o intervalo."""
    global r
    r0 = r
    for i, v in enumerate(values):
        ws.cell(row=r, column=1, value=label_fmt.format(i + 1)).font = F_BASE
        c = ws.cell(row=r, column=3, value=v)
        set_input(c)
        c.number_format = fmt
        ws.cell(row=r, column=6, value=unit).font = F_SUB
        if i == 0:
            ws.cell(row=r, column=7, value=note).font = F_SUB
        r += 1
    name(nm, ws, f"$C${r0}:$C${r - 1}")


def sec(text):
    global r
    r += 1
    section(ws, r, text, 7)
    r += 1


dv_simnao = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=False)
dv_amort = DataValidation(type="list", formula1='"Francês,Linear"', allow_blank=False)
dv_regime = DataValidation(type="list", formula1='"Simples,Composto"', allow_blank=False)
dv_tipo = DataValidation(type="list", formula1='"Hotel / Turismo,Turismo + Imobiliário,Imobiliário (venda),Outro"',
                         allow_blank=True)
for dv in (dv_simnao, dv_amort, dv_regime, dv_tipo):
    ws.add_data_validation(dv)

sec("1. IDENTIFICAÇÃO E CALENDÁRIO")
inp("Nome do projeto", "NomeProjeto", "Cocoon", fmt="@")
inp("Localização", "Localizacao", "Comporta / Alcácer do Sal", fmt="@")
inp("Tipo de ativo", "TipoAtivo", "Turismo + Imobiliário", fmt="@", valid=dv_tipo)
inp("Analista / responsável", "Analista", "Potentia Group", fmt="@")
inp("Ano civil do período 0", "AnoInicio", 2026, fmt=YEAR, unit="ano",
    note="Período 0 = ano do investimento inicial (aquisição / obra).")
inp("Período de abertura (1.º ano de operação)", "AnoAbertura", 1, fmt=INT, unit="período",
    note="Ex.: 1 = abre no ano seguinte ao investimento; 2 = obra dura 2 anos.", key=True)
inp("Período de saída / horizonte de análise", "AnoSaida", 15, fmt=INT, unit="período",
    note="Máximo 20. No período de saída considera-se a venda do ativo (valor terminal).", key=True)

# ---- cenários
sec("2. CENÁRIOS — VARIÁVEIS-CHAVE (Base / Pessimista / Otimista)")
for i, s in enumerate(SCEN):
    c = ws.cell(row=r, column=3 + i, value=s)
    c.font = F_HDR
    c.fill = FILL_HDR
    c.alignment = Alignment(horizontal="center")
name("ListaCenarios", ws, f"$C${r}:$E${r}")
ws.cell(row=r, column=1, value="Variável").font = F_BOLD
ws.cell(row=r, column=7, value="Notas").font = F_BOLD
r += 1
scen_inputs = [
    ("Ocupação — 1.º ano de operação", "S_Ocup1", (0.40, 0.32, 0.45), PCT, "%",
     "Dossier bancário indicava 30%; modelo original usava 40%. A validar."),
    ("Crescimento anual da ocupação", "S_OcupG", (0.05, 0.04, 0.05), PCT, "p.p./ano", ""),
    ("Ocupação máxima (estabilizada)", "S_OcupMax", (0.70, 0.62, 0.75), PCT, "%", ""),
    ("Ajuste ao ADR (sobre a tabela de ADR)", "S_AjADR", (0.0, -0.10, 0.05), PCT, "%", "0% = tabela tal como está."),
    ("Ajuste ao Capex (rubricas marcadas 'Sim')", "S_AjCapex", (0.0, 0.10, -0.05), PCT, "%",
     "Derrapagem / poupança de obra."),
    ("Ajuste aos custos operacionais", "S_AjCustos", (0.0, 0.05, -0.03), PCT, "%",
     "Aplica-se a payroll, FSE e custos diretos por noite."),
    ("Ajuste ao preço de venda das unidades", "S_AjPreco", (0.0, -0.10, 0.05), PCT, "%", ""),
    ("Exit yield (NOI / valor de saída)", "S_ExitYield", (0.075, 0.085, 0.07), PCT, "%",
     "Modelo original tinha 0% (sem valor terminal). Valor indicativo — validar com avaliador."),
]
for label, nm, vals, fmt, unit, note in scen_inputs:
    ws.cell(row=r, column=1, value=label).font = F_BASE
    for i, v in enumerate(vals):
        c = ws.cell(row=r, column=3 + i, value=v)
        set_input(c, key=(i == 0))
        c.number_format = fmt
    ws.cell(row=r, column=6, value=unit).font = F_SUB
    ws.cell(row=r, column=7, value=note).font = F_SUB
    name(nm, ws, f"$C${r}:$E${r}")
    r += 1

sec("3. OPERAÇÃO — OFERTA E RECEITAS")
inp("Unidades / quartos (chaves)", "Unidades", 48, fmt=INT, unit="unid.", key=True)
inp("Dias de operação por ano", "Dias", 365, fmt=INT, unit="dias")
inp_table("ADR — ano de operação {}", "ADR_tab", [240, 280, 330, 400, 0], fmt=EUR, unit="€ / noite",
          note="Preço médio por quarto ocupado. Deixe 0 para usar o crescimento abaixo a partir do ano anterior.")
inp("Crescimento anual do ADR (após a tabela)", "ADR_g", 0.015, fmt=PCT, unit="% / ano")
inp("Receita F&B por quarto ocupado", "FB_quarto", 62, fmt=EUR, unit="€ / noite")
inp("Crescimento anual da receita F&B por quarto", "FB_g", 0.0, fmt=PCT, unit="% / ano")
inp("Outros proveitos", "Outros_pct", 0.03, fmt=PCT, unit="% rec. quartos")

sec("4. OPERAÇÃO — CUSTOS DIRETOS")
inp("Amenities e consumíveis", "C_Amen", 4, fmt=EUR2, unit="€ / noite ocup.")
inp("Lavandaria e roupa", "C_Lav", 5, fmt=EUR2, unit="€ / noite ocup.")
inp("Inflação dos custos por noite", "InflDrivers", 0.03, fmt=PCT, unit="% / ano")
inp("Peso das OTA na receita de quartos", "OTA_peso", 0.35, fmt=PCT, unit="%")
inp("Comissão média das OTA", "OTA_com", 0.17, fmt=PCT, unit="% rec. OTA")
inp("Taxas de pagamento", "TaxasPag", 0.01, fmt=PCT, unit="% rec. total")
inp("Peso de Food na receita F&B", "Food_peso", 0.60, fmt=PCT, unit="%")
inp("Food COGS", "Food_cogs", 0.30, fmt=PCT, unit="% rec. Food")
inp("Beverage COGS", "Bev_cogs", 0.20, fmt=PCT, unit="% rec. Bev.")
inp("Consumíveis operacionais F&B", "FB_cons", 0.02, fmt=PCT, unit="% rec. F&B")
inp("Custo direto de outros proveitos", "CustoOutros_pct", 0.60, fmt=PCT, unit="% outros prov.")

sec("5. OPERAÇÃO — PESSOAL, FSE E GESTÃO")
inp("Inflação salarial anual", "InflSal", 0.03, fmt=PCT, unit="% / ano",
    note="O quadro de pessoal está na folha Pessoal.")
inp("Inflação anual dos FSE", "InflFSE", 0.0, fmt=PCT, unit="% / ano",
    note="Modelo original: 0%. Recomenda-se 2% — a validar.", key=True)
inp("Ocupação de referência dos FSE variáveis", "OcupRef", 1.0, fmt=PCT, unit="%",
    note="A parte variável dos FSE escala com ocupação / ocupação de referência. A rubrica está na folha FSE.")
inp("Management fee (base)", "MgmtFee", 0.05, fmt=PCT, unit="% rec. total")
inp("Incentive fee", "IncFee", 0.0, fmt=PCT, unit="% GOP")
inp("Crescimento anual dos shared services", "SS_g", 0.03, fmt=PCT, unit="% / ano",
    note="Custo base calculado na folha Pessoal (alocação por chave).")
inp("Reserva FF&E", "FFE_res", 0.04, fmt=PCT, unit="% rec. total")
inp("Reserva FF&E a partir do ano de operação", "FFE_inicio", 2, fmt=INT, unit="ano op.")
inp("CAPEX de manutenção adicional (ano 1)", "CapexManut", 0, fmt=EUR, unit="€ / ano")
inp("Crescimento do CAPEX de manutenção", "CapexManut_g", 0.03, fmt=PCT, unit="% / ano")

sec("6. VENDA DE UNIDADES (componente imobiliária — deixe 0 se não aplicável)")
inp_table("Unidades vendidas — ano de operação {}", "Vendas_tab", [3, 2, 1, 0, 0], fmt=INT, unit="unid.",
          note="Vendas ocorrem durante o ano indicado.")
inp("As unidades vendidas saem da exploração?", "Vendas_reduz", "Sim", fmt="@", valid=dv_simnao,
    note="Sim = a partir do ano seguinte à venda deixam de gerar receita hoteleira.")
inp("Preço de venda por unidade", "PrecoUnid", 2000000, fmt=EUR, unit="€ / unid.", key=True)
inp("Valorização anual do preço de venda", "PrecoUnid_g", 0.0, fmt=PCT, unit="% / ano")
inp("Custos de venda / comissões", "CustoVenda_pct", 0.05, fmt=PCT, unit="% preço",
    note="O modelo original aplicava 95% ao preço (5% de custos).")
inp("Custo contabilístico por unidade vendida", "CustoContab_unid", 345833, fmt=EUR, unit="€ / unid.",
    note="Para apurar a mais-valia tributável. Defeito: (aquisição + construção) / 48 unidades.")

sec("7. FISCALIDADE E AMORTIZAÇÕES")
inp("Taxa de IRC (incl. derrama)", "IRC", 0.20, fmt=PCT, unit="%")
inp("Limite de dedução de prejuízos fiscais", "LimitePrej", 0.65, fmt=PCT, unit="% lucro trib.",
    note="Regime PT: dedução limitada a 65% do lucro tributável; reporte sem prazo.")
inp("Vida útil — construção", "VidaConstr", 20, fmt=INT, unit="anos")
inp("Vida útil — FF&E", "VidaFFE", 5, fmt=INT, unit="anos")
inp("Vida útil — gastos amortizáveis (pré-abertura, etc.)", "VidaOutros", 3, fmt=INT, unit="anos")
inp("Imposto sobre a mais-valia na saída", "TaxaMaisValia", 0.20, fmt=PCT, unit="%",
    note="Aplicado a (valor de saída líquido − valor contabilístico). 0% se não aplicável.")

sec("8. FINANCIAMENTO")
inp("Dívida bancária — montante", "Divida", 6020000, fmt=EUR, unit="€", key=True,
    note="Cocoon: 70% × 8,6 M€ de aquisição (organigrama).")
inp("Período de desembolso", "DividaPer", 0, fmt=INT, unit="período")
inp("Taxa de juro anual", "Juro", 0.05, fmt=PCT, unit="%", key=True)
inp("Imposto do selo sobre juros", "Selo", 0.04, fmt=PCT, unit="% juros",
    note="O modelo original aplicava ×1,04 aos juros.")
inp("Prazo total (incl. carência)", "Prazo", 10, fmt=INT, unit="anos")
inp("Carência de capital", "Carencia", 3, fmt=INT, unit="anos", note="Durante a carência pagam-se só juros.")
inp("Tipo de amortização", "TipoAmort", "Francês", fmt="@", valid=dv_amort,
    note="Francês = prestação constante; Linear = capital constante.")
inp("Comissão de abertura / estruturação", "ComAbertura", 0.0, fmt=PCT, unit="% dívida")
inp("Subsídio a fundo perdido — montante", "Subsidio", 4610000, fmt=EUR, unit="€",
    note="Portugal 2030. Considerado não tributável (simplificação).")
inp("Subsídio — período de recebimento", "SubsidioPer", 1, fmt=INT, unit="período",
    note="Até lá o promotor tem de adiantar o valor (fica refletido no cash flow).")
inp("Investidor financeiro — nome", "InvNome", "Portugal Ventures", fmt="@")
inp("Investidor financeiro — montante investido", "InvMontante", 3773000, fmt=EUR, unit="€",
    note="Cocoon: 70% × (10 M€ obra − 4,61 M€ P2030). 0 se não existir.")
inp("Investidor financeiro — período de entrada", "InvPer", 0, fmt=INT, unit="período")
inp("Investidor financeiro — remuneração anual", "InvTaxa", 0.08, fmt=PCT, unit="%")
inp("Investidor financeiro — regime de juro", "InvRegime", "Simples", fmt="@", valid=dv_regime)
inp("Investidor financeiro — recompra após (anos)", "InvRecompra", 7, fmt=INT, unit="anos",
    note="Recompra paga pelo projeto/promotor. Se ultrapassar a saída, é paga na saída.")

sec("9. SAÍDA E AVALIAÇÃO")
inp("Custos de transação na saída", "CustosSaida_pct", 0.02, fmt=PCT, unit="% valor")
inp("Taxa de desconto — projeto (WACC)", "TaxaDescProj", 0.08, fmt=PCT, unit="%", key=True)
inp("Taxa de desconto — capital próprio", "CustoEquity", 0.12, fmt=PCT, unit="%", key=True)

sec("10. CRITÉRIOS DE VIABILIDADE (para o veredicto do Dashboard)")
inp("TIR do projeto mínima", "Crit_TIRproj", 0.08, fmt=PCT, unit="%")
inp("TIR do promotor mínima", "Crit_TIRprom", 0.12, fmt=PCT, unit="%")
inp("DSCR mínimo exigido", "Crit_DSCR", 1.25, fmt=MULT, unit="x")
inp("Payback máximo do promotor", "Crit_Payback", 10, fmt=INT, unit="anos")
inp("Múltiplo de capital (MOIC) mínimo", "Crit_MOIC", 1.5, fmt=MULT, unit="x")
ws.freeze_panes = "A4"
IN_LAST = r

# ================================================================ CAPEX
ws = ws_cx
CLASSES = ["Terreno / aquisição", "Construção", "FF&E", "Gastos amortizáveis", "Fundo de maneio"]
banner(ws, "Plano de investimento (Capex)",
       "Uma linha por rubrica. A classe define a amortização (terreno não amortiza; fundo de maneio é recuperado "
       "na saída). O valor é repartido em partes iguais pelo n.º de períodos.", 10)
widths(ws, {"A": 34, "B": 22, "C": 15, "D": 12, "E": 10, "F": 10, "G": 10, "H": 15, "I": 15, "J": 50})
header_row(ws, 4, ["Rubrica", "Classe", "Valor (€)", "Ajuste de cenário?", "Período início",
                   "N.º períodos", "Período fim", "Valor / período", "Ajustável / período", "Notas"])
ws.row_dimensions[4].height = 30
capex_rows = [
    ("Aquisição da sociedade", "Terreno / aquisição", 8600000, "Não", 0, 1,
     "Compra da Cocoon (inclui terreno e edificado existente)."),
    ("Construção / fit-out", "Construção", 8000000, "Sim", 0, 1, ""),
    ("FF&E / OS&E", "FF&E", 1000000, "Sim", 0, 1, ""),
    ("Pré-abertura", "Gastos amortizáveis", 300000, "Não", 0, 1, "Inclui payroll de pré-abertura."),
    ("Fundo de maneio", "Fundo de maneio", 400000, "Não", 0, 1, "Recuperado na saída."),
    ("Contingência", "Construção", 300000, "Sim", 0, 1, ""),
]
CX0, CX1 = 5, 24
dv_cls = DataValidation(type="list", formula1='"' + ",".join(CLASSES) + '"', allow_blank=True)
dv_cx_sn = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
ws.add_data_validation(dv_cls)
ws.add_data_validation(dv_cx_sn)
for i in range(CX0, CX1 + 1):
    data = capex_rows[i - CX0] if i - CX0 < len(capex_rows) else (None,) * 7
    for col, v in zip((1, 2, 3, 4, 5, 6, 10), data):
        c = ws.cell(row=i, column=col, value=v)
        set_input(c)
    ws.cell(row=i, column=3).number_format = EUR
    ws.cell(row=i, column=5).number_format = INT
    ws.cell(row=i, column=6).number_format = INT
    dv_cls.add(ws.cell(row=i, column=2))
    dv_cx_sn.add(ws.cell(row=i, column=4))
    ws.cell(row=i, column=7, value=f'=IF(AND(C{i}<>0,F{i}>0),E{i}+F{i}-1,"")').number_format = INT
    ws.cell(row=i, column=8, value=f'=IF(AND(N(C{i})<>0,N(F{i})>0),C{i}/F{i},0)').number_format = EUR
    ws.cell(row=i, column=9, value=f'=IF(D{i}="Sim",H{i},0)').number_format = EUR
TR = CX1 + 1
ws.cell(row=TR, column=1, value="TOTAL (cenário Base)").font = F_BOLD
ws.cell(row=TR, column=3, value=f"=SUM(C{CX0}:C{CX1})").number_format = EUR
style_range(ws, f"A{TR}:J{TR}", fill=FILL_TOT, border=B_TOP)
ws.cell(row=TR, column=3).font = F_BOLD
for nm, col in (("Capex_Classe", "B"), ("Capex_Valor", "C"), ("Capex_Ini", "E"), ("Capex_Fim", "G"),
                ("Capex_VPer", "H"), ("Capex_VAj", "I"), ("Capex_N", "F"), ("Capex_Aj", "D")):
    name(nm, ws, f"${col}${CX0}:${col}${CX1}")
r = TR + 2
section(ws, r, "RESUMO POR CLASSE (cenário Base)", 10)
r += 1
for k in CLASSES:
    ws.cell(row=r, column=1, value=k)
    ws.cell(row=r, column=3, value=f'=SUMIF(Capex_Classe,A{r},Capex_Valor)').number_format = EUR
    r += 1
ws.cell(row=r, column=1, value="Sem classe válida (deve ser 0)").font = F_BOLD
ws.cell(row=r, column=3, value=f"=C{TR}-SUM(C{r - 5}:C{r - 1})").number_format = EUR
name("Capex_SemClasse", ws, f"$C${r}")
name("Capex_Total", ws, f"$C${TR}")
ws.freeze_panes = "A5"

# ================================================================ PESSOAL
ws = ws_p
banner(ws, "Quadro de pessoal e shared services",
       "Custo anual = FTE × salário × meses × (1 + bónus) × (1 + encargos). Valores do 1.º ano de operação.", 12)
widths(ws, {"A": 20, "B": 30, "C": 8, "D": 14, "E": 8, "F": 9, "G": 10, "H": 14, "I": 12, "J": 12, "K": 15, "L": 30})
hdr = ["Departamento", "Função", "FTE", "Salário bruto / mês", "Meses", "Bónus %", "Encargos %",
       "Custo base", "Bónus", "Encargos", "Custo total anual", "Notas"]
header_row(ws, 4, hdr)
ws.row_dimensions[4].height = 30
fte = [("Management", "General Manager", 1, 5000, 0.10), ("Management", "Operations Manager", 0, 4200, 0.08),
       ("Front Office", "Front Office Manager", 1, 2500, 0.05), ("Front Office", "Receptionist", 5, 1200, 0.02),
       ("Front Office", "Night Auditor", 2, 1300, 0.02), ("Housekeeping", "Executive Housekeeper", 1, 2800, 0.05),
       ("Housekeeping", "Room Attendant", 8, 1100, 0.02), ("F&B", "F&B Manager", 1, 3200, 0.05),
       ("F&B", "Chef", 1, 3200, 0.05), ("F&B", "Sous Chef", 0, 2300, 0.03),
       ("F&B", "Kitchen / Stewarding", 4, 1200, 0.02), ("F&B", "Waiter / Bartender", 6, 1200, 0.02),
       ("Engineering", "Chief Engineer", 1, 3000, 0.05), ("Engineering", "Maintenance Technician", 3, 1300, 0.02),
       ("Sales & Marketing", "Sales Manager", 0, 3200, 0.05), ("Sales & Marketing", "Marketing / Digital", 0, 2300, 0.03),
       ("Finance / HR", "Finance Manager", 0, 3200, 0.05), ("Finance / HR", "Accountant", 0, 1900, 0.03),
       ("Finance / HR", "HR / Admin", 0, 1900, 0.03), ("IT / Security", "IT / Systems", 0, 2200, 0.03),
       ("IT / Security", "Security", 0, 1450, 0.02)]
P0, P1 = 5, 34


def staff_row(ws, i, data):
    for col, v in zip((1, 2, 3, 4, 5, 6, 7, 12), data):
        c = ws.cell(row=i, column=col, value=v)
        set_input(c)
    ws.cell(row=i, column=3).number_format = '0.0;(0.0);"-"'
    ws.cell(row=i, column=4).number_format = EUR
    ws.cell(row=i, column=5).number_format = INT
    ws.cell(row=i, column=6).number_format = PCT
    ws.cell(row=i, column=7).number_format = PCT
    ws.cell(row=i, column=8, value=f"=N(C{i})*N(D{i})*N(E{i})").number_format = EUR
    ws.cell(row=i, column=9, value=f"=H{i}*N(F{i})").number_format = EUR
    ws.cell(row=i, column=10, value=f"=(H{i}+I{i})*N(G{i})").number_format = EUR
    ws.cell(row=i, column=11, value=f"=SUM(H{i}:J{i})").number_format = EUR


for i in range(P0, P1 + 1):
    k = i - P0
    data = (fte[k][0], fte[k][1], fte[k][2], fte[k][3], 14, fte[k][4], 0.25, "") if k < len(fte) else (None,) * 8
    staff_row(ws, i, data)
TP = P1 + 1
ws.cell(row=TP, column=1, value="TOTAL PESSOAL DA UNIDADE").font = F_BOLD
for col in (3, 8, 9, 10, 11):
    c = ws.cell(row=TP, column=col, value=f"=SUM({L(col)}{P0}:{L(col)}{P1})")
    c.number_format = EUR if col != 3 else '0.0'
    c.font = F_BOLD
style_range(ws, f"A{TP}:L{TP}", fill=FILL_TOT, border=B_TOP)
name("Payroll_base", ws, f"$K${TP}")
name("FTE_total", ws, f"$C${TP}")
ws.cell(row=TP + 1, column=1, value="Custo médio por FTE")
ws.cell(row=TP + 1, column=11, value=f"=IF(C{TP}>0,K{TP}/C{TP},0)").number_format = EUR

r = TP + 3
section(ws, r, "SHARED SERVICES DO GRUPO (custo central alocado por chave)", 12)
r += 1
header_row(ws, r, hdr, 1)
r += 1
ss = [("Shared Services", "Procurement", 1, 2500, 14, 0.05, 0.25, "Interno"),
      ("Shared Services", "Sales Manager", 1, 3200, 14, 0.05, 0.25, "Interno"),
      ("Shared Services", "Accountant + TOC", 1, 6000, 12, 0, 0, "Serviço externo"),
      ("Shared Services", "HR / Admin", 1, 2500, 12, 0, 0, "Serviço externo"),
      ("Shared Services", "Marketing", 1, 3000, 12, 0, 0, "Serviço externo"),
      ("Shared Services", "IT", 1, 3000, 12, 0, 0, "Serviço externo"),
      ("Shared Services", "Legal", 1, 2000, 12, 0, 0, "Serviço externo")]
S0 = r
for k in range(10):
    staff_row(ws, r, ss[k] if k < len(ss) else (None,) * 8)
    r += 1
S1 = r - 1
ws.cell(row=r, column=1, value="TOTAL SHARED SERVICES").font = F_BOLD
c = ws.cell(row=r, column=11, value=f"=SUM(K{S0}:K{S1})")
c.number_format = EUR
c.font = F_BOLD
style_range(ws, f"A{r}:L{r}", fill=FILL_TOT, border=B_TOP)
SST = r
r += 2
ws.cell(row=r, column=1, value="Chaves do portefólio que partilham o custo").font = F_BOLD
r += 1
port = [("Cocoon (este projeto)", 48), ("Alisios", 117), ("Évora", 27), ("VDT", 43), (None, None), (None, None)]
K0 = r
for nm_, keys in port:
    c1 = ws.cell(row=r, column=2, value=nm_)
    c2 = ws.cell(row=r, column=3, value=keys)
    set_input(c1)
    set_input(c2)
    c2.number_format = INT
    r += 1
ws.cell(row=r, column=2, value="Total de chaves").font = F_BOLD
ws.cell(row=r, column=3, value=f"=SUM(C{K0}:C{r - 1})").number_format = INT
KT = r
r += 1
ws.cell(row=r, column=2, value="Custo anual por chave")
ws.cell(row=r, column=11, value=f"=IF(C{KT}>0,K{SST}/C{KT},0)").number_format = EUR
r += 1
ws.cell(row=r, column=2, value="Custo alocado a este projeto (ano 1)").font = F_BOLD
c = ws.cell(row=r, column=11, value=f"=K{r - 1}*Unidades")
c.number_format = EUR
c.font = F_BOLD
ws.cell(row=r, column=12, value="custo por chave × unidades do projeto").font = F_SUB
name("SS_anual", ws, f"$K${r}")
ws.freeze_panes = "C5"

# ================================================================ FSE
ws = ws_f
banner(ws, "Fornecimentos e serviços externos (FSE)",
       "Custo anual = base × [% fixo + % variável × (ocupação / ocupação de referência)] × inflação. "
       "Valores do 1.º ano de operação.", 7)
widths(ws, {"A": 36, "B": 15, "C": 10, "D": 12, "E": 15, "F": 15, "G": 40})
header_row(ws, 4, ["Rubrica", "Base anual (€)", "% fixo", "% variável", "Componente fixa", "Componente variável",
                   "Notas"])
ws.row_dimensions[4].height = 30
fse = [("Utilities", 230000, 0.75), ("IT", 75000, 1.0), ("Manutenção", 160000, 0.75),
       ("Jardinagem / piscina / exteriores", 60000, 0.75), ("Segurança / vigilância", 50000, 0.75),
       ("Outros FSE", 80000, 0.75)]
F0, F1 = 5, 19
for i in range(F0, F1 + 1):
    k = i - F0
    d = fse[k] if k < len(fse) else (None, None, None)
    for col, v in zip((1, 2, 3, 7), (d[0], d[1], d[2], None)):
        set_input(ws.cell(row=i, column=col, value=v))
    ws.cell(row=i, column=2).number_format = EUR
    ws.cell(row=i, column=3).number_format = PCT
    ws.cell(row=i, column=4, value=f'=IF(B{i}="","",1-N(C{i}))').number_format = PCT
    ws.cell(row=i, column=5, value=f"=N(B{i})*N(C{i})").number_format = EUR
    ws.cell(row=i, column=6, value=f"=N(B{i})*N(D{i})").number_format = EUR
TF = F1 + 1
ws.cell(row=TF, column=1, value="TOTAL FSE").font = F_BOLD
for col in (2, 5, 6):
    c = ws.cell(row=TF, column=col, value=f"=SUM({L(col)}{F0}:{L(col)}{F1})")
    c.number_format = EUR
    c.font = F_BOLD
style_range(ws, f"A{TF}:G{TF}", fill=FILL_TOT, border=B_TOP)
name("FSE_fix", ws, f"$E${TF}")
name("FSE_var", ws, f"$F${TF}")
name("FSE_base", ws, f"$B${TF}")
ws.cell(row=TF + 2, column=1,
        value="Nota: salários estão exclusivamente na folha Pessoal; custos por noite e comissões nos Pressupostos.").font = F_SUB
ws.freeze_panes = "A5"

# ================================================================ MOTOR DE CÁLCULO
NP = 21                      # períodos 0..20
C0 = 4                       # coluna D = período 0
PCOLS = [L(C0 + i) for i in range(NP)]
FIRST, LAST = PCOLS[0], PCOLS[-1]
HDR_ROW = 14


GROUPS = [("camen", "cout"), ("dstart", "dstart"), ("pab", "pfi"), ("u_ebt", "u_irc"), ("inv_n", "inv_n"),
          ("pbp", "pbj"), ("vpreco", "vcont"), ("ativo", "saida")]


def build_engine(ws, idx, scen):
    rows = {}
    banner(ws, f"Motor de cálculo — cenário {scen}",
           "Folha 100% calculada — não editar. Os valores alteram-se na folha Pressupostos.", C0 + NP - 1, nav=False)
    ws["A3"] = "Índice do cenário"
    ws["C3"] = idx
    ws["C3"].font = F_BOLD
    widths(ws, {"A": 44, "B": 12, "C": 15})
    for col in PCOLS:
        ws.column_dimensions[col].width = 12.5
    # parâmetros do cenário
    prm = {}
    rr = 4
    for label, nm, *_ in scen_inputs:
        ws.cell(row=rr, column=1, value=label)
        c = ws.cell(row=rr, column=3, value=f"=INDEX({nm},1,$C$3)")
        c.font = F_LINK
        c.number_format = PCT
        prm[nm] = f"$C${rr}"
        rr += 1
    # cabeçalho temporal
    header_row(ws, HDR_ROW, ["Rubrica", "Unidade", "Total / KPI"] + [None] * NP)
    for i, col in enumerate(PCOLS):
        ws[f"{col}{HDR_ROW}"] = i
        ws[f"{col}{HDR_ROW}"].number_format = '"P"0'
    ws.freeze_panes = f"D{HDR_ROW + 1}"

    cur = [HDR_ROW + 1]

    def P(col):
        return f"{col}${HDR_ROW}"

    specs = []

    def add(key, label, f, unit="€", fmt=EUR, total=None, bold=False, line=False):
        """f(col, prev, R, RP) -> fórmula. As fórmulas são escritas no fim (permite referências a linhas abaixo)."""
        rows[key] = cur[0]
        specs.append((cur[0], label, f, unit, fmt, total, bold, line))
        cur[0] += 1
        return rows[key]

    def write_specs():
        for r, label, f, unit, fmt, total, bold, line in specs:
            ws.cell(row=r, column=1, value=label).font = F_BOLD if bold else F_BASE
            ws.cell(row=r, column=2, value=unit).font = F_SUB
            for i, col in enumerate(PCOLS):
                prev = PCOLS[i - 1] if i > 0 else None

                def R(k, c=col):
                    return f"{c}{rows[k]}"

                def RP(k, p=prev):
                    return f"{p}{rows[k]}" if p else "0"
                c = ws.cell(row=r, column=C0 + i, value="=" + f(col, prev, R, RP))
                c.number_format = fmt
                if bold:
                    c.font = F_BOLD
            if total == "sum":
                ws.cell(row=r, column=3, value=f"=SUM({FIRST}{r}:{LAST}{r})").number_format = fmt
            if bold or line:
                style_range(ws, f"A{r}:{LAST}{r}", border=B_TOP)
                if bold:
                    style_range(ws, f"A{r}:{LAST}{r}", fill=FILL_TOT)

    def sec_(text):
        cur[0] += 1
        section(ws, cur[0], text, C0 + NP - 1)
        cur[0] += 1

    # --- calendário
    sec_("CALENDÁRIO")
    add("ano", "Ano civil", lambda c, p, R, RP: f"AnoInicio+{P(c)}", unit="", fmt=YEAR)
    add("ativo", "Período dentro do horizonte (1/0)", lambda c, p, R, RP: f"IF({P(c)}<=AnoSaida,1,0)", unit="flag",
        fmt=INT)
    add("op", "Ano de operação", lambda c, p, R, RP: f"IF(AND({R('ativo')}=1,{P(c)}>=AnoAbertura),{P(c)}-AnoAbertura+1,0)",
        unit="n.º", fmt=INT)
    add("saida", "Período de saída (1/0)", lambda c, p, R, RP: f"IF({P(c)}=AnoSaida,1,0)", unit="flag", fmt=INT)

    # --- operação
    sec_("OPERAÇÃO")
    add("vend", "Unidades vendidas no ano",
        lambda c, p, R, RP: f"IF(AND({R('op')}>=1,{R('op')}<=5),INDEX(Vendas_tab,MIN(5,MAX(1,{R('op')}))),0)",
        unit="unid.", fmt=INT, total="sum")
    add("vendac", "Unidades vendidas acumuladas", lambda c, p, R, RP: f"{RP('vendac')}+{R('vend')}", unit="unid.",
        fmt=INT)
    add("unid", "Unidades em exploração",
        lambda c, p, R, RP: f"IF({R('op')}>0,MAX(0,Unidades-IF(Vendas_reduz=\"Sim\",{R('vendac')}-{R('vend')},0)),0)",
        unit="unid.", fmt=INT)
    add("ndisp", "Noites disponíveis", lambda c, p, R, RP: f"{R('unid')}*Dias", unit="noites", total="sum")
    add("occ", "Taxa de ocupação",
        lambda c, p, R, RP: f"IF({R('op')}>0,MIN({prm['S_Ocup1']}+{prm['S_OcupG']}*({R('op')}-1),{prm['S_OcupMax']}),0)",
        unit="%", fmt=PCT)
    add("nocc", "Noites ocupadas", lambda c, p, R, RP: f"{R('ndisp')}*{R('occ')}", unit="noites", total="sum")
    add("adr", "ADR (preço médio por noite)",
        lambda c, p, R, RP: (f"IF({R('op')}=0,0,IF(AND({R('op')}<=5,INDEX(ADR_tab,MIN(5,MAX(1,{R('op')})))>0),"
                             f"INDEX(ADR_tab,MIN(5,MAX(1,{R('op')})))*(1+{prm['S_AjADR']}),{RP('adr')}*(1+ADR_g)))"),
        unit="€", fmt=EUR2)
    add("revpar", "RevPAR", lambda c, p, R, RP: f"IF({R('ndisp')}>0,{R('adr')}*{R('occ')},0)", unit="€", fmt=EUR2)

    sec_("CONTA DE EXPLORAÇÃO")
    add("rq", "Receita de quartos", lambda c, p, R, RP: f"{R('nocc')}*{R('adr')}", total="sum")
    add("rfb", "Receita F&B",
        lambda c, p, R, RP: f"IF({R('op')}>0,{R('nocc')}*FB_quarto*(1+FB_g)^({R('op')}-1),0)", total="sum")
    add("rout", "Outros proveitos", lambda c, p, R, RP: f"{R('rq')}*Outros_pct", total="sum")
    add("rtot", "RECEITA OPERACIONAL TOTAL", lambda c, p, R, RP: f"SUM({c}{rows['rq']}:{c}{rows['rout']})",
        total="sum", bold=True)
    aj = prm["S_AjCustos"]
    add("camen", "Amenities e consumíveis",
        lambda c, p, R, RP: f"IF({R('op')}>0,{R('nocc')}*C_Amen*(1+InflDrivers)^({R('op')}-1)*(1+{aj}),0)", total="sum")
    add("clav", "Lavandaria e roupa",
        lambda c, p, R, RP: f"IF({R('op')}>0,{R('nocc')}*C_Lav*(1+InflDrivers)^({R('op')}-1)*(1+{aj}),0)", total="sum")
    add("cota", "Comissões OTA", lambda c, p, R, RP: f"{R('rq')}*OTA_peso*OTA_com", total="sum")
    add("cpag", "Taxas de pagamento", lambda c, p, R, RP: f"{R('rtot')}*TaxasPag", total="sum")
    add("cfood", "Food COGS", lambda c, p, R, RP: f"{R('rfb')}*Food_peso*Food_cogs", total="sum")
    add("cbev", "Beverage COGS", lambda c, p, R, RP: f"{R('rfb')}*(1-Food_peso)*Bev_cogs", total="sum")
    add("cfbc", "Consumíveis F&B", lambda c, p, R, RP: f"{R('rfb')}*FB_cons", total="sum")
    add("cout", "Custo direto de outros proveitos", lambda c, p, R, RP: f"{R('rout')}*CustoOutros_pct", total="sum")
    add("cdir", "Total custos diretos", lambda c, p, R, RP: f"SUM({c}{rows['camen']}:{c}{rows['cout']})",
        total="sum", line=True)
    add("pay", "Pessoal (payroll)",
        lambda c, p, R, RP: f"IF({R('op')}>0,Payroll_base*(1+InflSal)^({R('op')}-1)*(1+{aj}),0)", total="sum")
    add("fse", "FSE",
        lambda c, p, R, RP: (f"IF({R('op')}>0,(FSE_fix+FSE_var*IF(OcupRef>0,{R('occ')}/OcupRef,1))"
                             f"*(1+InflFSE)^({R('op')}-1)*(1+{aj}),0)"), total="sum")
    add("gop", "GOP (resultado operacional bruto)",
        lambda c, p, R, RP: f"{R('rtot')}-{R('cdir')}-{R('pay')}-{R('fse')}", total="sum", bold=True)
    add("gopm", "Margem GOP", lambda c, p, R, RP: f"IF({R('rtot')}>0,{R('gop')}/{R('rtot')},0)", unit="%", fmt=PCT)
    add("mfee", "Management fee", lambda c, p, R, RP: f"{R('rtot')}*MgmtFee", total="sum")
    add("ifee", "Incentive fee", lambda c, p, R, RP: f"MAX(0,{R('gop')})*IncFee", total="sum")
    add("ss", "Shared services do grupo",
        lambda c, p, R, RP: f"IF({R('op')}>0,SS_anual*(1+SS_g)^({R('op')}-1),0)", total="sum")
    add("ebitda", "EBITDA", lambda c, p, R, RP: f"{R('gop')}-{R('mfee')}-{R('ifee')}-{R('ss')}", total="sum",
        bold=True)
    add("ebitdam", "Margem EBITDA", lambda c, p, R, RP: f"IF({R('rtot')}>0,{R('ebitda')}/{R('rtot')},0)", unit="%",
        fmt=PCT)
    add("ffe", "Reserva FF&E",
        lambda c, p, R, RP: f"IF(AND({R('op')}>0,{R('op')}>=FFE_inicio),{R('rtot')}*FFE_res,0)", total="sum")
    add("cxm", "CAPEX de manutenção adicional",
        lambda c, p, R, RP: f"IF({R('op')}>0,CapexManut*(1+CapexManut_g)^({R('op')}-1),0)", total="sum")
    add("noi", "NOI (EBITDA após reserva FF&E)", lambda c, p, R, RP: f"{R('ebitda')}-{R('ffe')}-{R('cxm')}",
        total="sum", bold=True)

    sec_("VENDA DE UNIDADES")
    add("vpreco", "Preço por unidade",
        lambda c, p, R, RP: f"IF({R('op')}>0,PrecoUnid*(1+PrecoUnid_g)^({R('op')}-1)*(1+{prm['S_AjPreco']}),0)",
        fmt=EUR)
    add("vrec", "Receita de venda de unidades", lambda c, p, R, RP: f"{R('vend')}*{R('vpreco')}", total="sum")
    add("vcus", "Custos de venda / comissões", lambda c, p, R, RP: f"{R('vrec')}*CustoVenda_pct", total="sum")
    add("vcont", "Custo contabilístico das unidades vendidas", lambda c, p, R, RP: f"{R('vend')}*CustoContab_unid",
        total="sum")
    add("vres", "Resultado com venda de unidades", lambda c, p, R, RP: f"{R('vrec')}-{R('vcus')}-{R('vcont')}",
        total="sum", bold=True)

    sec_("INVESTIMENTO E AMORTIZAÇÕES")
    ajc = prm["S_AjCapex"]
    cls_key = {"Terreno / aquisição": "cx_ter", "Construção": "cx_con", "FF&E": "cx_ffe",
               "Gastos amortizáveis": "cx_gas", "Fundo de maneio": "cx_fdm"}
    for k in CLASSES:
        add(cls_key[k], f"Capex — {k}",
            lambda c, p, R, RP, k=k: (f"{R('ativo')}*(SUMPRODUCT((Capex_Classe=\"{k}\")*(Capex_Ini<={P(c)})*"
                                      f"(Capex_Fim>={P(c)})*Capex_VPer)+{ajc}*SUMPRODUCT((Capex_Classe=\"{k}\")*"
                                      f"(Capex_Ini<={P(c)})*(Capex_Fim>={P(c)})*Capex_VAj))"), total="sum")
    add("cxtot", "CAPEX TOTAL", lambda c, p, R, RP: f"SUM({c}{rows['cx_ter']}:{c}{rows['cx_fdm']})", total="sum",
        bold=True)
    add("dstart", "Início de amortização do investimento do período",
        lambda c, p, R, RP: f"MAX({P(c)}+1,AnoAbertura)", unit="período", fmt=INT)
    rs = rows["dstart"]

    def dep(cx, life):
        return lambda c, p, R, RP: (f"{R('ativo')}*SUMPRODUCT(${FIRST}${rows[cx]}:${LAST}${rows[cx]},"
                                    f"--(${FIRST}${rs}:${LAST}${rs}<={P(c)}),"
                                    f"--(${FIRST}${rs}:${LAST}${rs}+{life}>{P(c)}))/MAX(1,{life})")
    add("dcon", "Amortização — construção", dep("cx_con", "VidaConstr"), total="sum")
    add("dffe", "Amortização — FF&E", dep("cx_ffe", "VidaFFE"), total="sum")
    add("dgas", "Amortização — gastos amortizáveis", dep("cx_gas", "VidaOutros"), total="sum")
    add("dtot", "Total amortizações", lambda c, p, R, RP: f"SUM({c}{rows['dcon']}:{c}{rows['dgas']})", total="sum",
        line=True)
    add("vcb", "Valor contabilístico do ativo (fim do período)",
        lambda c, p, R, RP: f"{RP('vcb')}+{R('cxtot')}-{R('cx_fdm')}-{R('dtot')}-{R('vcont')}")

    sec_("DÍVIDA BANCÁRIA")
    nper = "MAX(1,Prazo-Carencia)"
    add("ddes", "Desembolso", lambda c, p, R, RP: f"IF(AND({P(c)}=DividaPer,{R('ativo')}=1),Divida,0)", total="sum")
    add("dabe", "Capital em dívida — início", lambda c, p, R, RP: f"{RP('dfim')}")
    add("djur", "Juros", lambda c, p, R, RP: f"{R('dabe')}*Juro", total="sum")
    add("dsel", "Imposto do selo sobre juros", lambda c, p, R, RP: f"{R('djur')}*Selo", total="sum")
    add("dcom", "Comissão de abertura", lambda c, p, R, RP: f"{R('ddes')}*ComAbertura", total="sum")
    add("damo", "Amortização de capital",
        lambda c, p, R, RP: (f"IF(AND({R('dabe')}>0,{P(c)}-DividaPer-Carencia>=1,{P(c)}-DividaPer-Carencia<={nper}),"
                             f"MIN({R('dabe')},IF(TipoAmort=\"Linear\",Divida/{nper},"
                             f"PPMT(Juro,{P(c)}-DividaPer-Carencia,{nper},-Divida))),0)"), total="sum")
    add("dreemb", "Reembolso antecipado na saída",
        lambda c, p, R, RP: f"IF({R('saida')}=1,{R('dabe')}+{R('ddes')}-{R('damo')},0)", total="sum")
    add("dfim", "Capital em dívida — fim", lambda c, p, R, RP: f"{R('dabe')}+{R('ddes')}-{R('damo')}-{R('dreemb')}")
    add("dserv", "Serviço da dívida (juros + selo + capital)",
        lambda c, p, R, RP: f"{R('djur')}+{R('dsel')}+{R('damo')}", total="sum", line=True)

    sec_("DEMONSTRAÇÃO DE RESULTADOS")
    add("pl_noi", "NOI", lambda c, p, R, RP: f"{R('noi')}")
    add("pl_v", "Resultado com venda de unidades", lambda c, p, R, RP: f"{R('vres')}")
    add("pl_d", "Amortizações", lambda c, p, R, RP: f"-{R('dtot')}")
    add("ebit", "EBIT (resultado operacional)", lambda c, p, R, RP: f"{R('pl_noi')}+{R('pl_v')}+{R('pl_d')}",
        total="sum", bold=True)
    add("pl_fin", "Custos financeiros", lambda c, p, R, RP: f"-({R('djur')}+{R('dsel')}+{R('dcom')})", total="sum")
    add("ebt", "Resultado antes de impostos", lambda c, p, R, RP: f"{R('ebit')}+{R('pl_fin')}", total="sum",
        bold=True)
    add("pab", "Prejuízos fiscais reportáveis — início", lambda c, p, R, RP: f"{RP('pfi')}")
    add("puso", "Dedução de prejuízos", lambda c, p, R, RP: f"IF({R('ebt')}>0,MIN({R('pab')},{R('ebt')}*LimitePrej),0)")
    add("pfi", "Prejuízos fiscais reportáveis — fim",
        lambda c, p, R, RP: f"{R('pab')}+MAX(0,-{R('ebt')})-{R('puso')}")
    add("irc", "IRC", lambda c, p, R, RP: f"MAX(0,{R('ebt')}-{R('puso')})*IRC", total="sum")
    add("rl", "RESULTADO LÍQUIDO", lambda c, p, R, RP: f"{R('ebt')}-{R('irc')}", total="sum", bold=True)

    sec_("VALOR DE SAÍDA")
    add("tv", "Valor de saída (NOI / exit yield)",
        lambda c, p, R, RP: f"IF(AND({R('saida')}=1,{prm['S_ExitYield']}>0),{R('noi')}/{prm['S_ExitYield']},0)",
        total="sum")
    add("tvc", "Custos de transação", lambda c, p, R, RP: f"{R('tv')}*CustosSaida_pct", total="sum")
    add("tvt", "Imposto sobre mais-valia",
        lambda c, p, R, RP: f"IF({R('saida')}=1,MAX(0,{R('tv')}-{R('tvc')}-{R('vcb')})*TaxaMaisValia,0)", total="sum")
    add("tvf", "Recuperação do fundo de maneio",
        lambda c, p, R, RP: f"IF({R('saida')}=1,SUM(${FIRST}{rows['cx_fdm']}:{c}{rows['cx_fdm']}),0)", total="sum")
    add("tvl", "Valor de saída líquido", lambda c, p, R, RP: f"{R('tv')}-{R('tvc')}-{R('tvt')}+{R('tvf')}",
        total="sum", bold=True)

    sec_("CASH FLOW DO PROJETO (sem financiamento)")
    add("u_ebt", "Base tributável sem dívida (EBIT)", lambda c, p, R, RP: f"{R('ebit')}")
    add("u_pab", "Prejuízos reportáveis (sem dívida) — início", lambda c, p, R, RP: f"{RP('u_pfi')}")
    add("u_uso", "Dedução de prejuízos (sem dívida)",
        lambda c, p, R, RP: f"IF({R('u_ebt')}>0,MIN({R('u_pab')},{R('u_ebt')}*LimitePrej),0)")
    add("u_pfi", "Prejuízos reportáveis (sem dívida) — fim",
        lambda c, p, R, RP: f"{R('u_pab')}+MAX(0,-{R('u_ebt')})-{R('u_uso')}")
    add("u_irc", "IRC sem dívida", lambda c, p, R, RP: f"MAX(0,{R('u_ebt')}-{R('u_uso')})*IRC", total="sum")
    add("cf_noi", "NOI", lambda c, p, R, RP: f"{R('noi')}", total="sum")
    add("cf_v", "Venda de unidades (líquida de custos)", lambda c, p, R, RP: f"{R('vrec')}-{R('vcus')}", total="sum")
    add("cf_cx", "Investimento (capex)", lambda c, p, R, RP: f"-{R('cxtot')}", total="sum")
    add("cf_sub", "Subsídio a fundo perdido",
        lambda c, p, R, RP: f"IF(AND({P(c)}=SubsidioPer,{R('ativo')}=1),Subsidio,0)", total="sum")
    add("cf_irc", "IRC (sem dívida)", lambda c, p, R, RP: f"-{R('u_irc')}", total="sum")
    add("cf_tv", "Valor de saída líquido", lambda c, p, R, RP: f"{R('tvl')}", total="sum")
    add("fcfp", "CASH FLOW LIVRE DO PROJETO",
        lambda c, p, R, RP: f"SUM({c}{rows['cf_noi']}:{c}{rows['cf_tv']})", total="sum", bold=True)
    add("fcfpa", "Cash flow do projeto acumulado", lambda c, p, R, RP: f"{RP('fcfpa')}+{R('fcfp')}")

    sec_("CASH FLOW DO CAPITAL PRÓPRIO")
    add("e_fcf", "Cash flow livre do projeto", lambda c, p, R, RP: f"{R('fcfp')}", total="sum")
    add("e_tax", "Efeito fiscal da dívida (IRC sem dívida − IRC)", lambda c, p, R, RP: f"{R('u_irc')}-{R('irc')}",
        total="sum")
    add("e_des", "Desembolso de dívida", lambda c, p, R, RP: f"{R('ddes')}", total="sum")
    add("e_srv", "Juros, selo e comissões", lambda c, p, R, RP: f"-({R('djur')}+{R('dsel')}+{R('dcom')})",
        total="sum")
    add("e_amo", "Amortização e reembolso de capital", lambda c, p, R, RP: f"-({R('damo')}+{R('dreemb')})",
        total="sum")
    add("fcfe", "CASH FLOW DO CAPITAL PRÓPRIO (total)",
        lambda c, p, R, RP: f"SUM({c}{rows['e_fcf']}:{c}{rows['e_amo']})", total="sum", bold=True)
    add("inv_n", "Anos até à recompra do investidor", lambda c, p, R, RP: "MIN(InvPer+InvRecompra,AnoSaida)-InvPer",
        unit="anos", fmt=INT)
    add("inv_in", "Entrada do investidor financeiro",
        lambda c, p, R, RP: f"IF(AND({P(c)}=InvPer,{R('ativo')}=1),InvMontante,0)", total="sum")
    add("inv_out", "Recompra ao investidor financeiro",
        lambda c, p, R, RP: (f"IF({P(c)}=MIN(InvPer+InvRecompra,AnoSaida),InvMontante*IF(InvRegime=\"Simples\","
                             f"1+InvTaxa*{R('inv_n')},(1+InvTaxa)^{R('inv_n')}),0)"), total="sum")
    add("fcfs", "CASH FLOW DO PROMOTOR",
        lambda c, p, R, RP: f"{R('fcfe')}+{R('inv_in')}-{R('inv_out')}", total="sum", bold=True)
    add("fcfsa", "Cash flow do promotor acumulado", lambda c, p, R, RP: f"{RP('fcfsa')}+{R('fcfs')}")
    add("fcfi", "Cash flow do investidor financeiro", lambda c, p, R, RP: f"-{R('inv_in')}+{R('inv_out')}",
        total="sum")

    add("bcf", "Cash flow do banco", lambda c, p, R, RP: f"-{R('ddes')}+{R('djur')}+{R('dcom')}+{R('damo')}+{R('dreemb')}",
        total="sum")

    sec_("RÁCIOS DE DÍVIDA")
    add("cfads", "CFADS (NOI + vendas líquidas − IRC)",
        lambda c, p, R, RP: f"{R('noi')}+{R('vrec')}-{R('vcus')}-{R('irc')}", total="sum")
    add("dscr", "DSCR (CFADS / serviço da dívida)",
        lambda c, p, R, RP: f"IF(AND({R('dserv')}>0,{R('op')}>0),{R('cfads')}/{R('dserv')},\"\")", unit="x", fmt=MULT)
    add("pbp", "Payback do promotor (período de recuperação)",
        lambda c, p, R, RP: f"IF(AND({R('fcfsa')}>=0,{RP('fcfsa')}<0),{P(c)},\"\")", unit="período", fmt=INT)
    add("pbj", "Payback do projeto (período de recuperação)",
        lambda c, p, R, RP: f"IF(AND({R('fcfpa')}>=0,{RP('fcfpa')}<0),{P(c)},\"\")", unit="período", fmt=INT)

    write_specs()

    # --- KPIs
    def rng(k):
        return f"{FIRST}{rows[k]}:{LAST}{rows[k]}"

    def rng1(k):
        return f"{PCOLS[1]}{rows[k]}:{LAST}{rows[k]}"
    cur[0] += 1
    section(ws, cur[0], "INDICADORES (KPI)", C0 + NP - 1)
    cur[0] += 1
    kpis = [
        ("k_cxt", "Investimento total", f"SUM({rng('cxtot')})", EUR),
        ("k_div", "Dívida bancária", f"SUM({rng('ddes')})", EUR),
        ("k_ltc", "Dívida / investimento (LTC)", f"IF(C{{k_cxt}}>0,C{{k_div}}/C{{k_cxt}},0)", PCT),
        ("k_sub", "Subsídio a fundo perdido", f"SUM({rng('cf_sub')})", EUR),
        ("k_inv", "Investidor financeiro", f"SUM({rng('inv_in')})", EUR),
        ("k_eqmax", "Exposição máxima do promotor", f"MAX(0,-MIN({rng('fcfsa')}))", EUR),
        ("k_tirp", "TIR do projeto (sem dívida)", f"IFERROR(IRR({rng('fcfp')}),\"n.a.\")", PCT),
        ("k_valp", "VAL do projeto", f"NPV(TaxaDescProj,{rng1('fcfp')})+{FIRST}{rows['fcfp']}", EUR),
        ("k_tire", "TIR do capital próprio (total)", f"IFERROR(IRR({rng('fcfe')}),\"n.a.\")", PCT),
        ("k_tirs", "TIR do promotor", f"IFERROR(IRR({rng('fcfs')}),\"n.a.\")", PCT),
        ("k_vals", "VAL do promotor", f"NPV(CustoEquity,{rng1('fcfs')})+{FIRST}{rows['fcfs']}", EUR),
        ("k_moic", "Múltiplo de capital do promotor (MOIC)",
         f"IF(SUMIF({rng('fcfs')},\"<0\")<0,SUMIF({rng('fcfs')},\">0\")/-SUMIF({rng('fcfs')},\"<0\"),\"n.a.\")", MULT),
        ("k_pbs", "Payback do promotor (anos)", f"IF(COUNT({rng('pbp')})=0,\"n.a.\",MAX({rng('pbp')}))", INT),
        ("k_pbj", "Payback do projeto (anos)", f"IF(COUNT({rng('pbj')})=0,\"n.a.\",MAX({rng('pbj')}))", INT),
        ("k_dmin", "DSCR mínimo", f"IF(COUNT({rng('dscr')})=0,\"n.a.\",MIN({rng('dscr')}))", MULT),
        ("k_dmed", "DSCR médio", f"IF(COUNT({rng('dscr')})=0,\"n.a.\",AVERAGE({rng('dscr')}))", MULT),
        ("k_tiri", "TIR do investidor financeiro", f"IFERROR(IRR({rng('fcfi')}),\"n.a.\")", PCT),
        ("k_r1", "Receita — 1.º ano de operação",
         f"IFERROR(INDEX({rng('rtot')},MATCH(1,{rng('op')},0)),0)", EUR),
        ("k_rs", "Receita — ano 5 de operação (estabilizado)",
         f"IFERROR(INDEX({rng('rtot')},MATCH(5,{rng('op')},0)),0)", EUR),
        ("k_es", "EBITDA — ano 5 de operação", f"IFERROR(INDEX({rng('ebitda')},MATCH(5,{rng('op')},0)),0)", EUR),
        ("k_em", "Margem EBITDA — ano 5", f"IF(C{{k_rs}}>0,C{{k_es}}/C{{k_rs}},0)", PCT),
        ("k_ns", "NOI — ano 5 de operação", f"IFERROR(INDEX({rng('noi')},MATCH(5,{rng('op')},0)),0)", EUR),
        ("k_yoc", "Yield on cost (NOI ano 5 / investimento)", f"IF(C{{k_cxt}}>0,C{{k_ns}}/C{{k_cxt}},0)", PCT),
        ("k_tv", "Valor de saída (bruto)", f"SUM({rng('tv')})", EUR),
        ("k_vend", "Receita de venda de unidades", f"SUM({rng('vrec')})", EUR),
        ("k_rl", "Resultado líquido acumulado", f"SUM({rng('rl')})", EUR),
        ("k_lucro", "Lucro do promotor (cash in − cash out)", f"SUM({rng('fcfs')})", EUR),
        ("k_dfim", "Dívida por reembolsar após a saída", f"{LAST}{rows['dfim']}", EUR),
        ("k_sneg", "Promotor — total investido", f"-SUMIF({rng('fcfs')},\"<0\")", EUR),
        ("k_spos", "Promotor — total recebido", f"SUMIF({rng('fcfs')},\">0\")", EUR),
        ("k_invout", "Investidor — total recebido", f"SUM({rng('inv_out')})", EUR),
        ("k_brec", "Banco — total recebido (juros, comissões e capital)",
         f"SUM({rng('djur')})+SUM({rng('dcom')})+SUM({rng('damo')})+SUM({rng('dreemb')})", EUR),
        ("k_tirb", "TIR do banco", f"IFERROR(IRR({rng('bcf')}),\"n.a.\")", PCT),
        ("k_ter", "Capex — terreno / aquisição", f"SUM({rng('cx_ter')})", EUR),
        ("k_con", "Capex — construção", f"SUM({rng('cx_con')})", EUR),
        ("k_ffe", "Capex — FF&E", f"SUM({rng('cx_ffe')})", EUR),
        ("k_gas", "Capex — gastos amortizáveis", f"SUM({rng('cx_gas')})", EUR),
        ("k_fdm", "Capex — fundo de maneio", f"SUM({rng('cx_fdm')})", EUR),
    ]
    for key, label, f, fmt in kpis:
        rows[key] = cur[0]
        cur[0] += 1
    for key, label, f, fmt in kpis:
        rr = rows[key]
        ws.cell(row=rr, column=1, value=label).font = F_BOLD
        f = f.replace("{k_cxt}", str(rows["k_cxt"])).replace("{k_div}", str(rows["k_div"])) \
             .replace("{k_rs}", str(rows["k_rs"])).replace("{k_es}", str(rows["k_es"])) \
             .replace("{k_ns}", str(rows["k_ns"]))
        c = ws.cell(row=rr, column=3, value="=" + f)
        c.number_format = fmt
        c.font = F_BOLD
        c.fill = FILL_TOT
    # linhas de detalhe agrupadas (podem ser recolhidas com o botão "−" à esquerda)
    for a_, b_ in GROUPS:
        ws.row_dimensions.group(rows[a_], rows[b_], outline_level=1, hidden=False)
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_properties.tabColor = "A6A6A6"
    ws.sheet_view.showGridLines = False
    return rows


ENG = None
for i, (wse, s) in enumerate(zip(ws_eng, SCEN)):
    ENG = build_engine(wse, i + 1, s)

# ================================================================ DASHBOARD
ws = ws_dash
ws.sheet_view.showGridLines = False
widths(ws, {"A": 2, **{L(c): 12.5 for c in range(2, 14)}, "N": 2, "O": 2, "P": 22, "Q": 14})
banner(ws, '="Avaliação de investimento  ·  "&NomeProjeto',
       '=Localizacao&"   ·   "&TipoAtivo&"   ·   Análise: "&Analista', 13, start_col=2)

ws["B5"] = "Cenário  ▸"
ws["B5"].font = F_BOLD
ws["B5"].alignment = Alignment(vertical="center")
c = merge(ws, 5, 3, 5, 4, value="Base", font=Font(name=FONT, size=11, bold=True, color="0000FF"),
          fill=FILL_KEY, border=B_IN, align=Alignment(horizontal="center", vertical="center"))
dv_sc = DataValidation(type="list", formula1="ListaCenarios", allow_blank=False,
                       promptTitle="Cenário", prompt="Escolha Base, Pessimista ou Otimista.")
ws.add_data_validation(dv_sc)
dv_sc.add(ws["C5"])
name("CenarioAtivo", ws, "$C$5")
ws["P5"] = "=MATCH(CenarioAtivo,ListaCenarios,0)"
ws["P5"].font = Font(name=FONT, size=8, color="BFBFBF")
name("CenIdx", ws, "$P$5")
ws["F5"] = "Estado do modelo"
ws["F5"].font = F_BOLD
ws["F5"].alignment = Alignment(horizontal="right", vertical="center")
merge(ws, 5, 7, 5, 7, value="=ChecksStatus", font=F_BOLD, align=Alignment(horizontal="center", vertical="center"),
      border=B_BOX)
ws["I5"] = "Horizonte"
ws["I5"].font = F_BOLD
ws["I5"].alignment = Alignment(horizontal="right", vertical="center")
c = ws["J5"]
c.value = "=AnoSaida"
c.number_format = ANOS
c.alignment = Alignment(horizontal="center")
ws["K5"] = "Abertura"
ws["K5"].font = F_BOLD
ws["K5"].alignment = Alignment(horizontal="right", vertical="center")
ws["L5"] = "=AnoInicio+AnoAbertura"
ws["L5"].alignment = Alignment(horizontal="center")
ws.row_dimensions[5].height = 22
for d_ in ("G5",):
    ws.conditional_formatting.add(d_, CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_OK))
    ws.conditional_formatting.add(d_, CellIsRule(operator="equal", formula=['"ERRO"'], fill=FILL_BAD))
    ws.conditional_formatting.add(d_, CellIsRule(operator="equal", formula=['"OK c/ avisos"'], fill=FILL_WARN))


def eng_ref(key, col="C"):
    """Valor do motor do cenário ativo."""
    return "CHOOSE(CenIdx," + ",".join(f"'{w.title}'!${col}${ENG[key]}" for w in ws_eng) + ")"


# --- critérios (linhas 17-21) — calculados primeiro porque o veredicto e os cartões dependem deles
CR0 = 17
crit = [
    ("TIR do projeto (sem dívida)", "k_tirp", "Crit_TIRproj", PCT, ">="),
    ("TIR do promotor", "k_tirs", "Crit_TIRprom", PCT, ">="),
    ("DSCR mínimo", "k_dmin", "Crit_DSCR", MULT, ">="),
    ("Payback do promotor", "k_pbs", "Crit_Payback", ANOS, "<="),
    ("Múltiplo de capital do promotor", "k_moic", "Crit_MOIC", MULT, ">="),
]
section(ws, 15, "CRITÉRIOS DE VIABILIDADE", 6, start_col=2)
for col, t in ((2, "Critério"), (5, "Resultado"), (6, "Limite"), (7, "Cumpre?")):
    ws.cell(row=16, column=col, value=t)
header_row(ws, 16, ["Critério", None, None, "Resultado", "Limite", "Cumpre?"], start_col=2)
ws.merge_cells("B16:D16")
for i, (lab, k, cn, fmt, op) in enumerate(crit):
    rr = CR0 + i
    merge(ws, rr, 2, rr, 4, value=lab, border=B_ROW)
    c = ws.cell(row=rr, column=5, value="=" + eng_ref(k))
    c.number_format = fmt
    c.font = F_BOLD
    c.alignment = Alignment(horizontal="center")
    c = ws.cell(row=rr, column=6, value=f"={cn}")
    c.number_format = fmt
    c.font = F_LINK
    c.alignment = Alignment(horizontal="center")
    if k == "k_dmin":
        f = f'=IF(ISNUMBER(E{rr}),IF(E{rr}{op}F{rr},"✓","✗"),IF(Divida=0,"✓","✗"))'
    else:
        f = f'=IF(ISNUMBER(E{rr}),IF(E{rr}{op}F{rr},"✓","✗"),"✗")'
    c = ws.cell(row=rr, column=7, value=f)
    c.alignment = Alignment(horizontal="center")
    c.font = Font(name=FONT, size=11, bold=True)
    for col in range(2, 8):
        ws.cell(row=rr, column=col).border = B_ROW
CRIT_RNG = f"G{CR0}:G{CR0 + 4}"
ws.conditional_formatting.add(CRIT_RNG, CellIsRule(operator="equal", formula=['"✓"'], fill=FILL_OK,
                                                   font=Font(color=GREEN, bold=True)))
ws.conditional_formatting.add(CRIT_RNG, CellIsRule(operator="equal", formula=['"✗"'], fill=FILL_BAD,
                                                   font=Font(color=RED, bold=True)))

# --- veredicto
ws["P7"] = (f'=IF(COUNTIF({CRIT_RNG},"✗")=0,"VIÁVEL",IF(COUNTIF({CRIT_RNG},"✗")<=1,'
            f'"VIÁVEL COM RESERVAS","NÃO VIÁVEL"))')
ws["P7"].font = Font(name=FONT, size=8, color="BFBFBF")
merge(ws, 7, 2, 8, 13, value=f'=P7&"     ·     cumpre "&COUNTIF({CRIT_RNG},"✓")&" de 5 critérios"',
      font=Font(name=FONT, size=16, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=GREY),
      align=Alignment(horizontal="center", vertical="center"))
for v_, bg in (("VIÁVEL", GREEN), ("VIÁVEL COM RESERVAS", "B7791F"), ("NÃO VIÁVEL", RED)):
    ws.conditional_formatting.add("B7:M8", FormulaRule(formula=[f'$P$7="{v_}"'],
                                                       fill=PatternFill("solid", fgColor=bg, bgColor=bg)))
ws.row_dimensions[7].height = 20
ws.row_dimensions[8].height = 20

# --- cartões de KPI (linhas 10-13)
cards = [
    ("TIR DO PROJETO", "k_tirp", PCT, '="Mínimo "&ROUND(Crit_TIRproj*100,1)&"%"', f"$G${CR0}"),
    ("VAL DO PROJETO", "k_valp", MEUR, '="Taxa de desconto "&ROUND(TaxaDescProj*100,1)&"%"', None),
    ("TIR DO PROMOTOR", "k_tirs", PCT, '="Mínimo "&ROUND(Crit_TIRprom*100,1)&"%"', f"$G${CR0 + 1}"),
    ("MÚLTIPLO (MOIC)", "k_moic", MULT, '="Mínimo "&Crit_MOIC&"x"', f"$G${CR0 + 4}"),
    ("DSCR MÍNIMO", "k_dmin", MULT, '="Mínimo "&Crit_DSCR&"x"', f"$G${CR0 + 2}"),
    ("PAYBACK DO PROMOTOR", "k_pbs", ANOS, '="Máximo "&Crit_Payback&" anos"', f"$G${CR0 + 3}"),
]
for i, (lab, k, fmt, sub, crit_cell) in enumerate(cards):
    c1 = 2 + 2 * i
    merge(ws, 10, c1, 10, c1 + 1, value=lab, font=Font(name=FONT, size=8, bold=True, color=GREY),
          fill=FILL_CARD, align=Alignment(horizontal="center", vertical="bottom"))
    v = merge(ws, 11, c1, 12, c1 + 1, value="=" + eng_ref(k), font=Font(name=FONT, size=20, bold=True, color=NAVY),
              fill=FILL_CARD, fmt=fmt, align=Alignment(horizontal="center", vertical="center"))
    merge(ws, 13, c1, 13, c1 + 1, value=sub, font=Font(name=FONT, size=8, italic=True, color=GREY),
          fill=FILL_CARD, align=Alignment(horizontal="center", vertical="top"))
    for col in (c1, c1 + 1):
        ws.cell(row=10, column=col).border = Border(top=Side(style="thick", color=BLUE))
    rng_v = f"{L(c1)}11:{L(c1 + 1)}12"
    if crit_cell:
        ws.conditional_formatting.add(rng_v, FormulaRule(formula=[f'{crit_cell}="✓"'], font=Font(color=GREEN)))
        ws.conditional_formatting.add(rng_v, FormulaRule(formula=[f'{crit_cell}="✗"'], font=Font(color=RED)))
    else:
        ws.conditional_formatting.add(rng_v, FormulaRule(formula=[f'{L(c1)}11<0'], font=Font(color=RED)))
        ws.conditional_formatting.add(rng_v, FormulaRule(formula=[f'{L(c1)}11>0'], font=Font(color=GREEN)))
ws.row_dimensions[10].height = 18
ws.row_dimensions[11].height = 20
ws.row_dimensions[12].height = 20

# --- fontes e aplicações (direita, linhas 15-22)
section(ws, 15, "FONTES DE FINANCIAMENTO", 5, start_col=9)
header_row(ws, 16, ["Fonte", None, "Montante", "% do total", "Entidade"], start_col=9)
ws.merge_cells("I16:J16")
FS0 = 17
src = [("Dívida bancária", "=" + eng_ref("k_div"), "=BancoNome"),
       ("Subsídio a fundo perdido", "=" + eng_ref("k_sub"), "=SubNome"),
       ("Investidor financeiro", "=" + eng_ref("k_inv"), "=InvNome"),
       ("Capital próprio do promotor", f"=K{FS0 + 4}-SUM(K{FS0}:K{FS0 + 2})", "=PromotorNome")]
for i, (lab, f, ent) in enumerate(src):
    rr = FS0 + i
    merge(ws, rr, 9, rr, 10, value=lab, border=B_ROW)
    c = ws.cell(row=rr, column=11, value=f)
    c.number_format = MEUR
    c.border = B_ROW
    c = ws.cell(row=rr, column=12, value=f"=IF($K${FS0 + 4}>0,K{rr}/$K${FS0 + 4},0)")
    c.number_format = PCT
    c.border = B_ROW
    c = ws.cell(row=rr, column=13, value=ent)
    c.font = Font(name=FONT, size=8, color=GREY)
    c.border = B_ROW
    c.alignment = Alignment(shrink_to_fit=True)
rr = FS0 + 4
merge(ws, rr, 9, rr, 10, value="Total = investimento", font=F_BOLD, border=B_TOT)
c = ws.cell(row=rr, column=11, value="=" + eng_ref("k_cxt"))
c.number_format = MEUR
c.font = F_BOLD
c.border = B_TOT
for col in (12, 13):
    ws.cell(row=rr, column=col).border = B_TOT
ws.cell(row=rr, column=12, value=f"=SUM(L{FS0}:L{FS0 + 3})").number_format = PCT
ws.conditional_formatting.add(f"L{FS0}:L{FS0 + 3}",
                              DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                          color="9DC3E6", showValue=True))
rr += 1
merge(ws, rr, 9, rr, 10, value="Exposição máxima do promotor", font=F_SUB)
c = ws.cell(row=rr, column=11, value="=" + eng_ref("k_eqmax"))
c.number_format = MEUR
c.font = F_SUB
ws.cell(row=rr, column=12, value="inclui adiantamento do subsídio").font = Font(name=FONT, size=7, italic=True,
                                                                                  color=GREY)

# --- comparação de cenários (esquerda, 24+)
section(ws, 24, "COMPARAÇÃO DE CENÁRIOS", 6, start_col=2)
header_row(ws, 25, ["Indicador", None, None] + SCEN, start_col=2)
ws.merge_cells("B25:D25")
comp = [("TIR do projeto", "k_tirp", PCT), ("VAL do projeto", "k_valp", MEUR), ("TIR do promotor", "k_tirs", PCT),
        ("VAL do promotor", "k_vals", MEUR), ("Múltiplo do promotor (MOIC)", "k_moic", MULT),
        ("Payback do promotor", "k_pbs", ANOS), ("DSCR mínimo", "k_dmin", MULT),
        ("Exposição máxima do promotor", "k_eqmax", MEUR), ("Receita — ano 5", "k_rs", MEUR),
        ("EBITDA — ano 5", "k_es", MEUR), ("Investimento total", "k_cxt", MEUR)]
rr = 25
for lab, k, fmt in comp:
    rr += 1
    merge(ws, rr, 2, rr, 4, value=lab, border=B_ROW)
    for j, w in enumerate(ws_eng):
        c = ws.cell(row=rr, column=5 + j, value=f"='{w.title}'!$C${ENG[k]}")
        c.number_format = fmt
        c.border = B_ROW
        c.alignment = Alignment(horizontal="center")
ws.conditional_formatting.add(f"E26:G{rr}", FormulaRule(formula=["E$25=CenarioAtivo"],
                                                        fill=PatternFill("solid", fgColor="DDEBF7", bgColor="DDEBF7"),
                                                        font=Font(bold=True)))
COMP_END = rr

# --- indicadores operacionais (direita, 24+)
section(ws, 24, "INDICADORES DO PROJETO (cenário ativo)", 5, start_col=9)
header_row(ws, 25, ["Indicador", None, None, "Valor", None], start_col=9)
ws.merge_cells("I25:K25")
ws.merge_cells("L25:M25")
key_list = [("Investimento total", "k_cxt", MEUR), ("Dívida / investimento (LTC)", "k_ltc", PCT),
            ("TIR do capital próprio total", "k_tire", PCT), ("TIR do investidor financeiro", "k_tiri", PCT),
            ("Payback do projeto", "k_pbj", ANOS), ("DSCR médio", "k_dmed", MULT),
            ("Receita — 1.º ano de operação", "k_r1", MEUR), ("Receita — ano 5", "k_rs", MEUR),
            ("Margem EBITDA — ano 5", "k_em", PCT), ("Yield on cost (NOI ano 5)", "k_yoc", PCT),
            ("Valor de saída (bruto)", "k_tv", MEUR), ("Receita de venda de unidades", "k_vend", MEUR),
            ("Lucro líquido do promotor", "k_lucro", MEUR)]
rr = 25
for lab, k, fmt in key_list:
    rr += 1
    merge(ws, rr, 9, rr, 11, value=lab, border=B_ROW)
    merge(ws, rr, 12, rr, 13, value="=" + eng_ref(k), fmt=fmt, font=F_BOLD, border=B_ROW,
          align=Alignment(horizontal="right"))
KPI_END = rr
DASH_CHART_ROW = max(COMP_END, KPI_END) + 2
section(ws, DASH_CHART_ROW, "EVOLUÇÃO E COMPOSIÇÃO (cenário ativo)", 12, start_col=2)

# dados auxiliares para o gráfico de aplicações
ws["P9"] = "Dados dos gráficos"
ws["P9"].font = Font(name=FONT, size=8, bold=True, color=GREY)
AP0 = 10
for i, (lab, k) in enumerate((("Terreno / aquisição", "k_ter"), ("Construção", "k_con"), ("FF&E", "k_ffe"),
                              ("Gastos amortizáveis", "k_gas"), ("Fundo de maneio", "k_fdm"))):
    ws.cell(row=AP0 + i, column=16, value=lab).font = Font(name=FONT, size=8, color=GREY)
    c = ws.cell(row=AP0 + i, column=17, value="=" + eng_ref(k))
    c.number_format = EUR
    c.font = Font(name=FONT, size=8, color=GREY)

# ================================================================ ESTRUTURA
ws = ws_est
ws.sheet_view.showGridLines = False
widths(ws, {"A": 2, **{L(c): 5.4 for c in range(2, 26)}, "Z": 2})
banner(ws, "Estrutura de investimento",
       "Árvore de financiadores, veículos e aplicações  ·  mapa de alocação de fundos  ·  retorno por financiador "
       "(cenário ativo)", 25, nav_span=2, start_col=1)
LINE = Side(style="medium", color="8497B0")


def vline(r1, r2, col):
    for r_ in range(r1, r2 + 1):
        c = ws.cell(row=r_, column=col)
        b = c.border
        c.border = Border(left=LINE, right=b.right, top=b.top, bottom=b.bottom)


def hline(row, c1, c2):
    for c_ in range(c1, c2 + 1):
        c = ws.cell(row=row, column=c_)
        b = c.border
        c.border = Border(left=b.left, right=b.right, top=b.top, bottom=LINE)


def lbl(row, col, value, fmt=None, bold=False, align="left", c2=None):
    c = merge(ws, row, col, row, c2 or col, value=value,
              font=Font(name=FONT, size=8 if bold else 7, bold=bold, italic=not bold, color=NAVY if bold else GREY),
              fmt=fmt, align=Alignment(horizontal=align, vertical="center"))
    return c


def tbox(r, c, header, name_f, amount_f, detail_f, color, amount_fmt=MEUR, name_fmt=None):
    fill_h = PatternFill("solid", fgColor=color)
    merge(ws, r, c, r, c + 3, value=header, font=Font(name=FONT, size=7, bold=True, color="FFFFFF"), fill=fill_h,
          align=Alignment(horizontal="center", vertical="center"))
    merge(ws, r + 1, c, r + 1, c + 3, value=name_f, font=Font(name=FONT, size=9, bold=True, color=NAVY),
          fill=FILL_CARD, fmt=name_fmt, align=Alignment(horizontal="center", vertical="center", shrink_to_fit=True))
    merge(ws, r + 2, c, r + 2, c + 3, value=amount_f, font=Font(name=FONT, size=13, bold=True, color=color),
          fill=FILL_CARD, fmt=amount_fmt, align=Alignment(horizontal="center", vertical="center"))
    merge(ws, r + 3, c, r + 3, c + 3, value=detail_f, font=Font(name=FONT, size=7, italic=True, color=GREY),
          fill=FILL_CARD, align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    outline_box(ws, r, c, r + 3, c + 3, color)
    amt = f"${L(c)}${r + 2}"
    ws.conditional_formatting.add(f"{L(c)}{r}:{L(c + 3)}{r}", FormulaRule(
        formula=[f"N({amt})=0"], fill=PatternFill("solid", fgColor="BFBFBF", bgColor="BFBFBF")))
    ws.conditional_formatting.add(f"{L(c)}{r + 1}:{L(c + 3)}{r + 3}", FormulaRule(
        formula=[f"N({amt})=0"], font=Font(color="A6A6A6")))


section(ws, 5, "ÁRVORE DE INVESTIMENTO E FLUXOS DE CAPITAL (cenário ativo)", 24, start_col=2)
for r_ in (9, 16, 23, 29):
    ws.row_dimensions[r_].height = 22
for r_ in (10, 17, 24, 30):
    ws.row_dimensions[r_].height = 26
C_PROM, C_BANK, C_SPV, C_PROJ, C_SUB, C_INV = "1B2A41", "2E75B6", "596780", "9C7A16", GREEN, "7B3F98"
# nível 1 — promotor e banco
tbox(7, 6, "PROMOTOR", "=PromotorNome", "=Est_Prom",
     '="Capital "&ROUND(Est_Cap/1000000,2)&" M€  ·  suprimentos "&ROUND(Est_Sup/1000000,2)&" M€"', C_PROM)
tbox(7, 18, "FINANCIAMENTO BANCÁRIO", "=BancoNome", "=" + eng_ref("k_div"),
     '="Juro "&ROUND(Juro*100,2)&"%  ·  "&Prazo&" anos  ·  carência "&Carencia', C_BANK)
vline(11, 12, 8)
vline(11, 12, 20)
hline(12, 8, 19)
vline(13, 13, 14)
lbl(11, 9, '=IF(UsarSPV="Sim","Capital + suprimentos","Direto à sociedade do projeto")', c2=13)
lbl(12, 9, "=Est_Prom", fmt=MEUR, bold=True, c2=11)
lbl(11, 21, '=IF(Est_DivSPV>0,"Empréstimo à SPV","Empréstimo à sociedade do projeto")', c2=25)
lbl(12, 21, "=" + eng_ref("k_div"), fmt=MEUR, bold=True, c2=23)
# nível 2 — SPV
tbox(14, 12, "VEÍCULO INTERMÉDIO (SPV)", '=IF(UsarSPV="Sim",SPVNome,"Sem SPV")', "=Est_SPVin",
     '=IF(UsarSPV="Sim","Detida a "&ROUND(PartPromSPV*100,1)&"% por "&PromotorNome,"Fundos entram diretamente no projeto")',
     C_SPV)
vline(18, 20, 14)
lbl(18, 9, "Paga a aquisição", align="right", c2=13)
lbl(19, 9, "=Est_SPVaq", fmt=MEUR, bold=True, align="right", c2=13)
lbl(18, 15, "Transfere para o projeto", c2=19)
lbl(19, 15, "=Est_SPVtr", fmt=MEUR, bold=True, c2=17)
# nível 3 — subsídio, sociedade do projeto, investidor
tbox(21, 2, "SUBSÍDIO A FUNDO PERDIDO", "=SubNome", "=" + eng_ref("k_sub"),
     '="Recebido no ano "&SubsidioPer&"  ·  não reembolsável"', C_SUB)
tbox(21, 12, "SOCIEDADE DO PROJETO", "=NomeProjeto", "=" + eng_ref("k_cxt"),
     '="Detida a "&ROUND(PartSPVProj*100,1)&"% por "&IF(UsarSPV="Sim",SPVNome,PromotorNome)'
     '&IF(PartInvProj>0,"  ·  "&ROUND(PartInvProj*100,1)&"% "&InvNome,"")', C_PROJ)
tbox(21, 22, "INVESTIDOR FINANCEIRO", "=InvNome", "=" + eng_ref("k_inv"),
     '="Recompra no ano "&MIN(InvPer+InvRecompra,AnoSaida)&"  ·  "&ROUND(InvTaxa*100,2)&"% "&LOWER(InvRegime)', C_INV)
hline(22, 6, 11)
hline(22, 16, 21)
lbl(21, 6, "Fundo perdido", align="center", c2=11)
lbl(23, 6, "=" + eng_ref("k_sub"), fmt=MEUR, bold=True, align="center", c2=11)
lbl(21, 16, "Capital com recompra", align="center", c2=21)
lbl(23, 16, "=" + eng_ref("k_inv"), fmt=MEUR, bold=True, align="center", c2=21)
for col, ch_, al in ((11, "►", "right"), (16, "◄", "left")):
    c = ws.cell(row=22, column=col, value=ch_)
    c.font = Font(name=FONT, size=9, color="8497B0")
    c.alignment = Alignment(horizontal=al, vertical="bottom")
# nível 4 — aplicações
vline(25, 25, 14)
hline(25, 4, 23)
for col in (4, 9, 14, 19, 24):
    vline(26, 26, col)
uses = [("TERRENO / AQUISIÇÃO", "k_ter"), ("CONSTRUÇÃO", "k_con"), ("FF&E", "k_ffe"),
        ("GASTOS AMORTIZÁVEIS", "k_gas"), ("FUNDO DE MANEIO", "k_fdm")]
for i, (lab, k) in enumerate(uses):
    det = ('=IF(Est_SPVaq>0,"Paga pela SPV","Paga pela sociedade do projeto")' if k == "k_ter"
           else '="Paga pela sociedade do projeto"')
    tbox(27, 2 + 5 * i, lab, f"=IF({eng_ref('k_cxt')}>0,{eng_ref(k)}/{eng_ref('k_cxt')},0)", "=" + eng_ref(k), det,
         "44546A", name_fmt='0.0%" do investimento"')
lbl(32, 2, "Valores do cenário ativo (Dashboard). Caixas a cinzento = fonte ou aplicação sem montante.", c2=25)

# --- configuração
r = 34
section(ws, r, "CONFIGURAÇÃO DA ESTRUTURA", 24, start_col=2)
r += 1
dv_e_sn = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=False)
dv_e_via = DataValidation(type="list", formula1='"SPV,Sociedade do projeto"', allow_blank=False)
ws.add_data_validation(dv_e_sn)
ws.add_data_validation(dv_e_via)


def est_inp(label, nm, value, fmt="@", note="", dv=None, formula=False):
    global r
    merge(ws, r, 2, r, 12, value=label, border=B_ROW)
    c = merge(ws, r, 13, r, 17, value=value, fmt=fmt, align=Alignment(horizontal="center"))
    if formula:
        c.font = F_BOLD if formula == "calc" else F_LINK
        for col in range(13, 18):
            ws.cell(row=r, column=col).border = B_ROW
    else:
        for col in range(13, 18):
            set_input(ws.cell(row=r, column=col))
        c.font = F_INPUT
    if dv is not None:
        dv.add(c)
    merge(ws, r, 18, r, 25, value=note, font=Font(name=FONT, size=7, italic=True, color=GREY), border=B_ROW,
          align=Alignment(wrap_text=True, vertical="center"))
    if len(note) > 55:
        ws.row_dimensions[r].height = 24
    if nm:
        name(nm, ws, f"$M${r}")
    r += 1


est_inp("Nome do promotor", "PromotorNome", "Potentia Group")
est_inp("Utilizar veículo intermédio (SPV)?", "UsarSPV", "Sim", dv=dv_e_sn,
        note="Sim = promotor (e banco, se escolhido) financiam a SPV, que compra a sociedade.")
est_inp("Nome do veículo", "SPVNome", "SPV Cocoon (a constituir)")
est_inp("Participação do promotor na SPV", "PartPromSPV", 1.0, fmt=PCT)
est_inp("Participação da SPV / promotor na sociedade do projeto", "PartSPVProj", 1.0, fmt=PCT)
est_inp("Participação do investidor financeiro no capital do projeto", "PartInvProj", 0.0, fmt=PCT,
        note="Informativo. O original admite 100% da SPV após a recompra.")
est_inp("Capital social realizado pelo promotor", "CapSocial", 100000, fmt=EUR,
        note="O restante entra como suprimentos.")
est_inp("Nome do banco", "BancoNome", "Banco financiador (a definir)")
est_inp("Dívida bancária contratada por", "DividaVia", "SPV", dv=dv_e_via)
est_inp("Entidade do subsídio", "SubNome", "Portugal 2030")
est_inp("Investidor financeiro (definido em Pressupostos)", None, "=InvNome", formula="link")

r += 1
section(ws, r, "FLUXOS CALCULADOS (cenário ativo)", 24, start_col=2)
r += 1
inv_tot = eng_ref("k_cxt")
est_inp("Contribuição total do promotor (remanescente)", "Est_Prom",
        f"={inv_tot}-{eng_ref('k_div')}-{eng_ref('k_sub')}-{eng_ref('k_inv')}", fmt=EUR, formula="calc")
est_inp("   dos quais capital social", "Est_Cap", "=MAX(0,MIN(CapSocial,Est_Prom))", fmt=EUR, formula="calc")
est_inp("   dos quais suprimentos", "Est_Sup", "=Est_Prom-Est_Cap", fmt=EUR, formula="calc")
est_inp("Dívida bancária contratada pela SPV", "Est_DivSPV",
        f'=IF(AND(UsarSPV="Sim",DividaVia="SPV"),{eng_ref("k_div")},0)', fmt=EUR, formula="calc")
est_inp("Entradas de fundos na SPV", "Est_SPVin", '=IF(UsarSPV="Sim",Est_Prom+Est_DivSPV,0)', fmt=EUR,
        formula="calc")
est_inp("SPV — pagamento da aquisição", "Est_SPVaq", f'=IF(UsarSPV="Sim",MIN(Est_SPVin,{eng_ref("k_ter")}),0)',
        fmt=EUR, formula="calc", note="Até ao limite dos fundos que recebe.")
est_inp("SPV — transferência para a sociedade do projeto", "Est_SPVtr", "=Est_SPVin-Est_SPVaq", fmt=EUR,
        formula="calc", note="Suprimentos / prestações acessórias.")
est_inp("Controlo: fontes − aplicações (deve ser 0)", "Est_Ctrl",
        f"=Est_Prom+{eng_ref('k_div')}+{eng_ref('k_sub')}+{eng_ref('k_inv')}-{inv_tot}", fmt=EUR, formula="calc")

# --- mapa de alocação
r += 1
section(ws, r, "MAPA DE ALOCAÇÃO DE FUNDOS — QUEM FINANCIA CADA RUBRICA (cenário ativo)", 24, start_col=2)
r += 1
merge(ws, r, 2, r, 25, value="Preencha quanto de cada fonte financia cada rubrica. O promotor cobre o "
      "remanescente de cada linha. As rubricas vêm da folha Capex.", font=F_SUB)
r += 1
AL_COLS = [(2, 7, "Rubrica"), (8, 10, "Classe"), (11, 13, "Valor"), (14, 16, "Dívida bancária"),
           (17, 19, "Subsídio"), (20, 22, "Investidor"), (23, 25, "Promotor")]
for c1, c2, t in AL_COLS:
    merge(ws, r, c1, r, c2, value=t, font=F_HDR, fill=FILL_HDR,
          align=Alignment(horizontal="center", vertical="center", wrap_text=True))
ws.row_dimensions[r].height = 24
r += 1
AL0 = r
aloc_def = {0: (6020000, 0, 0), 1: (0, 4610000, 3390000), 2: (0, 0, 383000)}
for k in range(CX1 - CX0 + 1):
    i = CX0 + k
    rr = AL0 + k
    merge(ws, rr, 2, rr, 7, value=f"=IF(Capex!A{i}=\"\",\"\",Capex!A{i})", font=F_LINK, border=B_ROW)
    merge(ws, rr, 8, rr, 10, value=f"=IF(Capex!A{i}=\"\",\"\",Capex!B{i})",
          font=Font(name=FONT, size=8, color=GREY), border=B_ROW)
    merge(ws, rr, 11, rr, 13, value=(f"=N(Capex!C{i})*(1+IF(Capex!D{i}=\"Sim\",INDEX(S_AjCapex,1,CenIdx),0))"),
          fmt=EUR, border=B_ROW)
    dv_ = aloc_def.get(k, (None, None, None))
    for j, c1 in enumerate((14, 17, 20)):
        merge(ws, rr, c1, rr, c1 + 2, value=dv_[j], fmt=EUR)
        for col in range(c1, c1 + 3):
            set_input(ws.cell(row=rr, column=col))
    merge(ws, rr, 23, rr, 25, value=f"=K{rr}-N{rr}-Q{rr}-T{rr}", fmt=EUR, font=F_BOLD, border=B_ROW)
AL1 = AL0 + (CX1 - CX0)
ws.conditional_formatting.add(f"W{AL0}:Y{AL1}", CellIsRule(operator="lessThan", formula=["-0.5"], fill=FILL_BAD))
r = AL1 + 1
merge(ws, r, 2, r, 10, value="Total alocado", font=F_BOLD, border=B_TOT)
for c1 in (11, 14, 17, 20, 23):
    merge(ws, r, c1, r, c1 + 2, value=f"=SUM({L(c1)}{AL0}:{L(c1)}{AL1})", fmt=EUR, font=F_BOLD, border=B_TOT)
TOTR = r
r += 1
merge(ws, r, 2, r, 10, value="Montante disponível da fonte")
merge(ws, r, 11, r, 13, value=f"={inv_tot}", fmt=EUR)
for c1, f in ((14, eng_ref("k_div")), (17, eng_ref("k_sub")), (20, eng_ref("k_inv")), (23, "Est_Prom")):
    merge(ws, r, c1, r, c1 + 2, value="=" + f, fmt=EUR, font=F_LINK)
DISPR = r
r += 1
merge(ws, r, 2, r, 10, value="Por alocar (deve ser 0)", font=F_BOLD)
for c1 in (11, 14, 17, 20, 23):
    merge(ws, r, c1, r, c1 + 2, value=f"={L(c1)}{DISPR}-{L(c1)}{TOTR}", fmt=EUR, font=F_BOLD)
FALTR = r
ws.conditional_formatting.add(f"K{FALTR}:Y{FALTR}", FormulaRule(formula=[f"ABS(K{FALTR})>=1"], fill=FILL_WARN))
ws.conditional_formatting.add(f"K{FALTR}:Y{FALTR}", FormulaRule(formula=[f"ABS(K{FALTR})<1"], fill=FILL_OK))
name("Aloc_Falta", ws, f"$N${FALTR}:$V${FALTR}")
name("Aloc_PromMin", ws, f"$W${AL0}:$W${AL1}")

# --- retorno por financiador
r += 2
section(ws, r, "RETORNO POR FINANCIADOR (cenário ativo)", 24, start_col=2)
r += 1
RT_COLS = [(2, 7, "Financiador"), (8, 10, "Tipo"), (11, 13, "Investido"), (14, 16, "Recebido"),
           (17, 19, "Resultado"), (20, 22, "TIR"), (23, 25, "Múltiplo")]
for c1, c2, t in RT_COLS:
    merge(ws, r, c1, r, c2, value=t, font=F_HDR, fill=FILL_HDR, align=Alignment(horizontal="center", vertical="center"))
ws.row_dimensions[r].height = 22
ret = [("=PromotorNome", "Capital próprio", eng_ref("k_sneg"), eng_ref("k_spos"), eng_ref("k_tirs")),
       ("=InvNome", "Capital c/ recompra", eng_ref("k_inv"), eng_ref("k_invout"), eng_ref("k_tiri")),
       ("=BancoNome", "Dívida bancária", eng_ref("k_div"), eng_ref("k_brec"), eng_ref("k_tirb")),
       ("=SubNome", "Fundo perdido", eng_ref("k_sub"), "0", '"n.a."')]
for nm_, tipo, inv_, rec_, tir_ in ret:
    r += 1
    merge(ws, r, 2, r, 7, value=nm_, font=F_BOLD, border=B_ROW)
    merge(ws, r, 8, r, 10, value=tipo, font=Font(name=FONT, size=8, color=GREY), border=B_ROW)
    merge(ws, r, 11, r, 13, value="=" + inv_, fmt=EUR, border=B_ROW)
    merge(ws, r, 14, r, 16, value="=" + rec_, fmt=EUR, border=B_ROW)
    merge(ws, r, 17, r, 19, value=f"=N{r}-K{r}", fmt=EUR, font=F_BOLD, border=B_ROW)
    merge(ws, r, 20, r, 22, value="=" + tir_, fmt=PCT, font=F_BOLD, border=B_ROW,
          align=Alignment(horizontal="right"))
    merge(ws, r, 23, r, 25, value=f'=IF(K{r}>0,N{r}/K{r},"n.a.")', fmt=MULT, border=B_ROW,
          align=Alignment(horizontal="right"))
r += 1
merge(ws, r, 2, r, 25, value="Promotor: investido = soma dos cash flows negativos (inclui o adiantamento do subsídio e "
      "a recompra ao investidor); recebido = soma dos positivos (inclui o valor de saída).", font=F_SUB)
ws.freeze_panes = "A4"

# ================================================================ RELATÓRIO
ws = ws_rep
ws.sheet_view.showGridLines = False
widths(ws, {"A": 44, "B": 10, "C": 15})
for col in PCOLS:
    ws.column_dimensions[col].width = 12.5
banner(ws, '="Relatório  ·  "&NomeProjeto&"  ·  cenário "&CenarioAtivo',
       "Exploração, resultados, investimento, dívida, saída e cash flows ano a ano. Use os botões − / + à "
       "esquerda para recolher o detalhe. Anos após a saída a cinzento.", C0 + NP - 1)
src_ws = ws_eng[0]
last_row = max(v for k, v in ENG.items() if not k.startswith("k_"))
REP_OFF = 4 - HDR_ROW
for r_ in range(HDR_ROW, last_row + 1):
    for col in range(1, C0 + NP):
        s_ = src_ws.cell(row=r_, column=col)
        d = ws.cell(row=r_ + REP_OFF, column=col)
        if s_.value is None:
            if s_.fill and s_.fill.fill_type:
                d.fill = copy(s_.fill)
                d.border = copy(s_.border)
            continue
        if col <= 2 or r_ == HDR_ROW:
            d.value = s_.value
        else:
            refs = ",".join(f"'{w.title}'!{L(col)}{r_}" for w in ws_eng)
            d.value = f"=CHOOSE(CenIdx,{refs})"
        d.font = copy(s_.font)
        d.fill = copy(s_.fill)
        d.border = copy(s_.border)
        d.number_format = s_.number_format
        d.alignment = copy(s_.alignment)
    ws.row_dimensions[r_ + REP_OFF].height = src_ws.row_dimensions[r_].height
for col in PCOLS:
    ws[f"{col}4"].number_format = '"Ano "0'
ws["C4"] = "Total"


def rep_row(key):
    return ENG[key] + REP_OFF


data_rng = f"D5:{LAST}{last_row + REP_OFF}"
ws.conditional_formatting.add(data_rng, FormulaRule(formula=[f"D${rep_row('ativo')}=0"],
                                                    font=Font(color="C0C0C0")))
ws.conditional_formatting.add(data_rng, FormulaRule(formula=[f"D${rep_row('saida')}=1"],
                                                    fill=PatternFill("solid", fgColor="FFF4E0", bgColor="FFF4E0")))
for a_, b_ in GROUPS:
    ws.row_dimensions.group(rep_row(a_), rep_row(b_), outline_level=1, hidden=False)
ws.sheet_properties.outlinePr.summaryBelow = True
ws.freeze_panes = "D5"

# ---- gráficos (no Dashboard, com dados do Relatório)
ch = LineChart()
ch.title = "Receita, EBITDA e NOI"
ch.height, ch.width = 7.5, 14.6
for k in ("rtot", "ebitda", "noi"):
    ch.add_data(Reference(ws_rep, min_col=C0 + 1, max_col=C0 + NP - 1, min_row=rep_row(k)), from_rows=True,
                titles_from_data=False)
for s_, t in zip(ch.series, ("Receita", "EBITDA", "NOI")):
    s_.tx = SeriesLabel(v=t)
ch.set_categories(Reference(ws_rep, min_col=C0 + 1, max_col=C0 + NP - 1, min_row=4))
style_chart(ch, [NAVY, BLUE, GOLD])
ws_dash.add_chart(ch, f"B{DASH_CHART_ROW + 2}")

bc = BarChart()
bc.title = "Cash flow do promotor"
bc.height, bc.width = 7.5, 14.6
bc.add_data(Reference(ws_rep, min_col=C0, max_col=C0 + NP - 1, min_row=rep_row("fcfs")), from_rows=True,
            titles_from_data=False)
bc.series[0].tx = SeriesLabel(v="Anual")
bc.set_categories(Reference(ws_rep, min_col=C0, max_col=C0 + NP - 1, min_row=4))
style_chart(bc, [BLUE], kind="bar")
lc = LineChart()
lc.add_data(Reference(ws_rep, min_col=C0, max_col=C0 + NP - 1, min_row=rep_row("fcfsa")), from_rows=True,
            titles_from_data=False)
lc.series[0].tx = SeriesLabel(v="Acumulado")
lc.series[0].graphicalProperties.line.solidFill = GOLD
lc.series[0].graphicalProperties.line.width = 28575
lc.series[0].smooth = False
bc += lc
ws_dash.add_chart(bc, f"H{DASH_CHART_ROW + 2}")


def pie(title, labels_ref, data_ref, anchor, colors):
    pc = PieChart()
    pc.title = title
    pc.height, pc.width = 7.5, 14.6
    pc.add_data(data_ref, titles_from_data=False)
    pc.set_categories(labels_ref)
    pc.dataLabels = DataLabelList()
    pc.dataLabels.showPercent = True
    pc.dataLabels.showVal = False
    pc.dataLabels.showCatName = False
    pc.dataLabels.showSerName = False
    pc.dataLabels.showLeaderLines = False
    pc.legend.position = "r"
    from openpyxl.chart.marker import DataPoint
    for i, col in enumerate(colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = col
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        pc.series[0].dPt.append(pt)
    ws_dash.add_chart(pc, anchor)


pie("Fontes de financiamento", Reference(ws_dash, min_col=9, min_row=FS0, max_row=FS0 + 3),
    Reference(ws_dash, min_col=11, min_row=FS0, max_row=FS0 + 3), f"B{DASH_CHART_ROW + 18}",
    [C_BANK, C_SUB, C_INV, C_PROM])
pie("Aplicações (investimento)", Reference(ws_dash, min_col=16, min_row=AP0, max_row=AP0 + 4),
    Reference(ws_dash, min_col=17, min_row=AP0, max_row=AP0 + 4), f"H{DASH_CHART_ROW + 18}",
    [NAVY, BLUE, GOLD, "8497B0", "BFBFBF"])

# ================================================================ CHECKS
ws = ws_chk
ws.sheet_view.showGridLines = False
banner(ws, "Verificações do modelo",
       "Controlos automáticos de coerência dos pressupostos e da estrutura de cálculo. ERRO = corrigir; "
       "AVISO = confirmar.", 4)
ws["A5"] = "Estado global"
ws["A5"].font = Font(name=FONT, size=11, bold=True, color=NAVY)
widths(ws, {"A": 60, "B": 16, "C": 12, "D": 70})
header_row(ws, 7, ["Verificação", "Valor", "Estado", "O que fazer se falhar"])
eng_b = ws_eng[0].title
checks = [
    ("Período de abertura ≥ 0 e < período de saída", "=AnoSaida-AnoAbertura",
     "=IF(AND(AnoAbertura>=0,AnoSaida>AnoAbertura),\"OK\",\"ERRO\")", "Corrigir os períodos na secção 1."),
    ("Período de saída ≤ 20", "=AnoSaida", "=IF(AnoSaida<=20,\"OK\",\"ERRO\")", "O modelo cobre no máximo 20 anos."),
    ("ADR do 1.º ano de operação > 0", "=INDEX(ADR_tab,1)", "=IF(INDEX(ADR_tab,1)>0,\"OK\",\"ERRO\")",
     "Preencher o ADR do ano 1."),
    ("Ocupações entre 0% e 100% (todos os cenários)", "=MAX(S_Ocup1,S_OcupMax)",
     "=IF(AND(MIN(S_Ocup1,S_OcupMax)>=0,MAX(S_Ocup1,S_OcupMax)<=1),\"OK\",\"ERRO\")", "Rever a secção 2."),
    ("Ocupação do ano 1 ≤ ocupação máxima", "=SUMPRODUCT(--(S_Ocup1>S_OcupMax))",
     "=IF(SUMPRODUCT(--(S_Ocup1>S_OcupMax))=0,\"OK\",\"AVISO\")", "A ocupação máxima limita a do ano 1."),
    ("Unidades vendidas ≤ unidades totais", "=SUM(Vendas_tab)",
     "=IF(SUM(Vendas_tab)<=Unidades,\"OK\",\"ERRO\")", "Reduzir as vendas na secção 6."),
    ("Todas as rubricas de capex têm classe válida", "=Capex_SemClasse",
     "=IF(ABS(Capex_SemClasse)<1,\"OK\",\"ERRO\")", "Escolher a classe na folha Capex (lista)."),
    ("Rubricas de capex com n.º de períodos ≥ 1",
     "=SUMPRODUCT((Capex_Valor<>0)*(Capex_N<1))",
     "=IF(SUMPRODUCT((Capex_Valor<>0)*(Capex_N<1))=0,\"OK\",\"ERRO\")", "Preencher a coluna N.º períodos."),
    ("Capex do motor = total da folha Capex (cenário Base)",
     f"='{eng_b}'!$C${ENG['k_cxt']}-Capex_Total-INDEX(S_AjCapex,1)*SUM(Capex_VAj)",
     f"=IF(ABS('{eng_b}'!$C${ENG['k_cxt']}-Capex_Total-INDEX(S_AjCapex,1)*SUM(Capex_VAj))<1,\"OK\",\"ERRO\")",
     "Algum investimento está fora do horizonte de análise."),
    ("FSE: % fixo entre 0% e 100%", "=MAX('FSE'!C5:C19)",
     "=IF(AND(MIN('FSE'!C5:C19)>=0,MAX('FSE'!C5:C19)<=1),\"OK\",\"ERRO\")", "Corrigir a coluna % fixo."),
    ("Quadro de pessoal com FTE > 0", "=FTE_total", "=IF(FTE_total>0,\"OK\",\"AVISO\")", "Preencher a folha Pessoal."),
    ("Carência inferior ao prazo da dívida", "=Prazo-Carencia",
     "=IF(OR(Divida=0,Carencia<Prazo),\"OK\",\"ERRO\")", "Rever a secção 8."),
    ("Dívida totalmente reembolsada após a saída (todos os cenários)",
     "=" + "+".join(f"ABS('{w.title}'!$C${ENG['k_dfim']})" for w in ws_eng),
     "=IF(" + "+".join(f"ABS('{w.title}'!$C${ENG['k_dfim']})" for w in ws_eng) + "<1,\"OK\",\"ERRO\")",
     "Período de desembolso deve estar dentro do horizonte."),
    ("Dívida ≤ investimento total", f"='{eng_b}'!$C${ENG['k_ltc']}",
     f"=IF('{eng_b}'!$C${ENG['k_ltc']}<=1,\"OK\",\"AVISO\")", "Dívida superior ao investimento — confirmar."),
    ("Subsídio recebido dentro do horizonte", "=SubsidioPer",
     "=IF(OR(Subsidio=0,SubsidioPer<=AnoSaida),\"OK\",\"ERRO\")", "Rever o período do subsídio."),
    ("Exit yield > 0 em todos os cenários", "=MIN(S_ExitYield)",
     "=IF(MIN(S_ExitYield)>0,\"OK\",\"AVISO\")", "Sem exit yield não há valor terminal (TIR subestimada)."),
    ("Cash flow do capital próprio reconcilia (Base)",
     f"=SUM('{eng_b}'!{FIRST}{ENG['fcfe']}:{LAST}{ENG['fcfe']})-SUM('{eng_b}'!{FIRST}{ENG['fcfp']}:{LAST}{ENG['fcfp']})"
     f"-SUM('{eng_b}'!{FIRST}{ENG['e_tax']}:{LAST}{ENG['e_tax']})-SUM('{eng_b}'!{FIRST}{ENG['ddes']}:{LAST}{ENG['ddes']})"
     f"+SUM('{eng_b}'!{FIRST}{ENG['djur']}:{LAST}{ENG['dsel']})+SUM('{eng_b}'!{FIRST}{ENG['dcom']}:{LAST}{ENG['dcom']})"
     f"+SUM('{eng_b}'!{FIRST}{ENG['damo']}:{LAST}{ENG['dreemb']})",
     None, "Erro estrutural — contactar quem mantém o modelo."),
    ("GOP reconcilia com receitas − custos (Base)",
     f"='{eng_b}'!$C${ENG['gop']}-('{eng_b}'!$C${ENG['rtot']}-'{eng_b}'!$C${ENG['cdir']}"
     f"-'{eng_b}'!$C${ENG['pay']}-'{eng_b}'!$C${ENG['fse']})",
     None, "Erro estrutural — contactar quem mantém o modelo."),
    ("Unidades em exploração nunca negativas (Base)",
     f"=MIN('{eng_b}'!{FIRST}{ENG['unid']}:{LAST}{ENG['unid']})",
     f"=IF(MIN('{eng_b}'!{FIRST}{ENG['unid']}:{LAST}{ENG['unid']})>=0,\"OK\",\"ERRO\")", ""),
    ("Estrutura: fontes = aplicações", "=Est_Ctrl", None, "Erro estrutural — contactar quem mantém o modelo."),
    ("Mapa de alocação: dívida, subsídio e investidor totalmente alocados", "=SUMPRODUCT(ABS(Aloc_Falta))",
     "=IF(SUMPRODUCT(ABS(Aloc_Falta))<1,\"OK\",\"AVISO\")", "Completar o mapa de alocação na folha Estrutura."),
    ("Mapa de alocação: nenhuma rubrica com parte do promotor negativa", "=MIN(Aloc_PromMin)",
     "=IF(MIN(Aloc_PromMin)>-1,\"OK\",\"AVISO\")", "Há fontes a financiar mais do que o valor da rubrica."),
    ("Capital social ≤ contribuição do promotor", "=CapSocial-Est_Prom",
     "=IF(CapSocial<=Est_Prom,\"OK\",\"AVISO\")", "Rever o capital social na folha Estrutura."),
]
for i, (lab, val, st, todo) in enumerate(checks):
    rr = 8 + i
    ws.cell(row=rr, column=1, value=lab).border = B_ROW
    c = ws.cell(row=rr, column=2, value=val)
    c.number_format = '#,##0.00;(#,##0.00);"-"'
    if st is None:
        st = f'=IF(ABS(B{rr})<1,"OK","ERRO")'
    c = ws.cell(row=rr, column=3, value=st)
    c.alignment = Alignment(horizontal="center")
    c.font = F_BOLD
    ws.cell(row=rr, column=4, value=todo).font = F_SUB
CHK_END = 8 + len(checks) - 1
ws["B5"] = (f'=IF(COUNTIF(C8:C{CHK_END},"ERRO")>0,"ERRO",IF(COUNTIF(C8:C{CHK_END},"AVISO")>0,"OK c/ avisos","OK"))')
ws["B5"].font = Font(name=FONT, size=11, bold=True)
ws["B5"].alignment = Alignment(horizontal="center")
name("ChecksStatus", ws, "$B$5")
for rng_ in (f"C8:C{CHK_END}", "B5"):
    ws.conditional_formatting.add(rng_, CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_OK))
    ws.conditional_formatting.add(rng_, CellIsRule(operator="equal", formula=['"ERRO"'], fill=FILL_BAD))
    ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'ISNUMBER(SEARCH("AVISO",{rng_.split(":")[0]}))'],
                                                    fill=FILL_WARN))
ws.conditional_formatting.add("B5", CellIsRule(operator="equal", formula=['"OK c/ avisos"'], fill=FILL_WARN))
ws.freeze_panes = "A8"

# ================================================================ INÍCIO (instruções)
ws = ws0
ws.sheet_view.showGridLines = False
widths(ws, {"A": 2, "B": 26, "C": 100})
banner(ws, "Modelo de avaliação de investimentos",
       "Rentabilidade e viabilidade de projetos de turismo e imobiliário  ·  projeção anual até 20 anos  ·  "
       "3 cenários", 2, start_col=2)
rows_txt = [
    ("COMO USAR", None),
    ("1. Pressupostos", "Preencha todas as células a azul (fundo amarelo). Comece pelas secções 1–3 e 8."),
    ("2. Capex", "Liste as rubricas de investimento, a classe (lista), o período e em quantos períodos é pago."),
    ("3. Pessoal / FSE", "Ajuste o quadro de pessoal e os FSE da unidade (valores do 1.º ano de operação)."),
    ("4. Estrutura", "Defina o promotor, a SPV, o banco e o mapa de alocação (que fonte financia cada rubrica)."),
    ("5. Dashboard", "Escolha o cenário (célula amarela) e leia o veredicto, os KPI e a comparação de cenários."),
    ("6. Checks", "Confirme que o estado é OK. Um ERRO significa que algum pressuposto é incoerente."),
    ("7. Relatório", "P&L, cash flow, dívida e saída do cenário ativo, ano a ano — pronto para imprimir."),
    ("", None),
    ("LEGENDA", None),
    ("Texto azul / fundo amarelo claro", "Input editável."),
    ("Fundo amarelo forte", "Pressuposto-chave com grande impacto — validar com cuidado."),
    ("Texto verde", "Valor ligado de outra folha."),
    ("Texto preto", "Fórmula — não editar."),
    ("Folhas M_ (cinzentas)", "Motores de cálculo, um por cenário. Não editar."),
    ("", None),
    ("LÓGICA DO MODELO", None),
    ("Receita", "Noites ocupadas × ADR + F&B por noite ocupada + outros proveitos."),
    ("GOP", "Receita − custos diretos (por noite / % receita) − pessoal − FSE (fixo + variável com ocupação)."),
    ("EBITDA / NOI", "GOP − management fees − shared services = EBITDA; − reserva FF&E − capex manutenção = NOI."),
    ("Resultado", "NOI + venda de unidades − amortizações − juros − IRC (com reporte de prejuízos, limite 65%)."),
    ("Cash flow do projeto", "NOI + vendas − capex + subsídio − IRC sem dívida + valor de saída (NOI / exit yield)."),
    ("Cash flow do promotor", "Projeto + dívida (desembolso − serviço) + investidor financeiro (entrada − recompra)."),
    ("Indicadores", "TIR e VAL do projeto e do promotor, MOIC, payback, DSCR, yield on cost, exposição máxima."),
    ("Veredicto", "VIÁVEL se cumprir todos os critérios da secção 10; COM RESERVAS se falhar 1; senão NÃO VIÁVEL."),
    ("", None),
    ("CONVENÇÕES", None),
    ("Períodos", "Anuais. Período 0 = investimento inicial. Fluxos no fim de cada período."),
    ("Saída", "No período de saída o ativo é vendido e a dívida remanescente reembolsada."),
    ("Simplificações", "Subsídio não tributável; reserva FF&E tratada como gasto; sem IVA nem saldo mínimo de caixa; "
                       "cash flows negativos após o período 0 são suprimentos adicionais do promotor."),
]
rr = 5
for a, b in rows_txt:
    if b is None and a:
        section(ws, rr, "", 3)
        ws.cell(row=rr, column=2, value=a).font = F_SEC
    else:
        ws.cell(row=rr, column=2, value=a).font = F_BOLD
        ws.cell(row=rr, column=3, value=b).font = F_BASE
        ws.cell(row=rr, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    rr += 1
set_input(ws["B15"])
ws["B15"].font = Font(name=FONT, size=10, bold=True, color="0000FF")
ws["B16"].fill = FILL_KEY
ws["B17"].font = Font(name=FONT, size=10, bold=True, color="008000")

# ================================================================ acabamentos
for w in wb.worksheets:
    for row in w.iter_rows():
        for c in row:
            f = c.font
            if f.name != FONT:
                c.font = Font(name=FONT, size=f.size or 10, bold=f.bold, italic=f.italic, color=f.color)
    w.sheet_view.zoomScale = 90
    w.sheet_view.showGridLines = False
ws0.sheet_properties.tabColor = NAVY
ws_dash.sheet_properties.tabColor = NAVY
ws_est.sheet_properties.tabColor = BLUE
for w in (ws_in, ws_cx, ws_p, ws_f):
    w.sheet_properties.tabColor = GOLD
ws_rep.sheet_properties.tabColor = GREEN
ws_chk.sheet_properties.tabColor = RED
ws_dash.print_area = f"A1:N{DASH_CHART_ROW + 34}"
for w in (ws_dash, ws_est, ws_rep, ws_in, ws_chk, ws_cx, ws_p, ws_f) + tuple(ws_eng):
    w.page_setup.orientation = "landscape"
    w.page_setup.paperSize = w.PAPERSIZE_A4
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.active = 1
wb.save(OUT)
print("Gerado:", OUT)
