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
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.workbook.properties import CalcProperties

OUT = sys.argv[1] if len(sys.argv) > 1 else "Modelo_Avaliacao_Investimentos.xlsx"

# ---------------------------------------------------------------- estilos
FONT = "Arial"
NAVY = "1F3864"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_TITLE = Font(name=FONT, size=16, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_SEC = Font(name=FONT, size=11, bold=True, color=NAVY)
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_TOT = PatternFill("solid", fgColor="F2F2F2")
FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_BAD = PatternFill("solid", fgColor="FFC7CE")
FILL_WARN = PatternFill("solid", fgColor="FFEB9C")
THIN = Side(style="thin", color="BFBFBF")
B_TOP = Border(top=Side(style="thin", color="000000"))
B_BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

EUR = '#,##0;(#,##0);"-"'
EUR2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.00"x";(0.00"x");"-"'
INT = '0;(0);"-"'
YEAR = '0'

wb = Workbook()
names = {}


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


def section(ws, row, text, ncols):
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).fill = FILL_SEC
    c = ws.cell(row=row, column=1, value=text)
    c.font = F_SEC


def set_input(c, key=False):
    c.font = F_INPUT
    c.fill = FILL_KEY if key else FILL_IN
    c.border = B_BOX


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


def default_font(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size if c.font else 10,
                              bold=c.font.bold if c.font else False,
                              italic=c.font.italic if c.font else False,
                              color=c.font.color if c.font else None)


# ================================================================ INÍCIO
ws0 = wb.active
ws0.title = "Início"
ws_dash = wb.create_sheet("Dashboard")
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
ws["A1"] = "PRESSUPOSTOS DO PROJETO"
ws["A1"].font = F_TITLE
ws["A2"] = ("Preencha apenas as células a azul com fundo amarelo. Fundo amarelo forte = pressuposto-chave "
            "a validar. Os valores por defeito são os do projeto Cocoon (caso de teste).")
ws["A2"].font = F_SUB
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
ws.freeze_panes = "B4"
IN_LAST = r

# ================================================================ CAPEX
ws = ws_cx
CLASSES = ["Terreno / aquisição", "Construção", "FF&E", "Gastos amortizáveis", "Fundo de maneio"]
ws["A1"] = "PLANO DE INVESTIMENTO (CAPEX)"
ws["A1"].font = F_TITLE
ws["A2"] = ("Uma linha por rubrica. Classe define a amortização: Terreno não amortiza; Fundo de maneio é "
            "recuperado na saída. O valor é distribuído em partes iguais pelo n.º de períodos.")
ws["A2"].font = F_SUB
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
ws["A1"] = "QUADRO DE PESSOAL E SHARED SERVICES"
ws["A1"].font = F_TITLE
ws["A2"] = "Custo anual = FTE × salário × meses × (1 + bónus) × (1 + encargos). Valores do ano 1 de operação."
ws["A2"].font = F_SUB
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
ws["A1"] = "FORNECIMENTOS E SERVIÇOS EXTERNOS (FSE)"
ws["A1"].font = F_TITLE
ws["A2"] = ("Custo anual = base × [% fixo + % variável × (ocupação / ocupação de referência)] × inflação. "
            "Valores-base a preços do ano 1 de operação.")
ws["A2"].font = F_SUB
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


def build_engine(ws, idx, scen):
    rows = {}
    ws["A1"] = f"MOTOR DE CÁLCULO — CENÁRIO {scen.upper()}"
    ws["A1"].font = F_TITLE
    ws["A2"] = "Folha 100% calculada — não editar. Altere os valores na folha Pressupostos."
    ws["A2"].font = F_SUB
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
    ws.sheet_properties.tabColor = "808080"
    return rows


ENG = None
for i, (wse, s) in enumerate(zip(ws_eng, SCEN)):
    ENG = build_engine(wse, i + 1, s)

# ================================================================ DASHBOARD
ws = ws_dash
ws.sheet_view.showGridLines = False
widths(ws, {"A": 2, "B": 40, "C": 18, "D": 18, "E": 18, "F": 4, "G": 34, "H": 16, "I": 12, "J": 12})
ws["B1"] = "=\"AVALIAÇÃO DE INVESTIMENTO — \"&UPPER(NomeProjeto)"
ws["B1"].font = F_TITLE
ws["B2"] = '=Localizacao&"  |  "&TipoAtivo&"  |  Análise: "&Analista'
ws["B2"].font = F_SUB
ws["B4"] = "Cenário em análise  ▶"
ws["B4"].font = F_BOLD
ws["C4"] = "Base"
set_input(ws["C4"], key=True)
ws["C4"].font = Font(name=FONT, size=11, bold=True, color="0000FF")
dv_sc = DataValidation(type="list", formula1="ListaCenarios", allow_blank=False)
ws.add_data_validation(dv_sc)
dv_sc.add(ws["C4"])
name("CenarioAtivo", ws, "$C$4")
ws["D4"] = "=MATCH(CenarioAtivo,ListaCenarios,0)"
ws["D4"].font = Font(name=FONT, size=8, color="A6A6A6")
name("CenIdx", ws, "$D$4")
ws["G4"] = "Estado do modelo (Checks)"
ws["G4"].font = F_BOLD
ws["H4"] = "=ChecksStatus"
ws["H4"].font = F_BOLD
ws["H4"].alignment = Alignment(horizontal="center")


def eng_ref(key, col="C"):
    """Valor do motor do cenário ativo."""
    return "CHOOSE(CenIdx," + ",".join(f"'{w.title}'!${col}${ENG[key]}" for w in ws_eng) + ")"


# veredicto
ws["B6"] = "VEREDICTO"
ws["B6"].font = F_SEC
ws.merge_cells("C6:E6")
ws["C6"] = ('=IF(COUNTIF($E$10:$E$14,"✗")=0,"VIÁVEL",IF(COUNTIF($E$10:$E$14,"✗")<=1,'
            '"VIÁVEL COM RESERVAS","NÃO VIÁVEL"))')
ws["C6"].font = Font(name=FONT, size=14, bold=True)
ws["C6"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[6].height = 28
ws.conditional_formatting.add("C6:E6", FormulaRule(formula=['$C$6="VIÁVEL"'], fill=FILL_OK))
ws.conditional_formatting.add("C6:E6", FormulaRule(formula=['$C$6="VIÁVEL COM RESERVAS"'], fill=FILL_WARN))
ws.conditional_formatting.add("C6:E6", FormulaRule(formula=['$C$6="NÃO VIÁVEL"'], fill=FILL_BAD))

header_row(ws, 9, ["Critério", "Resultado", "Mínimo / máximo", "Cumpre?"], start_col=2)
crit = [
    ("TIR do projeto", "k_tirp", "Crit_TIRproj", PCT, ">="),
    ("TIR do promotor", "k_tirs", "Crit_TIRprom", PCT, ">="),
    ("DSCR mínimo", "k_dmin", "Crit_DSCR", MULT, ">="),
    ("Payback do promotor (anos)", "k_pbs", "Crit_Payback", INT, "<="),
    ("Múltiplo de capital (MOIC)", "k_moic", "Crit_MOIC", MULT, ">="),
]
for i, (lab, k, cn, fmt, op) in enumerate(crit):
    rr = 10 + i
    ws.cell(row=rr, column=2, value=lab)
    c = ws.cell(row=rr, column=3, value="=" + eng_ref(k))
    c.number_format = fmt
    c = ws.cell(row=rr, column=4, value=f"={cn}")
    c.number_format = fmt
    c.font = F_LINK
    if k == "k_dmin":
        # sem dívida não há DSCR: critério cumprido
        f = f'=IF(ISNUMBER(C{rr}),IF(C{rr}{op}D{rr},"✓","✗"),IF(Divida=0,"✓","✗"))'
    else:
        f = f'=IF(ISNUMBER(C{rr}),IF(C{rr}{op}D{rr},"✓","✗"),"✗")'
    c = ws.cell(row=rr, column=5, value=f)
    c.alignment = Alignment(horizontal="center")
    c.font = F_BOLD
    for col in range(2, 6):
        ws.cell(row=rr, column=col).border = B_BOX
ws.conditional_formatting.add("E10:E14", CellIsRule(operator="equal", formula=['"✓"'], fill=FILL_OK))
ws.conditional_formatting.add("E10:E14", CellIsRule(operator="equal", formula=['"✗"'], fill=FILL_BAD))

# indicadores principais (cenário ativo) à direita
ws["G6"] = "INDICADORES-CHAVE (cenário ativo)"
ws["G6"].font = F_SEC
key_list = [("Investimento total", "k_cxt", EUR), ("Dívida bancária", "k_div", EUR),
            ("LTC (dívida / investimento)", "k_ltc", PCT), ("Subsídio a fundo perdido", "k_sub", EUR),
            ("Investidor financeiro", "k_inv", EUR), ("Exposição máxima do promotor", "k_eqmax", EUR),
            ("VAL do projeto", "k_valp", EUR), ("VAL do promotor", "k_vals", EUR),
            ("TIR do capital próprio total", "k_tire", PCT), ("TIR do investidor financeiro", "k_tiri", PCT),
            ("Payback do projeto (anos)", "k_pbj", INT), ("DSCR médio", "k_dmed", MULT),
            ("Receita — 1.º ano", "k_r1", EUR), ("Receita — ano 5", "k_rs", EUR),
            ("Margem EBITDA — ano 5", "k_em", PCT), ("Yield on cost (NOI ano 5)", "k_yoc", PCT),
            ("Valor de saída (bruto)", "k_tv", EUR), ("Receita de venda de unidades", "k_vend", EUR),
            ("Lucro líquido do promotor", "k_lucro", EUR)]
for i, (lab, k, fmt) in enumerate(key_list):
    rr = 7 + i
    ws.cell(row=rr, column=7, value=lab).border = B_BOX
    c = ws.cell(row=rr, column=8, value="=" + eng_ref(k))
    c.number_format = fmt
    c.border = B_BOX
    c.font = F_BOLD

# comparação de cenários
rr = 17
ws.cell(row=rr, column=2, value="COMPARAÇÃO DE CENÁRIOS").font = F_SEC
rr += 1
header_row(ws, rr, ["Indicador"] + SCEN, start_col=2)
comp = [("TIR do projeto", "k_tirp", PCT), ("VAL do projeto", "k_valp", EUR), ("TIR do promotor", "k_tirs", PCT),
        ("VAL do promotor", "k_vals", EUR), ("MOIC do promotor", "k_moic", MULT),
        ("Payback do promotor (anos)", "k_pbs", INT), ("DSCR mínimo", "k_dmin", MULT),
        ("Exposição máxima do promotor", "k_eqmax", EUR), ("Receita — ano 5", "k_rs", EUR),
        ("EBITDA — ano 5", "k_es", EUR), ("Investimento total", "k_cxt", EUR)]
for lab, k, fmt in comp:
    rr += 1
    ws.cell(row=rr, column=2, value=lab).border = B_BOX
    for j, w in enumerate(ws_eng):
        c = ws.cell(row=rr, column=3 + j, value=f"='{w.title}'!$C${ENG[k]}")
        c.number_format = fmt
        c.border = B_BOX
COMP_END = rr

# fontes e aplicações (período 0 + subsídio)
rr = COMP_END + 2
ws.cell(row=rr, column=2, value="FONTES E APLICAÇÕES (cenário ativo)").font = F_SEC
rr += 1
header_row(ws, rr, ["Fonte", "Montante (€)", "% do total"], start_col=2)
s0 = rr + 1
src = [("Dívida bancária", "=" + eng_ref("k_div")), ("Subsídio a fundo perdido", "=" + eng_ref("k_sub")),
       ("Investidor financeiro", "=" + eng_ref("k_inv")),
       ("Capital próprio do promotor", f"=C{s0 + 4}-SUM(C{s0}:C{s0 + 2})")]
for lab, f in src:
    rr += 1
    ws.cell(row=rr, column=2, value=lab).border = B_BOX
    c = ws.cell(row=rr, column=3, value=f)
    c.number_format = EUR
    c.border = B_BOX
    c = ws.cell(row=rr, column=4, value=f"=IF($C${s0 + 4}>0,C{rr}/$C${s0 + 4},0)")
    c.number_format = PCT
    c.border = B_BOX
rr += 1
ws.cell(row=rr, column=2, value="Total = investimento (aplicações)").font = F_BOLD
c = ws.cell(row=rr, column=3, value="=" + eng_ref("k_cxt"))
c.number_format = EUR
c.font = F_BOLD
rr += 1
ws.cell(row=rr, column=2, value="Exposição máxima do promotor (inclui adiantamento do subsídio)")
c = ws.cell(row=rr, column=3, value="=" + eng_ref("k_eqmax"))
c.number_format = EUR
DASH_CHART_ROW = rr + 3
ws.freeze_panes = "A5"

# ================================================================ RELATÓRIO
ws = ws_rep
ws["A1"] = '="RELATÓRIO — "&UPPER(NomeProjeto)&" — CENÁRIO "&UPPER(CenarioAtivo)'
ws["A1"].font = F_TITLE
ws["A2"] = "Espelho do motor de cálculo do cenário escolhido no Dashboard. Não editar."
ws["A2"].font = F_SUB
widths(ws, {"A": 44, "B": 10, "C": 15})
for col in PCOLS:
    ws.column_dimensions[col].width = 12.5
src_ws = ws_eng[0]
last_row = max(v for k, v in ENG.items() if not k.startswith("k_"))
for r_ in range(HDR_ROW, last_row + 1):
    for col in range(1, C0 + NP):
        s = src_ws.cell(row=r_, column=col)
        d = ws.cell(row=r_ - HDR_ROW + 4, column=col)
        if s.value is None:
            if s.fill and s.fill.fgColor and s.fill.fill_type:
                d.fill = copy(s.fill)
            continue
        if col <= 2 or r_ == HDR_ROW:
            d.value = s.value
        else:
            refs = ",".join(f"'{w.title}'!{L(col)}{r_}" for w in ws_eng)
            d.value = f"=CHOOSE(CenIdx,{refs})"
        d.font = copy(s.font)
        d.fill = copy(s.fill)
        d.border = copy(s.border)
        d.number_format = s.number_format
        d.alignment = copy(s.alignment)
REP_OFF = 4 - HDR_ROW
ws.freeze_panes = "D5"


def rep_row(key):
    return ENG[key] + REP_OFF


# gráficos (no Dashboard, com dados do Relatório)
ch = LineChart()
ch.title = "Receita, EBITDA e NOI (cenário ativo)"
ch.y_axis.title = "€"
ch.x_axis.title = "Período"
ch.height, ch.width = 8, 17
for k in ("rtot", "ebitda", "noi"):
    ref = Reference(ws_rep, min_col=C0 + 1, max_col=C0 + NP - 1, min_row=rep_row(k))
    ch.add_data(ref, from_rows=True, titles_from_data=False)
for s_, t in zip(ch.series, ("Receita", "EBITDA", "NOI")):
    from openpyxl.chart.series import SeriesLabel
    s_.tx = SeriesLabel(v=t)
ch.set_categories(Reference(ws_rep, min_col=C0 + 1, max_col=C0 + NP - 1, min_row=4))
for s_ in ch.series:
    s_.smooth = False
ws_dash.add_chart(ch, f"B{DASH_CHART_ROW}")

bc = BarChart()
bc.title = "Cash flow do promotor (anual e acumulado)"
bc.height, bc.width = 8, 17
bc.add_data(Reference(ws_rep, min_col=C0, max_col=C0 + NP - 1, min_row=rep_row("fcfs")), from_rows=True,
            titles_from_data=False)
bc.series[0].tx = SeriesLabel(v="Anual")
bc.set_categories(Reference(ws_rep, min_col=C0, max_col=C0 + NP - 1, min_row=4))
lc = LineChart()
lc.add_data(Reference(ws_rep, min_col=C0, max_col=C0 + NP - 1, min_row=rep_row("fcfsa")), from_rows=True,
            titles_from_data=False)
lc.series[0].tx = SeriesLabel(v="Acumulado")
lc.series[0].smooth = False
bc += lc
ws_dash.add_chart(bc, f"G{DASH_CHART_ROW}")

# ================================================================ CHECKS
ws = ws_chk
ws["A1"] = "VERIFICAÇÕES DO MODELO"
ws["A1"].font = F_TITLE
ws["A2"] = "Estado global"
ws["A2"].font = F_BOLD
widths(ws, {"A": 60, "B": 16, "C": 12, "D": 70})
header_row(ws, 4, ["Verificação", "Valor", "Estado", "O que fazer se falhar"])
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
]
for i, (lab, val, st, todo) in enumerate(checks):
    rr = 5 + i
    ws.cell(row=rr, column=1, value=lab)
    c = ws.cell(row=rr, column=2, value=val)
    c.number_format = '#,##0.00;(#,##0.00);"-"'
    if st is None:
        st = f'=IF(ABS(B{rr})<1,"OK","ERRO")'
    c = ws.cell(row=rr, column=3, value=st)
    c.alignment = Alignment(horizontal="center")
    c.font = F_BOLD
    ws.cell(row=rr, column=4, value=todo).font = F_SUB
CHK_END = 5 + len(checks) - 1
ws["B2"] = (f'=IF(COUNTIF(C5:C{CHK_END},"ERRO")>0,"ERRO",IF(COUNTIF(C5:C{CHK_END},"AVISO")>0,"OK c/ avisos","OK"))')
ws["B2"].font = F_BOLD
name("ChecksStatus", ws, "$B$2")
for rng_ in (f"C5:C{CHK_END}", "B2"):
    ws.conditional_formatting.add(rng_, CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_OK))
    ws.conditional_formatting.add(rng_, CellIsRule(operator="equal", formula=['"ERRO"'], fill=FILL_BAD))
    ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'ISNUMBER(SEARCH("AVISO",{rng_.split(":")[0]}))'],
                                                    fill=FILL_WARN))
ws.conditional_formatting.add("B2", CellIsRule(operator="equal", formula=['"OK c/ avisos"'], fill=FILL_WARN))
for d_ in ("H4",):
    ws_dash.conditional_formatting.add(d_, CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_OK))
    ws_dash.conditional_formatting.add(d_, CellIsRule(operator="equal", formula=['"ERRO"'], fill=FILL_BAD))
    ws_dash.conditional_formatting.add(d_, CellIsRule(operator="equal", formula=['"OK c/ avisos"'], fill=FILL_WARN))

# ================================================================ INÍCIO (instruções)
ws = ws0
ws.sheet_view.showGridLines = False
widths(ws, {"A": 2, "B": 26, "C": 100})
ws["B2"] = "MODELO DE AVALIAÇÃO DE INVESTIMENTOS"
ws["B2"].font = F_TITLE
ws["B3"] = "Rentabilidade e viabilidade de projetos de turismo e imobiliário — projeção anual até 20 anos"
ws["B3"].font = F_SUB
rows_txt = [
    ("COMO USAR", None),
    ("1. Pressupostos", "Preencha todas as células a azul (fundo amarelo). Comece pelas secções 1–3 e 8."),
    ("2. Capex", "Liste as rubricas de investimento, a classe (lista), o período e em quantos períodos é pago."),
    ("3. Pessoal / FSE", "Ajuste o quadro de pessoal e os FSE da unidade (valores do 1.º ano de operação)."),
    ("4. Dashboard", "Escolha o cenário (célula amarela) e leia o veredicto, os KPI e a comparação de cenários."),
    ("5. Checks", "Confirme que o estado é OK. Um ERRO significa que algum pressuposto é incoerente."),
    ("6. Relatório", "P&L, cash flow, dívida e saída do cenário ativo, ano a ano — pronto para imprimir."),
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
ws_dash.sheet_properties.tabColor = "1F3864"
ws_in.sheet_properties.tabColor = "FFC000"
ws_cx.sheet_properties.tabColor = "FFC000"
ws_p.sheet_properties.tabColor = "FFC000"
ws_f.sheet_properties.tabColor = "FFC000"
ws_rep.sheet_properties.tabColor = "70AD47"
ws_chk.sheet_properties.tabColor = "C00000"
for w in (ws_dash, ws_rep, ws_in, ws_chk) + tuple(ws_eng):
    w.page_setup.orientation = "landscape"
    w.page_setup.paperSize = w.PAPERSIZE_A4
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.active = 1
wb.save(OUT)
print("Gerado:", OUT)
