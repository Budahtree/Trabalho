"""Caso Villa Rosa (Peso da Régua) — pressupostos para o modelo de avaliação.

Fonte: VILLA_ROSA_INVESTOR_DECK_updated_AUG26.pdf (vendedor) + pressupostos prudentes da Potentia.
Cenários: coluna 1 = Prudente (base de decisão), 2 = Stress, 3 = Plano do vendedor.
"""

# ---- quadro de pessoal prudente (19,5 FTE) — (departamento, função, FTE, salário bruto/mês, bónus)
FTE = [
    ("Direção", "Diretor(a) geral", 1, 3000, 0.10),
    ("Front office", "Receção / guest relations", 3, 1050, 0.03),
    ("Housekeeping", "Empregado(a) de andares", 2, 950, 0.02),
    ("Cozinha", "Chef executivo", 1, 2500, 0.05),
    ("Cozinha", "Sous-chef", 1, 1600, 0.03),
    ("Cozinha", "Cozinheiro(a)", 2, 1150, 0.02),
    ("Cozinha", "Copa / ajudante", 1, 920, 0.02),
    ("Sala", "Maître / chefe de sala", 1, 1500, 0.03),
    ("Sala", "Empregado(a) de mesa", 3, 1000, 0.02),
    ("Sala", "Bar / escanção", 1, 1100, 0.02),
    ("Spa", "Terapeuta", 1.5, 1100, 0.03),
    ("Manutenção", "Manutenção / piscina", 1, 1100, 0.02),
    ("Comercial", "Vendas, eventos e reservas", 1, 1500, 0.05),
]
PAYROLL = sum(f * s * 14 * (1 + b) * 1.25 for _, _, f, s, b in FTE)

# ---- FSE prudentes — (rubrica, base anual a 60% de ocupação, % fixo)
FSE = [
    ("Energia, água e gás (piscina, spa, cozinha)", 70000, 0.60),
    ("Manutenção e conservação", 30000, 0.80),
    ("Seguros", 12000, 1.00),
    ("Sistemas, licenças e comunicações", 15000, 1.00),
    ("Marketing e comercial", 40000, 0.70),
    ("Contabilidade, jurídico e auditoria", 18000, 1.00),
    ("IMI, taxas e licenças", 20000, 1.00),
    ("Limpeza, lavandaria externa e outros", 25000, 0.60),
]
OCUP_REF = 0.60
FSE_FIX = sum(b * f for _, b, f in FSE)
FSE_VAR = sum(b * (1 - f) for _, b, f in FSE)

# ---- plano do vendedor (P&L "year 2 on cruise")
VEND_PESSOAL = 362048
VEND_OVERHEADS = 98834 + 20080          # services & marketing + taxes
VEND_OCUP = 0.62
AJ_PESSOAL_V = round(VEND_PESSOAL / PAYROLL - 1, 4)
AJ_FSE_V = round(VEND_OVERHEADS / (FSE_FIX + FSE_VAR * VEND_OCUP / OCUP_REF) - 1, 4)
EV_PRUDENTE = 4800000                    # €3,85M + €0,95M de dívida assumida
AJ_CAPEX_V = round(3850000 / EV_PRUDENTE - 1, 5)

CASE = {
    "subtitulo": "Caso: Villa Rosa (Peso da Régua). Cenários: Prudente (base de decisão) / Stress / Plano do vendedor.",
    "nomes_cenarios": ["Prudente", "Stress", "Plano do vendedor"],
    # identificação
    "NomeProjeto": "Villa Rosa", "Localizacao": "Peso da Régua (Douro)", "TipoAtivo": "Hotel / Turismo",
    "Analista": "Potentia Group", "AnoInicio": 2026, "AnoAbertura": 1, "AnoSaida": 10,
    "nota_AnoAbertura": "Aquisição em 2026 (período 0); primeiro ano completo de operação em 2027.",
    "nota_AnoSaida": "Horizonte de 10 anos com venda do ativo no fim.",
    # cenários: Prudente / Stress / Vendedor
    "S_Ocup1": (0.45, 0.35, 0.55),
    "nota_S_Ocup1": "Vendedor: média de 5,6 de 9 quartos ocupados (62%) no ano de cruzeiro.",
    "S_OcupG": (0.05, 0.05, 0.07),
    "S_OcupMax": (0.60, 0.50, 0.62),
    "S_AjADR": (0.0, -0.08, 0.085),
    "nota_S_AjADR": "Vendedor: ADR médio ponderado ≈ €266 (€534k / 2.010 noites).",
    "S_AjCapex": (0.0, 0.0, AJ_CAPEX_V),
    "nota_S_AjCapex": "Prudente/Stress: €3,85M + €0,95M de dívida assumida (EV €4,8M). Vendedor: €3,85M já com a dívida.",
    "S_AjCustos": (0.0, 0.0, 0.0),
    "S_AjPessoal": (0.0, -0.10, AJ_PESSOAL_V),
    "nota_S_AjPessoal": "Vendedor: €362k (18,6% da receita). Stress: equipa mais curta (−10%).",
    "S_AjFSE": (0.0, 0.0, AJ_FSE_V),
    "nota_S_AjFSE": "Vendedor: só €119k (marketing + taxas); sem energia, manutenção, seguros nem sistemas.",
    "S_Ramp1": (0.70, 0.60, 0.85),
    "nota_S_Ramp1": "Ativo sem histórico de operação: 1.º ano é de lançamento.",
    "S_Ramp2": (0.90, 0.80, 1.00),
    "S_AjPreco": (0.0, 0.0, 0.0),
    "S_ExitYield": (0.085, 0.095, 0.075),
    "nota_S_ExitYield": "Ativo pequeno, operacional e com restauração: prémio de risco face a hotéis maiores.",
    # operação
    "Unidades": 9, "Dias": 365,
    "ADR_tab": [235, 245, 250, 255, 0], "ADR_g": 0.025,
    "nota_ADR_tab": "Vendedor: €250–€350 consoante o mês. Prudente: €235 no arranque, a subir com a reputação.",
    "FB_quarto": 0, "FB_g": 0.0, "Outros_pct": 0.0,
    "nota_FB_quarto": "A restauração (hóspedes + externos) está na secção 3B.",
    "outras_receitas": [
        ("Restaurante (hóspedes + externos)", 475200, 252000, 1236000, 0.32,
         "Prudente: 30 refeições/dia × 330 dias × €48. Vendedor: 22.500 refeições × €55."),
        ("Bar", 40000, 25000, 0, 0.25, "Vendedor: €106k no gráfico, ausente da demonstração de resultados."),
        ("Spa & wellness", 72600, 40000, 160800, 0.15, "Prudente: 2 tratamentos/dia × 330 × €110."),
        ("Eventos privados", 40000, 20000, 0, 0.35, ""),
        ("Venda de vinho", 40000, 25000, 0, 0.15, "Custo do stock entra via amortização (4 anos)."),
        ("Experiências", 30000, 15000, 17000, 0.70, "Receita líquida de parceiros (cruzeiros, quintas)."),
    ],
    "Rec_g": 0.025,
    "C_Amen": 18, "nota_C_Amen": "Inclui custo do pequeno-almoço.", "C_Lav": 9, "InflDrivers": 0.025,
    "OTA_peso": 0.40, "OTA_com": 0.18, "TaxasPag": 0.015,
    "CustoOutros_pct": 0.0,
    "InflSal": 0.025, "InflFSE": 0.02, "nota_InflFSE": "", "OcupRef": OCUP_REF,
    "MgmtFee": 0.0, "nota_MgmtFee": "Gestão pela Potentia (custo via shared services).", "IncFee": 0.0,
    "SS_g": 0.03, "FFE_res": 0.03, "FFE_inicio": 2, "CapexManut": 0,
    "Vendas_tab": [0, 0, 0, 0, 0], "Vendas_reduz": "Não", "PrecoUnid": 0, "CustoVenda_pct": 0.0,
    "CustoContab_unid": 0, "nota_CustoVenda_pct": "", "nota_CustoContab_unid": "Não aplicável.",
    # fiscal
    "IRC": 0.205, "nota_IRC": "IRC 19% (2026) + derrama municipal até 1,5%.",
    "VidaConstr": 30, "VidaFFE": 8, "VidaOutros": 4, "TaxaMaisValia": 0.205,
    # financiamento
    "Divida": 950000,
    "nota_Divida": "Empréstimo existente. Termos por confirmar: taxa, prazo, garantias, cláusula de mudança de controlo.",
    "DividaPer": 0, "Juro": 0.05, "Selo": 0.04, "Prazo": 12, "Carencia": 1, "TipoAmort": "Francês",
    "ComAbertura": 0.0,
    "Subsidio": 0, "nota_Subsidio": "Não aplicável.", "SubsidioPer": 1,
    "InvNome": "—", "InvMontante": 0, "nota_InvMontante": "Não aplicável.", "InvPer": 0, "InvTaxa": 0.0,
    "InvRecompra": 5,
    "CustosSaida_pct": 0.02, "TaxaDescProj": 0.09, "CustoEquity": 0.12,
    "Crit_TIRproj": 0.09, "Crit_TIRprom": 0.12, "Crit_DSCR": 1.30, "Crit_Payback": 10, "Crit_MOIC": 1.5,
    # capex: EV prudente €4,8M + custos da Potentia
    "capex_rows": [
        ("Aquisição — terreno (alocação estimada)", "Terreno / aquisição", 900000, "Sim", 0, 1,
         "Alocação do preço a validar com avaliação independente."),
        ("Aquisição — edifício e obras de recuperação", "Construção", 3300000, "Sim", 0, 1, ""),
        ("Aquisição — mobiliário, cozinha, spa e equipamento", "FF&E", 300000, "Sim", 0, 1,
         "Vendedor: «pacote turnkey» > €200k."),
        ("Aquisição — stock de vinho (12 barricas, 4.800 l)", "Gastos amortizáveis", 300000, "Sim", 0, 1,
         "Valor do vendedor. Valor realizável estimado ≈ €180k."),
        ("Custos de transação (IMT eventual, selo, due diligence)", "Gastos amortizáveis", 200000, "Não", 0, 1,
         "≈4% do EV. IMT depende da estrutura (art. 2.º CIMT)."),
        ("Pré-abertura e lançamento", "Gastos amortizáveis", 150000, "Não", 0, 1,
         "Recrutamento, formação, marketing, stocks iniciais."),
        ("Fundo de maneio", "Fundo de maneio", 100000, "Não", 0, 1, ""),
        ("Contingência (correções, licenciamento)", "Construção", 100000, "Não", 0, 1, ""),
    ],
    "fte": FTE,
    "port": [("Villa Rosa (este projeto)", 9), ("Cocoon", 48), ("Alisios", 117), ("Évora", 27), ("VDT", 43),
             (None, None)],
    "fse": FSE,
    # estrutura
    "PromotorNome": "Potentia Group", "UsarSPV": "Não", "SPVNome": "—", "PartPromSPV": 1.0, "PartSPVProj": 1.0,
    "PartInvProj": 0.0, "CapSocial": 50000, "BancoNome": "Empréstimo existente (banco a identificar)",
    "DividaVia": "Sociedade do projeto", "SubNome": "—",
    "aloc_def": {1: (950000, 0, 0)},
    "sens_steps": {
        "adr": [-0.10, -0.05, 0.0, 0.05, 0.10],
        "occ": [-0.10, -0.05, 0.0, 0.05, 0.10],
        "cpx": [-0.30, -0.20, -0.10, 0.0, 0.10],
        "ey": [0.075, 0.08, 0.085, 0.09, 0.095],
    },
}
