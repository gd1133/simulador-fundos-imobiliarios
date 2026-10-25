# -*- coding: utf-8 -*-
"""Gera o simulador de fundos imobiliários (exercício de Excel)."""

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = r"C:\Users\Oem\Desktop\n8nmoto\simulador-fiis\simulador-fiis.xlsx"

TIPOS = [
    "Papel (recebíveis)",
    "Logística",
    "Shoppings",
    "Lajes corporativas",
    "Híbrido",
    "Fundo de fundos",
]

# Exemplo didático. Cada perfil soma 1. Não é recomendação de investimento.
PERFIS = {
    "Conservador": [0.45, 0.20, 0.10, 0.10, 0.10, 0.05],
    "Moderado": [0.25, 0.20, 0.15, 0.15, 0.15, 0.10],
    "Arrojado": [0.10, 0.20, 0.20, 0.20, 0.20, 0.10],
    "Foco em renda": [0.40, 0.20, 0.10, 0.10, 0.15, 0.05],
}

NAVY = "0B3A4A"
TEAL = "0E6B62"
TEAL_DARK = "0A4F49"
GOLD = "C8922A"
INPUT = "FFF3C4"
CALC = "E6E6E6"
WHITE = "FFFFFF"
INK = "1C2833"
MUTED = "5D6D7E"
CARD = "F4F7F7"
LINE = "D5D8DC"
OK_GREEN = "C6EFCE"
RESULT = "E8F6F3"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
thick_bottom = Border(bottom=Side(style="medium", color=TEAL))

font_base = Font(name="Calibri", size=12, color=INK)
font_label = Font(name="Calibri", size=12, color=INK)
font_small = Font(name="Calibri", size=10, color=MUTED)
font_white_bold = Font(name="Calibri", size=12, bold=True, color=WHITE)
font_section = Font(name="Calibri", size=13, bold=True, color=WHITE)
font_title = Font(name="Calibri", size=26, bold=True, color=WHITE)
font_sub = Font(name="Calibri", size=12, color="E8F6F3")
font_input = Font(name="Calibri", size=13, bold=True, color=INK)
font_calc = Font(name="Calibri", size=13, bold=True, color="2C3E50")
font_result = Font(name="Calibri", size=16, bold=True, color=TEAL_DARK)
font_header = Font(name="Calibri", size=11, bold=True, color=WHITE)

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_input = PatternFill("solid", fgColor=INPUT)
fill_calc = PatternFill("solid", fgColor=CALC)
fill_card = PatternFill("solid", fgColor=CARD)
fill_result = PatternFill("solid", fgColor=RESULT)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_gold = PatternFill("solid", fgColor="FBF6EA")
fill_header = PatternFill("solid", fgColor=TEAL_DARK)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_mid = Alignment(horizontal="left", vertical="center", wrap_text=True)
right_mid = Alignment(horizontal="right", vertical="center")

BRL = 'R$ #,##0.00'
PCT = '0.00%'
ANOS = '0" anos"'


def paint(cell, value=None, font=None, fill=None, align=None, fmt=None, border=None):
    if value is not None:
        cell.value = value
    cell.font = font or font_base
    if fill is not None:
        cell.fill = fill
    cell.alignment = align or left_mid
    if fmt:
        cell.number_format = fmt
    if border is not None:
        cell.border = border


def section(ws, cell_range, title):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    paint(cell, title, font_section, fill_teal, Alignment(horizontal="left", vertical="center", indent=1))
    # fill the rest of the merge so the bar is solid
    start, end = cell_range.split(":")
    # handled by merge display; still fill each cell
    from openpyxl.utils import range_boundaries
    min_c, min_r, max_c, max_r = range_boundaries(cell_range)
    for row in ws.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=max_c):
        for c in row:
            c.fill = fill_teal
            c.font = font_section
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def main():
    for nome, pesos in PERFIS.items():
        if abs(sum(pesos) - 1) > 1e-9:
            raise SystemExit(f"Perfil {nome} não soma 100% ({sum(pesos)})")
        if len(pesos) != len(TIPOS):
            raise SystemExit(f"Perfil {nome} não tem 6 tipos")

    wb = Workbook()
    sim = wb.active
    sim.title = "Simulador"
    apoio = wb.create_sheet("Apoio")

    montar_apoio(apoio)
    montar_simulador(sim)
    nomear(wb)
    validar_perfil(sim)
    grafico(sim)

    sim.sheet_view.showGridLines = False
    sim.sheet_view.zoomScale = 110
    sim.page_setup.orientation = "landscape"
    sim.page_setup.fitToWidth = 1
    sim.page_setup.fitToHeight = 1
    sim.page_setup.paperSize = sim.PAPERSIZE_A4
    sim.sheet_properties.pageSetUpPr.fitToPage = True
    sim.print_area = "A1:L34"
    sim.page_setup.horizontalCentered = True
    sim.sheet_properties.tabColor = TEAL
    apoio.sheet_properties.tabColor = "C8922A"
    apoio.sheet_view.showGridLines = False
    apoio.sheet_view.zoomScale = 120
    apoio.page_setup.orientation = "landscape"
    apoio.page_setup.fitToWidth = 1
    apoio.page_setup.fitToHeight = 1
    apoio.sheet_properties.pageSetUpPr.fitToPage = True
    apoio.print_area = "A1:J32"

    wb.properties.title = "Simulador de Fundos Imobiliários"
    wb.properties.subject = "Exercício: aporte, prazo, taxa, patrimônio, dividendos e perfil"
    wb.properties.creator = "Simulador FIIs"

    wb.save(OUT)
    print(OUT)


def montar_apoio(ws):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 36
    ws.merge_cells("A1:D1")
    paint(ws["A1"], "TABELA DE APOIO", font_title, fill_navy, Alignment(horizontal="left", vertical="center", indent=1))
    ws["A1"].font = Font(name="Calibri", size=20, bold=True, color=WHITE)
    for col in range(1, 5):
        ws.cell(1, col).fill = fill_navy
    ws.merge_cells("A2:D2")
    paint(
        ws["A2"],
        "A chave junta o perfil e o tipo de fundo. O PROCV da aba Simulador procura essa chave e devolve o percentual.",
        font_sub,
        fill_navy,
        left_mid,
    )
    ws["A2"].font = Font(name="Calibri", size=11, color="E8F6F3")
    for col in range(1, 5):
        ws.cell(2, col).fill = fill_navy

    headers = ["Chave", "Perfil", "Tipo de fundo", "Percentual"]
    for col, texto in enumerate(headers, 1):
        paint(ws.cell(3, col), texto, font_header, fill_header, center, border=thin)
    ws.row_dimensions[3].height = 22
    ws.auto_filter.ref = "A3:D27"

    row = 4
    for perfil, pesos in PERFIS.items():
        for tipo, peso in zip(TIPOS, pesos):
            ws.row_dimensions[row].height = 20
            # Chave composta: perfil|tipo. Cinza porque é fórmula.
            paint(ws.cell(row, 1), f'=B{row}&"|"&C{row}', font_calc, fill_calc, left_mid, border=thin)
            paint(ws.cell(row, 2), perfil, font_base, fill_white, left_mid, border=thin)
            paint(ws.cell(row, 3), tipo, font_base, fill_white, left_mid, border=thin)
            paint(ws.cell(row, 4), peso, font_input, fill_input, center, PCT, thin)
            row += 1

    # Lista usada na validação de dados
    paint(ws["F3"], "Perfis da lista", font_header, fill_header, center, border=thin)
    for i, perfil in enumerate(PERFIS):
        paint(ws.cell(4 + i, 6), perfil, font_base, fill_white, left_mid, border=thin)
    paint(ws["G3"], "Soma do perfil", font_header, fill_header, center, border=thin)
    for i, perfil in enumerate(PERFIS):
        r = 4 + i
        paint(
            ws.cell(r, 7),
            f'=SUMIF($B$4:$B$27,F{r},$D$4:$D$27)',
            font_calc,
            fill_calc,
            center,
            PCT,
            thin,
        )
        ws.row_dimensions[r].height = 20

    paint(ws["I3"], "Sugestão sobre o salário", font_header, fill_header, center, border=thin)
    paint(ws["I4"], 0.30, font_input, fill_input, center, PCT, thin)
    ws["I4"].comment = Comment(
        "30% é a regra didática do exercício: sugestão de aporte = salário × este percentual. "
        "Mude só se quiser testar outra regra. Não é conselho financeiro.",
        "Simulador",
        width=280,
        height=70,
    )

    ws.merge_cells("A30:G33")
    paint(
        ws["A30"],
        "De onde vêm os percentuais: são um exemplo criado para este exercício, para a divisão "
        "mudar quando o perfil muda. Conservador concentra em papel (recebíveis). Arrojado abre "
        "espaço para shoppings, lajes e híbrido. Foco em renda fica perto do conservador, com um "
        "pouco mais de híbrido. Cada linha de perfil soma 100%. Isso não é recomendação de investimento.\n\n"
        "Por que a chave existe: o mesmo tipo (Logística, por exemplo) aparece em todos os perfis. "
        "Um PROCV só pelo tipo devolveria sempre a primeira logística da tabela. A chave "
        "Moderado|Logística aponta para uma linha só.",
        font_small,
        fill_gold,
        Alignment(horizontal="left", vertical="top", wrap_text=True),
    )
    for col in range(1, 8):
        ws.cell(30, col).fill = fill_gold
        ws.cell(30, col).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[30].height = 30
    ws.row_dimensions[31].height = 30
    ws.row_dimensions[32].height = 30
    ws.row_dimensions[33].height = 30

    ws.conditional_formatting.add(
        "G4:G7",
        FormulaRule(formula=["ABS(G4-1)<0.000001"], fill=PatternFill("solid", fgColor=OK_GREEN)),
    )

    widths = {"A": 34, "B": 20, "C": 26, "D": 14, "E": 3, "F": 20, "G": 16, "H": 3, "I": 28}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[3].height = 22
    ws.freeze_panes = "A4"
    ws.oddHeader.left.text = "Apoio — não é a tela do usuário"


def montar_simulador(ws):
    widths = {
        "A": 48, "B": 20, "C": 22, "D": 42, "E": 18, "F": 22,
        "G": 3, "H": 16, "I": 16, "J": 16, "K": 16, "L": 18,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    ws.merge_cells("A3:L3")
    for col in range(1, 13):
        ws.cell(1, col).fill = fill_navy
        ws.cell(2, col).fill = fill_navy
        ws.cell(3, col).fill = fill_teal
    paint(
        ws["A1"],
        "SIMULADOR DE FUNDOS IMOBILIÁRIOS",
        font_title,
        fill_navy,
        Alignment(horizontal="left", vertical="center", indent=1),
    )
    paint(
        ws["A2"],
        "Quanto aportar, por quanto tempo, e o que isso vira em patrimônio e dividendos.",
        font_sub,
        fill_navy,
        Alignment(horizontal="left", vertical="center", indent=1),
    )
    paint(
        ws["A3"],
        "Amarelo: você edita.    Cinza: a planilha calcula.    Exemplo didático — não é recomendação de investimento.",
        Font(name="Calibri", size=11, color=WHITE),
        fill_teal,
        Alignment(horizontal="left", vertical="center", indent=1),
    )

    ws.row_dimensions[4].height = 10

    section(ws, "A5:B5", "1. O QUE VOCÊ DEFINE")
    section(ws, "D5:F5", "2. AS CINCO RESPOSTAS")
    section(ws, "H5:L5", "COMO A CONTA FUNCIONA")
    for r in (5,):
        ws.row_dimensions[r].height = 24

    entradas = [
        (6, "Salário mensal (exemplo)", 6000, BRL, True, "salario"),
        (7, "Sugestão de aporte (30% do salário)", "=salario*percentual_sugestao", BRL, False, None),
        (8, "Quanto investir por mês", 1500, BRL, True, None),
        (9, "Por quantos anos", 10, "0", True, None),
        (10, "Rendimento mensal da carteira", 0.008, PCT, True, None),
        (11, "Perfil do investidor", "Moderado", None, True, None),
    ]
    for row, rotulo, valor, fmt, editavel, _nome in entradas:
        ws.row_dimensions[row].height = 24
        paint(ws.cell(row, 1), rotulo, font_label, fill_card, left_mid, border=thin)
        fill = fill_input if editavel else fill_calc
        font = font_input if editavel else font_calc
        paint(ws.cell(row, 2), valor, font, fill, center, fmt, thin)

    ws["B10"].comment = Comment(
        "Taxa mensal, não anual. 0,80% ao mês é só um número redondo para o exercício. "
        "Na aba Apoio este mesmo conceito se chama taxa_mensal.",
        "Simulador",
        width=260,
        height=70,
    )
    ws.merge_cells("A12:B12")
    paint(
        ws["A12"],
        "Edite só o amarelo. O cinza se recalcula.",
        font_small,
        fill_card,
        Alignment(horizontal="left", vertical="center", indent=1),
        border=thin,
    )
    ws["B12"].fill = fill_card
    ws["B12"].border = thin

    ws["B11"].comment = Comment(
        "Abra a lista e troque o perfil. A tabela de baixo e o gráfico mudam juntos. "
        "A lista vem do intervalo nomeado lista_perfis, na aba Apoio.",
        "Simulador",
        width=260,
        height=70,
    )

    respostas = [
        (6, "1. Quanto investir por mês", "=aporte", BRL, False),
        (7, "2. Por quantos anos", "=prazo_anos", ANOS, False),
        (8, "3. Taxa de rendimento mensal", "=taxa_mensal", PCT, False),
        (9, "4. Patrimônio acumulado", "=FV(taxa_mensal,prazo_anos*12,-aporte)", BRL, True),
        (10, "5. Dividendos por mês", "=patrimonio*taxa_mensal", BRL, True),
        (11, "Total aportado no prazo", "=aporte*prazo_anos*12", BRL, False),
        (12, "Ganho acima do total aportado", "=patrimonio-F11", BRL, False),
    ]
    ws.row_dimensions[12].height = 24
    for row, rotulo, formula, fmt, destaque in respostas:
        ws.row_dimensions[row].height = 24
        paint(ws.cell(row, 4), rotulo, font_label, fill_card, left_mid, border=thin)
        paint(ws.cell(row, 5), None, font_calc, fill_card, center, border=thin)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
        fonte = font_result if destaque else font_calc
        fundo = fill_result if destaque else fill_calc
        paint(ws.cell(row, 6), formula, fonte, fundo, center, fmt, thin)

    ws["F9"].comment = Comment(
        "VF (FV em inglês) estima o patrimônio no fim do prazo.\n"
        "taxa = taxa_mensal (juros de cada mês).\n"
        "nper = prazo_anos × 12 (quantidade de aportes).\n"
        "pgto = -aporte. O sinal negativo é o dinheiro que sai do bolso.\n"
        "vp fica de fora: a conta começa do zero.\n"
        "tipo fica de fora: 0, aporte no fim de cada mês.\n"
        "Com pgto negativo, o VF volta positivo.",
        "Simulador",
        width=280,
        height=140,
    )
    ws["F10"].comment = Comment(
        "Dividendo do mês = patrimônio acumulado × taxa mensal. "
        "É o rendimento de um mês sobre o saldo do fim do prazo, sem novo aporte.",
        "Simulador",
        width=260,
        height=70,
    )

    ws.merge_cells("H6:L12")
    explicacao = (
        "Patrimônio (resposta 4) usa a função VF.\n"
        "Os três argumentos usados aqui:\n"
        "• taxa — rendimento de um mês (taxa_mensal);\n"
        "• nper — meses, ou seja, anos × 12;\n"
        "• pgto — aporte com sinal negativo.\n\n"
        "Dividendo (resposta 5) = patrimônio × taxa mensal.\n\n"
        "A divisão por fundo usa PROCV com chave composta: "
        "perfil & \"|\" & tipo. Exemplo: Moderado|Logística. "
        "A aba Apoio tem essa chave na 1ª coluna e o percentual na 4ª. "
        "O 4º argumento é FALSO: a busca é exata."
    )
    paint(
        ws["H6"],
        explicacao,
        Font(name="Calibri", size=11, color=INK),
        fill_gold,
        Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1),
    )
    for row in range(6, 13):
        for col in range(8, 13):
            ws.cell(row, col).fill = fill_gold
            ws.cell(row, col).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    ws.row_dimensions[13].height = 10

    section(ws, "A14:F14", "3. CENÁRIOS — O MESMO APORTE EM OUTROS PRAZOS")
    ws.row_dimensions[14].height = 24

    paint(ws["A15"], "Prazo", font_header, fill_header, left_mid, border=thin)
    for col, anos in enumerate([2, 5, 10, 20, 30], 2):
        paint(ws.cell(15, col), anos, font_header, fill_header, center, ANOS, thin)
    ws.row_dimensions[15].height = 22

    linhas_cenario = [
        (16, "Patrimônio com o aporte informado", "=FV(taxa_mensal,{col}15*12,-aporte)"),
        (17, "Dividendos por mês nesse patrimônio", "={col}16*taxa_mensal"),
        (18, "Patrimônio se aportar 30% do salário", "=FV(taxa_mensal,{col}15*12,-sugestao_aporte)"),
        (19, "Dividendos por mês nessa 2ª simulação", "={col}18*taxa_mensal"),
    ]
    for row, rotulo, modelo in linhas_cenario:
        ws.row_dimensions[row].height = 24
        paint(ws.cell(row, 1), rotulo, font_label, fill_card, left_mid, border=thin)
        for col in range(2, 7):
            letra = get_column_letter(col)
            formula = modelo.format(col=letra)
            paint(ws.cell(row, col), formula, font_calc, fill_calc, center, BRL, thin)

    ws["A16"].comment = Comment(
        "A fórmula é a mesma do patrimônio principal. "
        "taxa_mensal e aporte são intervalos nomeados: não andam quando a fórmula é arrastada para o lado. "
        "Só o prazo (a célula de cima) muda de coluna.",
        "Simulador",
        width=280,
        height=90,
    )

    ws.row_dimensions[20].height = 10
    section(ws, "A21:D21", "4. DIVISÃO DO APORTE PELO PERFIL")
    ws.row_dimensions[21].height = 24

    for col, texto in enumerate(["Tipo de fundo", "Percentual", "Valor do aporte", "Chave perfil|tipo"], 1):
        paint(ws.cell(22, col), texto, font_header, fill_header, center, border=thin)
    ws.row_dimensions[22].height = 22

    for i, tipo in enumerate(TIPOS):
        row = 23 + i
        ws.row_dimensions[row].height = 22
        paint(ws.cell(row, 1), tipo, font_label, fill_card, left_mid, border=thin)
        paint(
            ws.cell(row, 4),
            f'=perfil&"|"&A{row}',
            Font(name="Calibri", size=10, color="2C3E50"),
            fill_calc,
            center,
            border=thin,
        )
        paint(
            ws.cell(row, 2),
            f'=IFERROR(VLOOKUP(D{row},tabela_alocacao,4,FALSE),0)',
            font_calc,
            fill_calc,
            center,
            PCT,
            thin,
        )
        paint(
            ws.cell(row, 3),
            f"=B{row}*aporte",
            font_calc,
            fill_calc,
            center,
            BRL,
            thin,
        )

    ws["B23"].comment = Comment(
        "PROCV procura a chave da coluna D dentro de tabela_alocacao.\n"
        "Devolve a 4ª coluna (o percentual).\n"
        "FALSO = correspondência exata. Sem isso, o Excel aceitaria um valor parecido.",
        "Simulador",
        width=280,
        height=90,
    )

    ws.row_dimensions[29].height = 24
    paint(ws["A29"], "Soma do perfil (tem de ser 100%)", font_label, fill_card, left_mid, border=thin)
    paint(ws["B29"], "=SUM(B23:B28)", font_result, fill_calc, center, PCT, thin)
    paint(ws["C29"], "=SUM(C23:C28)", font_calc, fill_calc, center, BRL, thin)
    paint(
        ws["D29"],
        '=IF(ABS(B29-1)<0.0000001,"OK — soma 100%","Revise: não soma 100%")',
        font_calc,
        fill_calc,
        center,
        border=thin,
    )
    ws.conditional_formatting.add(
        "B29",
        FormulaRule(formula=["ABS(B29-1)<0.000001"], fill=PatternFill("solid", fgColor=OK_GREEN)),
    )
    ws.conditional_formatting.add(
        "D29",
        FormulaRule(formula=['B29=1'], fill=PatternFill("solid", fgColor=OK_GREEN)),
    )

    ws.merge_cells("A31:D33")
    paint(
        ws["A31"],
        "Segunda simulação: as linhas 18 e 19 dos cenários trocam o aporte informado pela sugestão "
        "de 30% do salário (sugestao_aporte). O prazo e a taxa continuam os mesmos. "
        "Os percentuais de cada perfil estão na aba Apoio e são um exemplo do exercício, não uma carteira recomendada. "
        "Salário, aporte e taxa desta tela também são números de exemplo.",
        font_small,
        fill_card,
        Alignment(horizontal="left", vertical="top", wrap_text=True),
    )
    for row in range(31, 34):
        ws.row_dimensions[row].height = 18
        for col in range(1, 5):
            ws.cell(row, col).fill = fill_card

    ws.row_dimensions[34].height = 8
    # respiro visual abaixo do banner já foi; congela só o título
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = None
    ws.sheet_view.view = "normal"
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.4
    ws.oddFooter.left.text = "Simulador de fundos imobiliários  |  exercício didático"
    ws.oddFooter.right.text = "Página &P"


def nomear(wb):
    nomes = {
        "salario": "Simulador!$B$6",
        "sugestao_aporte": "Simulador!$B$7",
        "aporte": "Simulador!$B$8",
        "prazo_anos": "Simulador!$B$9",
        "taxa_mensal": "Simulador!$B$10",
        "rendimento_carteira": "Simulador!$B$10",
        "perfil": "Simulador!$B$11",
        "patrimonio": "Simulador!$F$9",
        "dividendos_mensais": "Simulador!$F$10",
        "percentual_sugestao": "Apoio!$I$4",
        "lista_perfis": "Apoio!$F$4:$F$7",
        "tabela_alocacao": "Apoio!$A$4:$D$27",
    }
    for nome, ref in nomes.items():
        wb.defined_names.add(DefinedName(name=nome, attr_text=ref))


def validar_perfil(ws):
    dv = DataValidation(
        type="list",
        formula1="=lista_perfis",
        allow_blank=False,
        showDropDown=False,  # False mostra a seta no Excel; True esconde
        showErrorMessage=True,
        showInputMessage=True,
        errorTitle="Perfil",
        error="Escolha um perfil da lista.",
        promptTitle="Perfil do investidor",
        prompt="Conservador, Moderado, Arrojado ou Foco em renda.",
    )
    dv.add(ws["B11"])
    ws.add_data_validation(dv)


def grafico(ws):
    pizza = PieChart()
    pizza.title = "Divisão do aporte"
    rotulos = Reference(ws, min_col=1, min_row=23, max_row=28)
    dados = Reference(ws, min_col=3, min_row=22, max_row=28)
    pizza.add_data(dados, titles_from_data=True)
    pizza.set_categories(rotulos)
    pizza.dataLabels = DataLabelList()
    pizza.dataLabels.showPercent = True
    pizza.dataLabels.showVal = False
    pizza.dataLabels.showCatName = False
    pizza.dataLabels.showSerName = False
    pizza.width = 15
    pizza.height = 8
    if pizza.legend is not None:
        pizza.legend.position = "b"
    ws.add_chart(pizza, "H14")


if __name__ == "__main__":
    main()
