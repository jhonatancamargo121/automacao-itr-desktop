# -*- coding: utf-8 -*-
"""Gera a planilha-modelo de clientes/imoveis rurais para automacao do ITR."""
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

AZUL = "0066CC"
AZUL_CLARO = "E6F0FA"
CINZA = "F5F5F5"
BORDA = "D0D0D0"
PRETO = "000000"

F_INPUT = Font(color=AZUL)                 # valores digitados
F_CALC = Font(color=PRETO)                 # formulas de calculo
F_CROSS = Font(color="008000")             # referencias a outras abas
THIN = Side(style="thin", color=BORDA)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MUNICIPIOS = [
    "Campinaçu", "Santa Rosa de Goiás", "Crixás", "Itapaci", "Jaraguá",
    "São Francisco de Goiás", "Inhumas", "Trombas", "Pirenópolis",
    "Vila Propício", "Trindade", "Petrolina de Goiás",
]
CATEGORIAS = [
    "Lavoura aptidão boa", "Lavoura aptidão regular", "Lavoura aptidão restrita",
    "Pastagem plantada", "Silvicultura ou pastagem natural", "Preservação/fauna e flora",
]
VTN_PETROLINA = {
    "Lavoura aptidão boa": 34450.00,
    "Lavoura aptidão regular": 31500.00,
    "Lavoura aptidão restrita": 25320.00,
    "Pastagem plantada": 29490.00,
    "Silvicultura ou pastagem natural": 21220.00,
    "Preservação/fauna e flora": 18530.00,
}

wb = Workbook()

# ============================== CAPA ==============================
ws = wb.active
ws.title = "Capa"
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [3, 30, 22, 22, 22, 22, 22, 3]):
    ws.column_dimensions[col].width = w
ws.row_dimensions[2].height = 38
ws.merge_cells("B2:G2")
ws["B2"] = "Automação ITR — Clientes e Cálculo de VTN 2026"
ws["B2"].font = Font(size=18, bold=True)
ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells("B4:G4")
ws["B4"] = "Planilha de entrada do robô de preenchimento do programa ITR (Receita Federal)"
ws["B4"].font = Font(size=12, color="808080")
ws["B4"].alignment = Alignment(horizontal="center")

metricas = [
    ("Municípios atendidos", 12),
    ("Categorias de VTN", 6),
    ("Ano de referência", 2026),
    ("Município com laudo carregado", "Petrolina de Goiás"),
]
ws["B6"] = "Resumo"
ws["B6"].font = Font(size=13, bold=True)
for i, (k, v) in enumerate(metricas):
    r = 7 + i
    ws.cell(row=r, column=2, value=k).border = BORDER
    c = ws.cell(row=r, column=3, value=v)
    c.border = BORDER
    c.font = F_INPUT
    ws.cell(row=r, column=2).fill = PatternFill("solid", start_color=CINZA)

ws["B13"] = "Índice das abas"
ws["B13"].font = Font(size=13, bold=True)
indice = [
    ("Clientes", "Cadastro dos imóveis/clientes — o robô lê esta aba linha por linha"),
    ("VTN 2026", "Tabela de Valor da Terra Nua por município e categoria (R$/ha)"),
    ("Fontes", "Origem dos dados de VTN"),
]
for i, (nome, desc) in enumerate(indice):
    r = 14 + i
    ws.cell(row=r, column=2, value=nome).font = Font(bold=True)
    ws.cell(row=r, column=2).border = BORDER
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    ws.cell(row=r, column=3, value=desc).border = BORDER

ws["B19"] = "Como usar"
ws["B19"].font = Font(size=13, bold=True)
notas = [
    "1. Preencha uma linha por imóvel na aba Clientes (colunas azuis = digitação).",
    "2. Complete os valores de VTN dos demais municípios na aba VTN 2026 conforme os laudos.",
    "3. Deixe a coluna 'status' como PENDENTE nas linhas que o robô deve processar.",
    "4. Abra o programa ITR na tela inicial, execute o robô (executar.bat) e acompanhe.",
    "5. O robô NUNCA transmite sozinho: ele pausa e pede sua confirmação antes de enviar.",
]
for i, n in enumerate(notas):
    ws.cell(row=20 + i, column=2, value=n).font = Font(size=10, color="808080")

# ============================== VTN 2026 ==============================
wv = wb.create_sheet("VTN 2026")
wv.sheet_view.showGridLines = False
for col, w in zip("ABCDEF", [3, 26, 32, 16, 44, 40]):
    wv.column_dimensions[col].width = w
wv.row_dimensions[2].height = 30
wv.merge_cells("B2:E2")
wv["B2"] = "Valor da Terra Nua (VTN) por hectare — 2026"
wv["B2"].font = Font(size=14, bold=True)
wv.row_dimensions[4].height = 22
headers = ["Município", "Categoria", "Valor (R$/ha)", "Fonte", "Chave (automático)"]
for j, h in enumerate(headers):
    c = wv.cell(row=4, column=2 + j, value=h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color="333333")
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")

r = 5
for mun in MUNICIPIOS:
    for cat in CATEGORIAS:
        wv.cell(row=r, column=2, value=mun).border = BORDER
        wv.cell(row=r, column=3, value=cat).border = BORDER
        cval = wv.cell(row=r, column=4)
        cval.border = BORDER
        cval.number_format = '"R$" #,##0.00'
        cval.font = F_INPUT
        cfonte = wv.cell(row=r, column=5)
        cfonte.border = BORDER
        cfonte.font = F_INPUT
        if mun == "Petrolina de Goiás":
            cval.value = VTN_PETROLINA[cat]
            cfonte.value = "Laudo Técnico VTN Petrolina de Goiás 2026 (IN RFB 1877/2019)"
        cchave = wv.cell(row=r, column=6, value=f'=B{r}&"|"&C{r}')
        cchave.font = F_CALC
        r += 1
ULTIMA_VTN = r - 1  # linha 76

# ============================== CLIENTES ==============================
wc = wb.create_sheet("Clientes")
wc.sheet_view.showGridLines = False

colunas = [
    # (titulo, largura, tipo)  tipo: input/calc/cross
    ("nome_proprietario", 32, "input"),
    ("nome_imovel", 24, "input"),
    ("area_total_ha", 12, "input"),
    ("codigo_incra_nirf", 20, "input"),
    ("tipo_logradouro", 15, "input"),
    ("logradouro", 40, "input"),
    ("distrito", 16, "input"),
    ("uf", 6, "input"),
    ("municipio", 20, "input"),
    ("cep", 12, "input"),
    ("contribuinte_tipo", 16, "input"),
    ("cpf_cnpj", 18, "input"),
    ("condominio", 11, "input"),
    ("imune_ou_isento", 14, "input"),
    ("opcao_isencao", 13, "input"),
    ("categoria_vtn", 28, "input"),
    ("area_vtn_ha", 12, "input"),
    ("area_preservacao_permanente", 14, "input"),
    ("area_reserva_legal", 14, "input"),
    ("area_florestas_nativas", 14, "input"),
    ("area_pastagem", 12, "input"),
    ("area_benfeitorias", 12, "input"),
    ("area_nao_utilizada", 14, "input"),
    ("outras_areas", 12, "input"),
    ("grau_utilizacao_pct", 13, "input"),
    ("numero_car", 20, "input"),
    ("chave_vtn", 34, "cross"),
    ("valor_vtn_ha", 14, "cross"),
    ("vtn_total", 16, "calc"),
    ("alerta_isencao", 34, "calc"),
    ("status", 12, "input"),
    ("observacoes", 24, "input"),
]
COL = {nome: 2 + i for i, (nome, _, _) in enumerate(colunas)}  # indice da coluna
wc.column_dimensions["A"].width = 2
wc.row_dimensions[2].height = 30
wc.merge_cells(start_row=2, start_column=2, end_row=2, end_column=12)
wc.cell(row=2, column=2, value="Clientes / Imóveis Rurais — entrada do robô ITR").font = Font(size=14, bold=True)
wc.row_dimensions[4].height = 34
for nome, largura, tipo in colunas:
    j = COL[nome]
    wc.column_dimensions[get_column_letter(j)].width = largura
    c = wc.cell(row=4, column=j, value=nome)
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", start_color="333333")
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")

MAX_LINHA = 704  # 700 linhas de dados (5..704)
for i in range(5, MAX_LINHA + 1):
    # formulas
    wc.cell(row=i, column=COL["chave_vtn"],
            value=f'=IF(OR({get_column_letter(COL["municipio"])}{i}="",{get_column_letter(COL["categoria_vtn"])}{i}=""),"",'
                  f'{get_column_letter(COL["municipio"])}{i}&"|"&{get_column_letter(COL["categoria_vtn"])}{i})')
    wc.cell(row=i, column=COL["valor_vtn_ha"],
            value=f"=IFERROR(INDEX('VTN 2026'!$D$5:$D${ULTIMA_VTN},"
                  f"MATCH({get_column_letter(COL['chave_vtn'])}{i},'VTN 2026'!$F$5:$F${ULTIMA_VTN},0)),\"\")")
    q = get_column_letter(COL["area_vtn_ha"])
    ab = get_column_letter(COL["valor_vtn_ha"])
    wc.cell(row=i, column=COL["vtn_total"],
            value=f'=IF(OR({ab}{i}="",{q}{i}=""),"",{q}{i}*{ab}{i})')
    n = get_column_letter(COL["imune_ou_isento"])
    cc = get_column_letter(COL["area_total_ha"])
    wc.cell(row=i, column=COL["alerta_isencao"],
            value=f'=IF(AND({n}{i}="Sim",{cc}{i}<30),"ISENTO (<30 ha, imóvel único) — marcar opção A no Diac","")')
    # estilos
    for nome, _, tipo in colunas:
        cell = wc.cell(row=i, column=COL[nome])
        cell.border = BORDER
        cell.font = {"input": F_INPUT, "calc": F_CALC, "cross": F_CROSS}[tipo]
    wc.cell(row=i, column=COL["valor_vtn_ha"]).number_format = '"R$" #,##0.00'
    wc.cell(row=i, column=COL["vtn_total"]).number_format = '"R$" #,##0.00'
    for nm in ("area_total_ha", "area_vtn_ha", "area_preservacao_permanente",
               "area_reserva_legal", "area_florestas_nativas", "area_pastagem",
               "area_benfeitorias", "area_nao_utilizada", "outras_areas"):
        wc.cell(row=i, column=COL[nm]).number_format = '0.0'

# linha de exemplo (do video: FAZENDA TAPERAo)
ex = {
    "nome_imovel": "FAZENDA TAPERÃO", "area_total_ha": 5.9,
    "codigo_incra_nirf": "950114.051926-1", "tipo_logradouro": "Rodovia",
    "logradouro": "GO 080 PETROLINA A NEROPOLIS KM 06 MARGEM ESQUERDA",
    "distrito": "", "uf": "GO", "municipio": "Petrolina de Goiás",
    "cep": "75480-000", "contribuinte_tipo": "Pessoa Física",
    "cpf_cnpj": "457.698.501-91", "condominio": "Sim",
    "imune_ou_isento": "Não", "opcao_isencao": "",
    "categoria_vtn": "Silvicultura ou pastagem natural", "area_vtn_ha": 5.9,
    "area_preservacao_permanente": 0.0, "area_reserva_legal": 0.0,
    "area_florestas_nativas": 0.0, "area_pastagem": 5.9,
    "area_benfeitorias": 0.0, "area_nao_utilizada": 0.0, "outras_areas": 0.0,
    "grau_utilizacao_pct": 0, "numero_car": "",
    "status": "PENDENTE", "observacoes": "exemplo do vídeo",
}
for nome, valor in ex.items():
    wc.cell(row=5, column=COL[nome], value=valor)
    wc.cell(row=5, column=COL[nome]).font = F_INPUT

# validacoes (listas suspensas)
def dv(lista, colnome):
    d = DataValidation(type="list", formula1='"' + ",".join(lista) + '"', allow_blank=True)
    wc.add_data_validation(d)
    letra = get_column_letter(COL[colnome])
    d.add(f"{letra}5:{letra}{MAX_LINHA}")

dv(["GO"], "uf")
dv(["Pessoa Física", "Pessoa Jurídica"], "contribuinte_tipo")
dv(["Sim", "Não"], "condominio")
dv(["Sim", "Não"], "imune_ou_isento")
dv(["A", "B"], "opcao_isencao")
dv(CATEGORIAS, "categoria_vtn")
dv(MUNICIPIOS, "municipio")
dv(["PENDENTE", "PRONTO", "ERRO"], "status")
wc.freeze_panes = "C5"

# ============================== FONTES ==============================
wf = wb.create_sheet("Fontes")
wf.sheet_view.showGridLines = False
for col, w in zip("ABCD", [3, 40, 70, 3]):
    wf.column_dimensions[col].width = w
wf.row_dimensions[2].height = 30
wf["B2"] = "Fontes dos dados"
wf["B2"].font = Font(size=14, bold=True)
fontes = [
    ("Valores de VTN — Petrolina de Goiás (2026)",
     "Laudo Técnico para Determinação do Valor da Terra Nua (VTN) municipal de Petrolina de Goiás-GO, Tabela 1, pág. 18 — documento fornecido pelo usuário (IN RFB nº 1877/2019)"),
    ("Regras de preenchimento (taxa única, isenção <30 ha, importação 2025→2026)",
     "Transcrição de áudio do responsável pelo processo (arquivo 'audios manuel.txt')"),
    ("Fluxo de telas do programa ITR",
     "Gravações de tela 'PC Manoel.mp4' e 'PC Manoel 2.mp4' fornecidas pelo usuário"),
]
for i, (a, b) in enumerate(fontes):
    r = 4 + i
    wf.cell(row=r, column=2, value=a).font = Font(bold=True)
    wf.cell(row=r, column=2).border = BORDER
    wf.cell(row=r, column=3, value=b).border = BORDER

wb.save("planilha_clientes.xlsx")
print("OK - planilha_clientes.xlsx gerada")
