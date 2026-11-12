# -*- coding: utf-8 -*-
"""Gera o guia de uso (PDF) da automação ITR — linguagem simples para o cliente."""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, KeepTogether)

AZUL = HexColor("#1F4E79")
CINZA_TXT = HexColor("#333333")
CINZA_CLARO = HexColor("#666666")

doc = SimpleDocTemplate(
    "Guia_de_Uso_Automacao_ITR.pdf", pagesize=A4,
    topMargin=2.5 * cm, bottomMargin=2.5 * cm, leftMargin=3 * cm, rightMargin=2.5 * cm)
doc.title = "Guia de Uso — Automação ITR"
doc.author = "Automação ITR"

ss = getSampleStyleSheet()
titulo = ParagraphStyle("titulo", parent=ss["Title"], fontName="Helvetica-Bold",
                        fontSize=22, leading=28, textColor=AZUL, alignment=1)
subtitulo = ParagraphStyle("subtitulo", parent=ss["Normal"], fontName="Helvetica",
                           fontSize=12, leading=17, textColor=CINZA_CLARO, alignment=1)
h1 = ParagraphStyle("h1", parent=ss["Heading1"], fontName="Helvetica-Bold",
                    fontSize=15, leading=19, textColor=AZUL, spaceBefore=16, spaceAfter=8,
                    keepWithNext=1)
corpo = ParagraphStyle("corpo", parent=ss["Normal"], fontName="Helvetica",
                       fontSize=11.5, leading=17, textColor=CINZA_TXT, spaceAfter=6)
passo = ParagraphStyle("passo", parent=corpo, leftIndent=22, firstLineIndent=-22, spaceAfter=7)
aviso = ParagraphStyle("aviso", parent=corpo, fontName="Helvetica-Bold",
                       textColor=HexColor("#8B0000"))
cap = ParagraphStyle("cap", parent=corpo, fontSize=9.5, textColor=CINZA_CLARO, alignment=1)


def header_footer(canvas, doc_):
    canvas.saveState()
    w, h = A4
    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(CINZA_CLARO)
    canvas.drawCentredString(w / 2, h - 1.4 * cm, "Guia de Uso — Automação ITR")
    canvas.line(3 * cm, h - 1.6 * cm, w - 2.5 * cm, h - 1.6 * cm)
    canvas.drawCentredString(w / 2, 1.4 * cm, f"Página {doc_.page}")
    canvas.restoreState()


story = []

# ---------------- capa ----------------
story.append(Spacer(1, 5 * cm))
story.append(Paragraph("Guia de Uso", titulo))
story.append(Spacer(1, 0.4 * cm))
story.append(Paragraph("Robô de Preenchimento do ITR (Receita Federal)", ParagraphStyle(
    "t2", parent=titulo, fontSize=16, leading=22)))
story.append(Spacer(1, 1 * cm))
story.append(Paragraph("Passo a passo simples — não precisa entender de computador", subtitulo))
story.append(Spacer(1, 6 * cm))
story.append(Paragraph("Versão 1.0 — Agosto/2026", cap))
story.append(PageBreak())

# ---------------- o que é ----------------
story.append(Paragraph("1. O que esse programa faz?", h1))
story.append(Paragraph(
    "Ele é um <b>robô</b> que preenche sozinho as declarações do programa "
    "<b>ITR da Receita Federal</b> (Imposto Territorial Rural), usando os dados "
    "de uma planilha do Excel. Ele move o mouse e digita como uma pessoa, "
    "campo por campo, imóvel por imóvel.", corpo))
story.append(Paragraph(
    "Você só precisa: <b>preencher a planilha</b>, <b>abrir o programa do ITR</b> e "
    "<b>dar dois cliques no robô</b>. No final de cada imóvel, ele pausa para você "
    "conferir antes de continuar.", corpo))
story.append(Spacer(1, 0.3 * cm))
story.append(Paragraph(
    "Importante: o robô NUNCA envia (transmite) a declaração para a Receita. "
    "O envio final é sempre feito por você, depois de conferir.", aviso))

# ---------------- o que vem na pasta ----------------
celula = ParagraphStyle("celula", parent=corpo, fontSize=10.5, leading=14, spaceAfter=0)
celula_b = ParagraphStyle("celulab", parent=celula, fontName="Helvetica-Bold")

def tabular(dados_tab):
    return [[Paragraph(str(c), celula_b if i == 0 else celula) for c in lin]
            for i, lin in enumerate(dados_tab)]

story.append(Paragraph("2. O que tem na pasta?", h1))
tabela = Table(tabular([
    ["Arquivo", "O que é"],
    ["AutomacaoITR.exe", "O robô. É ele que você executa (dois cliques)."],
    ["planilha_clientes.xlsx", "A planilha onde ficam os dados dos clientes e imóveis."],
    ["config.json", "Arquivo de ajustes do robô. Não precisa mexer."],
]), colWidths=[5.0 * cm, 10.3 * cm])
tabela.setStyle(TableStyle([
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
    ("FONTSIZE", (0, 0), (-1, -1), 10.5),
    ("LINEABOVE", (0, 0), (-1, 0), 1.5, HexColor("#000000")),
    ("LINEBELOW", (0, 0), (-1, 0), 0.75, HexColor("#000000")),
    ("LINEBELOW", (0, -1), (-1, -1), 1.5, HexColor("#000000")),
    ("TOPPADDING", (0, 0), (-1, -1), 6),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
]))
story.append(tabela)

# ---------------- passo a passo ----------------
story.append(Paragraph("3. Passo a passo para usar", h1))
passos = [
    ("Passo 1 — Abra a planilha",
     "Dê dois cliques em <b>planilha_clientes.xlsx</b> e vá na aba <b>Clientes</b>."),
    ("Passo 2 — Preencha os dados do imóvel",
     "Encontre a linha do cliente (pelo nome do proprietário) e preencha as colunas azuis: "
     "nome do imóvel (fazenda), área total, código do Incra (NIRF), endereço, município, "
     "áreas de pastagem/preservação etc. As colunas de VTN calculam sozinhas — não mexa nelas."),
    ("Passo 3 — Confira a coluna status",
     "A coluna <b>status</b> deve estar como <b>PENDENTE</b>. O robô só trabalha nas linhas "
     "que estão PENDENTE e têm o nome do imóvel preenchido."),
    ("Passo 4 — Salve e feche a planilha",
     "Salve (Ctrl+S) e <b>feche o Excel</b>. Se a planilha ficar aberta, o robô não consegue "
     "gravar o resultado."),
    ("Passo 5 — Abra o programa do ITR",
     "Abra o programa <b>ITR</b> da Receita Federal e deixe a janela <b>maximizada</b> "
     "(ocupando a tela inteira), na tela inicial."),
    ("Passo 6 — Execute o robô",
     "Dê dois cliques em <b>AutomacaoITR.exe</b>. Vai abrir uma janela preta (console). "
     "Você tem <b>8 segundos</b> para clicar na janela do ITR e deixá-la em primeiro plano."),
    ("Passo 7 — Acompanhe sem mexer no computador",
     "O robô vai mover o mouse e digitar sozinho. <b>Não toque no mouse nem no teclado</b> "
     "enquanto ele trabalha."),
    ("Passo 8 — Confira cada imóvel",
     "Ao terminar cada imóvel, o robô clica em <b>Verificar Pendências</b> e abre uma "
     "pergunta na tela. Confira no ITR se está tudo certo e clique em <b>Sim</b> "
     "(marca PRONTO e vai para o próximo) ou <b>Não</b> (marca ERRO e vai para o próximo)."),
    ("Passo 9 — Transmita manualmente",
     "Depois que o robô terminar, revise no ITR e faça a <b>transmissão você mesmo</b>, "
     "como sempre fez. O robô nunca transmite."),
]
for tit, txt in passos:
    story.append(Paragraph(f"<b>{tit}.</b> {txt}", passo))

story.append(PageBreak())

# ---------------- se algo der errado ----------------
story.append(Paragraph("4. Se algo der errado", h1))
tabela2 = Table(tabular([
    ["Situação", "O que fazer"],
    ["Quero parar o robô imediatamente",
     "Arraste o mouse com força para o CANTO SUPERIOR ESQUERDO da tela. O robô para na hora. "
     "Ou feche a janela preta (console)."],
    ["O robô clicou no lugar errado",
     "Pare o robô (veja acima), confira se a janela do ITR está maximizada e rode de novo. "
     "Nada se perde: a planilha é salva a cada imóvel."],
    ["O robô não faz nada",
     "Confira se há linhas com status PENDENTE e nome do imóvel preenchido na planilha. "
     "E se o ITR está aberto."],
    ["Errei os dados de um imóvel",
     "Corrija na planilha, mude o status de volta para PENDENTE e rode o robô de novo."],
    ["Fechei tudo sem querer",
     "Sem problema. Abra a planilha e veja a coluna status: o que está PRONTO já foi feito. "
     "Rode o robô de novo para o restante."],
]), colWidths=[5.2 * cm, 10.1 * cm])
tabela2.setStyle(TableStyle([
    ("LINEABOVE", (0, 0), (-1, 0), 1.5, HexColor("#000000")),
    ("LINEBELOW", (0, 0), (-1, 0), 0.75, HexColor("#000000")),
    ("LINEBELOW", (0, -1), (-1, -1), 1.5, HexColor("#000000")),
    ("TOPPADDING", (0, 0), (-1, -1), 6),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
]))
story.append(tabela2)

# ---------------- regras ----------------
story.append(Paragraph("5. Regrinhas de ouro", h1))
regras = [
    "Sempre deixe a janela do ITR <b>maximizada</b> antes de rodar o robô.",
    "Sempre <b>feche o Excel</b> antes de rodar o robô.",
    "Não mexa no mouse e no teclado enquanto o robô trabalha.",
    "A <b>transmissão</b> da declaração é sempre manual — o robô nunca envia nada.",
    "Imóvel de proprietário único com <b>menos de 30 hectares</b> é isento: marque "
    "\"Sim\" na coluna imune_ou_isento que o robô já cuida do resto.",
    "O arquivo <b>automacao_log.txt</b> guarda o histórico de tudo que o robô fez.",
]
for regra in regras:
    story.append(Paragraph("• " + regra, passo))

story.append(Spacer(1, 0.6 * cm))
story.append(Paragraph("6. Dúvidas?", h1))
story.append(Paragraph(
    "Se travar ou aparecer alguma mensagem estranha, anote o que estava escrito na "
    "janela preta (ou tire uma foto da tela) e chame o suporte. O histórico em "
    "<b>automacao_log.txt</b> ajuda a descobrir o que aconteceu.", corpo))

doc.build(story, onFirstPage=lambda c, d: None, onLaterPages=header_footer)
print("OK - Guia_de_Uso_Automacao_ITR.pdf gerado")
