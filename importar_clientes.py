# -*- coding: utf-8 -*-
"""Importa os clientes do Cliente.csv para a aba Clientes da planilha do robô."""
import csv

from openpyxl import load_workbook
from openpyxl.styles import Font

F_INPUT = Font(color="0066CC")

rows = list(csv.reader(open("Cliente.csv", encoding="utf-8-sig"), delimiter=";"))
cab, dados = rows[0], rows[1:]

wb = load_workbook("planilha_clientes.xlsx")
ws = wb["Clientes"]
COL = {str(c.value).strip(): c.column for c in ws[4] if c.value}

def setv(linha, coluna, valor):
    if coluna in COL and valor not in (None, ""):
        c = ws.cell(row=linha, column=COL[coluna], value=valor)
        c.font = F_INPUT

linha = 6  # linha 5 = exemplo do video
importados = 0
for r in dados:
    if len(r) < 19:
        continue
    nome = r[0].strip()
    if not nome:
        continue
    cnpj, cpf = r[2].strip(), r[3].strip()
    setv(linha, "nome_proprietario", nome)
    if cpf:
        setv(linha, "contribuinte_tipo", "Pessoa Física")
        setv(linha, "cpf_cnpj", cpf)
    elif cnpj:
        setv(linha, "contribuinte_tipo", "Pessoa Jurídica")
        setv(linha, "cpf_cnpj", cnpj)
    setv(linha, "cep", r[13].strip())
    setv(linha, "uf", r[14].strip())
    setv(linha, "municipio", r[15].strip())
    obs = []
    end = " ".join(x for x in [r[16].strip(), r[17].strip()] if x)
    if end:
        obs.append("End.: " + end + (f" - {r[18].strip()}" if r[18].strip() else ""))
    if r[11].strip():
        obs.append("Fone: " + r[11].strip())
    if r[10].strip():
        obs.append("Email: " + r[10].strip())
    if r[6].strip():
        obs.append("Obs.: " + r[6].strip())
    setv(linha, "observacoes", " | ".join(obs))
    setv(linha, "status", "PENDENTE")
    linha += 1
    importados += 1

wb.save("planilha_clientes.xlsx")
print(f"OK - {importados} clientes importados a partir da linha 6")
