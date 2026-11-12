# -*- coding: utf-8 -*-
"""
automacao_itr.py — Robô de preenchimento do programa ITR (Receita Federal).

O que ele faz:
  1. Lê a aba "Clientes" da planilha_clientes.xlsx (linhas com status = PENDENTE).
  2. Preenche as fichas do programa ITR (Diac, Diat, Área Não Utilizada) com
     cliques e digitação HUMANIZADOS (movimento em curva, velocidade variável).
  3. Clica em "Verificar Pendências" e PAUSA para você conferir.
  4. Marca a linha como PRONTO na planilha e passa para o próximo imóvel.

O que ele NUNCA faz:
  - Não transmite a declaração. A transmissão é sempre manual/sua.

Segurança:
  - Mova o mouse bruscamente para o CANTO SUPERIOR ESQUERDO da tela para abortar.
  - Ctrl+C na janela do console também interrompe.

Antes de usar: rode calibrar.py uma vez para ajustar as coordenadas à sua tela.
"""
import json
import math
import os
import random
import sys
import time
import traceback
from datetime import datetime

import pyautogui
import pyperclip
from openpyxl import load_workbook

pyautogui.FAILSAFE = True  # mouse no canto superior esquerdo = aborta

if getattr(sys, "frozen", False):
    # rodando como .exe (PyInstaller): usar a pasta do executável
    BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(BASE, "config.json")
LOG = os.path.join(BASE, "automacao_log.txt")


# ---------------------------------------------------------------- utilidades
def log(msg):
    linha = f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {msg}"
    print(linha)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(linha + "\n")


def carregar_config():
    with open(CONFIG, "r", encoding="utf-8") as f:
        return json.load(f)


# ------------------------------------------------------- movimento humanizado
def human_move(x, y):
    """Move o mouse em curva de Bézier com duração e desvio aleatórios."""
    x0, y0 = pyautogui.position()
    dist = math.hypot(x - x0, y - y0)
    dur = random.uniform(0.35, 0.85) + min(dist / 3000, 0.4)
    passos = max(18, int(dur * 60))
    # dois pontos de controle aleatórios -> curva natural
    c1 = (x0 + (x - x0) * random.uniform(0.2, 0.4) + random.uniform(-90, 90),
          y0 + (y - y0) * random.uniform(0.2, 0.4) + random.uniform(-90, 90))
    c2 = (x0 + (x - x0) * random.uniform(0.6, 0.8) + random.uniform(-90, 90),
          y0 + (y - y0) * random.uniform(0.6, 0.8) + random.uniform(-90, 90))
    for i in range(passos + 1):
        t = i / passos
        # easing suave no início/fim
        t = t * t * (3 - 2 * t)
        bx = ((1 - t) ** 3 * x0 + 3 * (1 - t) ** 2 * t * c1[0]
              + 3 * (1 - t) * t ** 2 * c2[0] + t ** 3 * x)
        by = ((1 - t) ** 3 * y0 + 3 * (1 - t) ** 2 * t * c1[1]
              + 3 * (1 - t) * t ** 2 * c2[1] + t ** 3 * y)
        pyautogui.moveTo(bx + random.uniform(-1.2, 1.2),
                         by + random.uniform(-1.2, 1.2),
                         duration=dur / passos)


def human_click(campo, campos):
    x, y = campos[campo]
    human_move(x + random.randint(-3, 3), y + random.randint(-2, 2))
    time.sleep(random.uniform(0.12, 0.35))
    pyautogui.mouseDown()
    time.sleep(random.uniform(0.04, 0.11))
    pyautogui.mouseUp()
    time.sleep(random.uniform(0.15, 0.45))


def digitar(texto):
    """Digita como humano; usa área de transferência para acentos."""
    texto = str(texto)
    if texto.isascii():
        for ch in texto:
            pyautogui.typewrite(ch, interval=random.uniform(0.03, 0.11))
    else:
        pyperclip.copy(texto)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(random.uniform(0.1, 0.3))


# ------------------------------------------------------------------ planilha
COLUNAS = None  # {nome_coluna: indice}


def ler_clientes(caminho):
    wb = load_workbook(caminho)
    ws = wb["Clientes"]
    global COLUNAS
    COLUNAS = {}
    for cell in ws[4]:
        if cell.value:
            COLUNAS[str(cell.value).strip()] = cell.column
    linhas = []
    for i in range(5, ws.max_row + 1):
        if ws.cell(row=i, column=COLUNAS["nome_imovel"]).value in (None, ""):
            continue
        dados = {}
        for nome, col in COLUNAS.items():
            v = ws.cell(row=i, column=col).value
            dados[nome] = "" if v is None else v
        linhas.append((i, dados))
    return wb, ws, linhas


def fmt(valor):
    """Formata número para o padrão do programa (vírgula decimal)."""
    if isinstance(valor, float):
        s = f"{valor:.2f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    return str(valor)


# -------------------------------------------------------------------- execução
def confirmar(titulo, mensagem):
    """Caixa de confirmação — o robô só continua se você clicar em Sim."""
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ok = messagebox.askyesno(titulo, mensagem, parent=root)
    root.destroy()
    return ok


def executar_linha(dados, cfg):
    campos = cfg["campos"]
    isento = str(dados.get("imune_ou_isento", "")).strip().lower() == "sim"
    for passo in cfg["etapas"]:
        if isento and passo.get("pular_se_isento"):
            continue
        tipo = passo["tipo"]
        if tipo == "click":
            human_click(passo["campo"], campos)
        elif tipo == "digitar":
            coluna = passo["coluna"]
            valor = fmt(dados.get(coluna, ""))
            if valor == "" and not passo.get("obrigatorio"):
                continue
            human_click(passo["campo"], campos)
            pyautogui.hotkey("ctrl", "a")
            digitar(valor)
        elif tipo == "click_radio":
            coluna = passo["coluna"]
            valor = str(dados.get(coluna, "")).strip()
            campo = passo["mapa"].get(valor)
            if campo:
                human_click(campo, campos)
        elif tipo == "tecla":
            pyautogui.press(passo["tecla"])
            time.sleep(random.uniform(0.2, 0.5))
        elif tipo == "esperar":
            time.sleep(passo["segundos"] * random.uniform(0.9, 1.3))
        else:
            log(f"  ! passo desconhecido ignorado: {tipo}")


def main():
    cfg = carregar_config()
    planilha = os.path.join(BASE, cfg.get("planilha", "planilha_clientes.xlsx"))
    wb, ws, linhas = ler_clientes(planilha)
    pendentes = [(i, d) for i, d in linhas
                 if str(d.get("status", "")).strip().upper() == "PENDENTE"]

    log("=" * 60)
    log(f"Início da automação — {len(pendentes)} imóvel(is) PENDENTE(s).")
    if not pendentes:
        log("Nada a fazer. Encerrando.")
        return

    log("Deixe o programa ITR ABERTO na tela inicial e a janela MAXIMIZADA.")
    log("Você tem 8 segundos para trocar para a janela do ITR...")
    log("(Para abortar a qualquer momento: mouse no canto superior esquerdo.)")
    time.sleep(8)

    processados = 0
    for linha, dados in pendentes:
        nome = dados.get("nome_imovel", "?")
        try:
            log(f"--> Linha {linha}: {nome}")
            executar_linha(dados, cfg)
            ok = confirmar(
                "Conferir antes de continuar",
                f"Imóvel: {nome}\n\n"
                "O robô clicou em 'Verificar Pendências'.\n"
                "Confira no programa ITR se está tudo certo.\n\n"
                "SIM  = marcar como PRONTO e ir para o próximo\n"
                "NÃO  = marcar como ERRO e ir para o próximo\n\n"
                "Lembre-se: a TRANSMISSÃO é sempre manual.")
            ws.cell(row=linha, column=COLUNAS["status"],
                    value="PRONTO" if ok else "ERRO")
            wb.save(planilha)
            processados += 1
            log(f"    {'PRONTO' if ok else 'ERRO'} — planilha atualizada.")
        except pyautogui.FailSafeException:
            log("ABORTADO pelo usuário (mouse no canto). Planilha salva.")
            wb.save(planilha)
            break
        except Exception:
            log("ERRO na linha " + str(linha) + ":\n" + traceback.format_exc())
            ws.cell(row=linha, column=COLUNAS["status"], value="ERRO")
            wb.save(planilha)

    log(f"Fim. {processados} imóvel(is) processado(s).")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log("FALHA GERAL:\n" + traceback.format_exc())
        sys.exit(1)
