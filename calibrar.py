# -*- coding: utf-8 -*-
"""
calibrar.py — Ajusta as coordenadas dos campos do programa ITR para a SUA tela.

Como usar:
  1. Abra o programa ITR, MAXIMIZADO, em uma declaração na ficha do Imóvel (Diac).
  2. Rode este script.
  3. Para cada campo pedido: posicione o mouse em cima do campo no ITR
     e pressione ENTER aqui no console (sem clicar no campo!).
  4. No final ele grava tudo no config.json.

Dica: aperte ENTER sem mover o mouse. Se errar um campo, rode de novo —
os valores anteriores ficam como padrão (ENTER com o mouse no lugar errado
basta não fazer; Ctrl+C sai sem salvar).
"""
import json
import os

import pyautogui

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(BASE, "config.json")

with open(CONFIG, "r", encoding="utf-8") as f:
    cfg = json.load(f)

campos = cfg["campos"]
print("=" * 64)
print("CALIBRAÇÃO DAS COORDENADAS — ITR")
print("=" * 64)
print("Deixe o ITR aberto e maximizado. Para cada campo:")
print("  -> leve o mouse até o campo no programa ITR (sem clicar)")
print("  -> volte aqui e pressione ENTER")
print("  -> ou pressione ENTER direto para MANTER o valor atual")
print("  -> Ctrl+C para sair sem salvar")
print()

novos = dict(campos)
for nome in campos:
    atual = campos[nome]
    resp = input(f"[{nome}]  atual={atual}  (Enter=mantém / posicione o mouse e Enter=grava) ")
    if resp.strip() == "":
        # se o mouse foi movido para longe do console, pergunta se quer gravar
        x, y = pyautogui.position()
        print(f"    mouse está em ({x}, {y})")
        conf = input("    gravar esta posição? [s/N] ").strip().lower()
        if conf == "s":
            novos[nome] = [x, y]
            print(f"    gravado: [{x}, {y}]")
        else:
            print("    mantido.")

cfg["campos"] = novos
with open(CONFIG, "w", encoding="utf-8") as f:
    json.dump(cfg, f, ensure_ascii=False, indent=2)
print()
print("Calibração salva em config.json. Pode rodar a automação (executar.bat).")
