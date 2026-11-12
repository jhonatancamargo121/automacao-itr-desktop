# Automação ITR — Guia rápido

Kit para preencher declarações do programa **ITR (Receita Federal)** automaticamente
a partir da planilha de clientes, com cliques e digitação humanizados.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `planilha_clientes.xlsx` | Cadastro dos imóveis (aba **Clientes**) + tabela **VTN 2026** |
| `automacao_itr.py` | O robô de preenchimento |
| `calibrar.py` | Ajuste das coordenadas dos campos para a sua tela (rodar 1x) |
| `config.json` | Coordenadas e ordem dos passos do robô |
| `executar.bat` | Roda o robô com um clique (precisa do Python instalado) |
| `calibrar.bat` | Roda a calibração com um clique |
| `gerar_exe.bat` | Gera o `AutomacaoITR.exe` (não precisa de Python depois) |
| `automacao_log.txt` | Registro de tudo que o robô fez (criado na 1ª execução) |

## Passo a passo (primeira vez)

1. **Instale as dependências** (uma vez só):
   ```
   python -m pip install -r requirements.txt
   ```
2. **Preencha a planilha** `planilha_clientes.xlsx`, aba **Clientes** — uma linha por
   imóvel. Deixe `status = PENDENTE` nas linhas a processar. As colunas de VTN
   (`valor_vtn_ha`, `vtn_total`, `alerta_isencao`) são calculadas sozinhas.
3. **Complete a aba VTN 2026** com os laudos dos demais municípios
   (Petrolina de Goiás já está preenchida).
4. **Abra o programa ITR maximizado** e rode `calibrar.bat` para ajustar as
   coordenadas à sua tela (as que vêm no `config.json` foram estimadas dos vídeos
   em 1280x720 — na sua máquina podem ser diferentes).
5. Rode `executar.bat`. O robô espera 8 segundos para você focar a janela do ITR
   e então processa cada linha PENDENTE.

## Durante a execução

- O robô preenche **Diac → Diat (Utilização do Imóvel) → Área Não Utilizada →
  Verificar Pendências**.
- Imóvel **isento** (`imune_ou_isento = Sim`, <30 ha): o robô marca a opção no
  Diac e **pula** as fichas de área (como o sistema faz ao marcar a opção A).
- Ao final de cada imóvel ele **pausa** e pede sua confirmação antes de marcar
  PRONTO na planilha.
- **A transmissão nunca é feita pelo robô** — é sempre manual, após sua conferência.

## Segurança / abortar

- Mova o mouse com força para o **canto superior esquerdo** da tela → aborta na hora.
- Ou feche a janela do console / Ctrl+C.
- A planilha é salva a cada linha — nada se perde.

## Gerar o .exe (opcional)

Rode `gerar_exe.bat`. Depois copie `dist\AutomacaoITR.exe` junto com
`planilha_clientes.xlsx` e `config.json` para qualquer pasta — não precisa de Python.

## Observações importantes

- O programa ITR precisa estar **sempre na mesma posição/tamanho de tela**
  (maximizado) entre calibração e execução.
- Se a Receita atualizar a versão do ITR e as telas mudarem, rode `calibrar.bat`
  de novo e, se necessário, ajuste a ordem dos passos no `config.json`.
- Regras aplicadas conforme o áudio do responsável: taxa única de VTN,
  sobra de área como terra não utilizável, isenção para imóvel único < 30 ha
  (opção A), declaração 2026 nova importando dados de 2025.
