# Automação do Programa ITR

Robô de desktop que preenche as declarações no programa ITR da Receita a partir de uma planilha de imóveis, com cliques e digitação humanizados — a transmissão continua manual.

![Python](https://img.shields.io/badge/Python-3.12-3776AB) ![PyAutoGUI](https://img.shields.io/badge/PyAutoGUI-desktop-4B8BBE) ![PyInstaller](https://img.shields.io/badge/PyInstaller-.exe-FFD43B) ![status](https://img.shields.io/badge/status-Conclu%C3%ADdo-blue) ![licença](https://img.shields.io/badge/licen%C3%A7a-propriet%C3%A1ria-lightgrey)

> Projeto real. Este repositório é uma **cópia de portfólio**: segredos, servidores e nomes de empresas foram removidos ou trocados pelo nome fictício "Natela".

## Sumário

1. [O que o projeto faz](#1-o-que-o-projeto-faz)
2. [Instalação e requisitos](#2-instalação-e-requisitos)
3. [Como rodar e usar](#3-como-rodar-e-usar)
4. [Demo visual](#4-demo-visual)
5. [Suporte e ajuda](#5-suporte-e-ajuda)
6. [Contribuição e manutenção](#6-contribuição-e-manutenção)
7. [Licença e status](#7-licença-e-status)

## 1. O que o projeto faz

**O que resolve:** O programa ITR é uma aplicação Java de desktop sem importação em lote: cada imóvel exige preencher várias fichas à mão.

**Como resolve:** Um robô que lê a planilha de imóveis, calcula o VTN, e preenche as fichas Diac, Diat, área não utilizada e verificação de pendências, pausando a cada imóvel para conferência.

**Para quem é:** Escritórios que entregam muitas declarações de ITR por ano.

**Por que importa:** Tira o trabalho braçal sem tirar o controle: nada é transmitido sem a pessoa conferir.

### Principais benefícios

| Preenche as fichas sozinho | Calibração para qualquer tela | Gera .exe sem Python |
|---|---|---|

### Funcionalidades

- Leitura da planilha de imóveis com status PENDENTE/PRONTO
- Cálculo do VTN pela tabela do município
- Preenchimento Diac → Diat → Área Não Utilizada → Verificar Pendências
- Tratamento de imóvel isento (pula as fichas de área)
- Calibração das coordenadas da tela (`calibrar.py`) e log de tudo o que foi feito
- Geração de executável com PyInstaller e guia de uso em PDF

## 2. Instalação e requisitos

### Pré-requisitos

| Ferramenta | Versão |
|---|---|
| Python | 3.12 |
| Windows | com o programa ITR instalado |
| Bibliotecas | pyautogui, pyperclip, openpyxl |

### Instalação

```bash
git clone https://github.com/jhonatancamargo121/automacao-itr-desktop.git
cd automacao-itr-desktop
python -m pip install -r requirements.txt
python gerar_planilha.py
```

Sem `.env`: as coordenadas e a ordem dos passos ficam no `config.json` gerado pela calibração; os imóveis, na planilha criada por `gerar_planilha.py`.

### Variáveis de ambiente

| Variável | Para que serve |
|---|---|
| `config.json` | Coordenadas dos campos e ordem dos passos |
| `planilha_clientes.xlsx` | Imóveis a declarar (aba Clientes) e tabela de VTN |

A lista completa, com comentários, está no `.env.example`. Nenhum valor real é versionado.

## 3. Como rodar e usar

### Quick start

```bash
calibrar.bat
executar.bat
```

O robô espera 8 segundos para você focar a janela do ITR e então processa cada linha PENDENTE.

### Exemplo de uso

Requisição:

```bash
python automacao_itr.py
```

Resposta esperada:

```text
Para cada imóvel PENDENTE: preenche as fichas, pausa para você conferir e só então marca PRONTO na planilha. A transmissão é sempre manual.
```

## 4. Demo visual

Fluxo principal:

```mermaid
flowchart LR
  PL[Planilha de imóveis] --> V[Calcula VTN]
  V --> D[Diac]
  D --> T[Diat]
  T --> N[Área não utilizada]
  N --> P[Verificar pendências]
  P --> C[Pessoa confere] --> OK[Marca PRONTO]
```

> Prints e GIF com dados fictícios estão no roadmap. Para ver o sistema rodando, siga o [Quick start](#3-como-rodar-e-usar).

## 5. Suporte e ajuda

- **Documentação:**
  - [Guia rápido](docs/DETALHES.md)
- **Issues:** abra uma [issue](https://github.com/jhonatancamargo121/automacao-itr-desktop/issues) para reportar problemas ou sugerir melhorias.
- **Discussions:** dúvidas gerais em [Discussions](https://github.com/jhonatancamargo121/automacao-itr-desktop/discussions).
- **Contato:** pelo perfil [@jhonatancamargo121](https://github.com/jhonatancamargo121)

## 6. Contribuição e manutenção

### Como contribuir

1. Faça um **fork** do repositório
2. Crie uma **branch** a partir da `main` (`git checkout -b feat/minha-melhoria`)
3. Abra um **Pull Request** descrevendo o que mudou e por quê

### Padrões

- Commits no padrão [Conventional Commits](https://www.conventionalcommits.org/pt-br/) (`feat:`, `fix:`, `docs:`…)
- Lint e formatação passando antes do PR
- Testes rodando antes do PR

### Testes

Este projeto ainda não tem suíte automatizada; ela está no roadmap.

### Mantenedores

| | |
|---|---|
| **Jhonatan Camargo** ([@jhonatancamargo121](https://github.com/jhonatancamargo121)) | Mantenedor |

## 7. Licença e status

### Licença

Código proprietário, publicado para fins de portfólio. Todos os direitos reservados; veja o arquivo [LICENSE](LICENSE).

### Status

![status](https://img.shields.io/badge/status-Conclu%C3%ADdo-blue)

### Próximos passos

- [ ] Reconhecimento de campos por imagem no lugar de coordenadas fixas
- [ ] Integração com o motor de apuração do ITR
