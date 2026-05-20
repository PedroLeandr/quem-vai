# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Projeto

<!-- Descreve o que este projeto faz, o seu objetivo principal e o público-alvo. -->

The project will be called "Quem Vai?". It will be a Windows application designed specifically for coach Luis Reis, so the app must be extremely simple, practical, and fast to use during activities.

The "Quem Vai?" app should allow the user to:

* Define the number of teams
* Name each team
* Define the number of youth players (CEF) and junior players (CP) for each team
* Add the names of all participants

The application must randomly select participants for activities.

Whenever participants are selected, the app should display a list containing:

* The participant's name
* Their respective team

Once a participant has been selected, they should not be eligible for selection again until every participant from the same team has already participated.

The main goal of the application is to provide a very intuitive and efficient tool for quickly organizing random selections during training activities.



## Stack Tecnológica

- **GUI:** CustomTkinter — interface moderna e componentes com boa legibilidade (botões grandes, texto claro)
- **Linguagem:** Python 3.11+
- **Base de dados:** SQLite (via `sqlite3` — incluído no Python) — persiste equipas e jogadores entre sessões
- **Distribuição:** PyInstaller — compila `quem-vai.py` num único `quem-vai.exe` sem necessitar de Python instalado
- **Dependências:** `customtkinter` (única dependência externa)

## Comandos

```bash
pip install customtkinter          # instalar dependência
python quem-vai.py                 # correr a app em desenvolvimento

# compilar para .exe (Windows)
pip install pyinstaller
pyinstaller --onefile --windowed --name quem-vai quem-vai.py
# o executável fica em dist/quem-vai.exe
```

## Arquitetura

A app é um ficheiro único `quem-vai.py`. A lógica está separada da interface:

- **Modelo de dados**: equipas e jogadores (CEF/CP) persistidos em SQLite (`quem-vai.db` gerado automaticamente na primeira execução); estado de sorteio vive apenas em memória e reseta a cada sessão
- **Lógica de seleção**: por equipa, um jogador só volta ao sorteio depois de todos os da mesma equipa terem sido selecionados — implementado com uma fila que se reinicia quando esgota
- **Interface CustomTkinter**: ecrãs sequenciais (configuração → participantes → sorteio → resultado)

## Convenções

- Toda a interface usa fontes grandes (mínimo 14px) e botões de altura ≥ 40px — a app é usada em contexto de atividade física, não em secretária
- Sem base de dados nem ficheiros externos — o estado vive em memória durante a sessão
