# Tutorial SGPilot — Guia Completo de Uso

> **Versão:** 5.1.0 | **Plataforma:** Windows | **Autor:** Vitor Glennon

---

## Sumário

1. [Instalação](#1-instalação)
2. [Primeira Execução](#2-primeira-execução)
3. [Conectando ao Chrome](#3-conectando-ao-chrome)
4. [Aba SGP — Automação de Ocorrências](#4-aba-sgp--automação-de-ocorrências)
5. [Aba Papervines — Loop de Atendimento](#5-aba-papervines--loop-de-atendimento)
6. [Configurando Hotkeys (Binds)](#6-configurando-hotkeys-binds)
7. [Tipos de Ação](#7-tipos-de-ação)
8. [Gerando o Executável `.exe`](#8-gerando-o-executável-exe)
9. [Solução de Problemas](#9-solução-de-problemas)

---

## 1. Instalação

### Requisitos mínimos

- Windows 10 ou superior
- Python 3.10+ ([baixar aqui](https://www.python.org/downloads/))
- Google Chrome instalado

### Opção A — Instalação via Python (recomendado para desenvolvimento)

```bash
# Clone o repositório
git clone https://github.com/vitorstewartglennon30/SGPilot.git
cd SGPilot

# Instale as dependências
pip install -r requirements.txt
```

### Opção B — Instalação pelo instalador Windows

1. Baixe o projeto ou clone com git
2. Dê duplo clique em **`instalar.bat`**
3. Aguarde a instalação terminar
4. Pronto!

### Dependências instaladas

O SGPilot usa as seguintes bibliotecas Python:

- `customtkinter` — interface gráfica moderna
- `selenium` — conexão com o Chrome via DevTools
- `keyboard` — hotkeys globais
- `pyautogui` + `pyperclip` — envio de texto via clipboard
- `Pillow` — exibição de imagens na interface

---

## 2. Primeira Execução

Na primeira vez que o SGPilot é aberto, ele cria automaticamente o arquivo `config.json` na mesma pasta, com as configurações padrão (hotkeys F1–F7 pré-configuradas).

### Como abrir o SGPilot

**Via Python:**
```bash
python main.py
```

**Via executável (se você gerou o `.exe`):**
```
Dê duplo clique em SGPilot.exe
```

A janela do SGPilot será aberta com duas abas: **SGP** e **Papervines**.

---

## 3. Conectando ao Chrome

O SGPilot controla o Chrome através do **protocolo DevTools** — isso significa que você não precisa instalar nenhuma extensão. Mas é necessário iniciar o Chrome de uma forma especial.

### Passo 1 — Feche todo o Chrome

Certifique-se de que **nenhuma janela do Chrome está aberta**. Verifique no Gerenciador de Tarefas se não há processos `chrome.exe` rodando.

### Passo 2 — Inicie o Chrome em modo debug

Dê duplo clique em **`chrome_debug.bat`**.

Isso executa o Chrome com o seguinte comando:
```
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\sgpilot-chrome-profile"
```

Uma janela do Chrome abrirá normalmente — use ela para acessar o SGP e o Papervines.

> **Por que um perfil separado?** O parâmetro `--user-data-dir` cria um perfil de Chrome exclusivo para o SGPilot, evitando conflitos com seu perfil pessoal.

### Passo 3 — Conecte no SGPilot

1. Abra o SGPilot
2. Clique no botão **"Conectar ao Chrome"** (no topo da interface)
3. Se a conexão for bem-sucedida, o status ficará **verde**

> Se a conexão falhar, verifique se o Chrome foi iniciado via `chrome_debug.bat` e se a porta `9222` está correta nas configurações.

---

## 4. Aba SGP — Automação de Ocorrências

A aba SGP é o coração do SGPilot. Ela permite automatizar o preenchimento de ocorrências com um único toque de tecla.

### Interface da aba SGP

```
┌──────────────────────────────────────────────────┐
│  [SGP]              [Papervines]                 │
├──────────────────────────────────────────────────┤
│  Status: ● Conectado                             │
│                                                  │
│  [F1] OFF LOSI/LOBI          [F2] OFF FTTx       │
│  [F3] Offline Dying-gasp     [F4] Sinal Atenuado │
│  [F6] Comprovante (SGP)      [F7] Link Pagamento │
│                                                  │
│  [+ Nova Bind]    [Configurações]                │
└──────────────────────────────────────────────────┘
```

### Como usar uma hotkey

1. Abra a ficha do cliente no SGP no Chrome
2. Pressione a hotkey correspondente (ex: `F1`)
3. O SGPilot lê o número do contrato diretamente da página
4. Executa a ação configurada (envia mensagem, preenche formulário, etc.)

### Hotkeys padrão

| Tecla | Nome | Tipo |
|-------|------|------|
| F1 | OFF LOSI/LOBI | Texto + Número |
| F2 | OFF FTTx com potência | Texto + Número |
| F3 | Offline Dying-gasp | Texto + Número |
| F4 | Sinal atenuado | Texto + Número |
| F6 | Comprovante (SGP completo) | Automação SGP |
| F7 | Link de pagamento | Link de Pagamento |

### Fluxo de automação SGP completo (ex: F6)

Quando uma bind do tipo **"Automação SGP"** é acionada:

1. O SGPilot localiza a página de nova ocorrência no Chrome
2. Preenche o campo **Tipo** (ex: financeiro/sus)
3. Preenche o campo **Origem** (ex: whatsapp)
4. Preenche o campo **Conteúdo** com a mensagem configurada
5. Opcionalmente: clica em **Cadastrar** automaticamente
6. Opcionalmente: abre e preenche a **OS** automaticamente

---

## 5. Aba Papervines — Loop de Atendimento

A aba Papervines automatiza o início do atendimento de novos clientes na fila.

### Como funciona o loop

1. Pressione `F9` (ou a tecla configurada) para **iniciar** o loop
2. O SGPilot executa em sequência:
   - Clica em **"Novos"** para ver a fila de clientes aguardando
   - Clica no primeiro cliente da lista
   - Clica em **"Iniciar"** para começar o atendimento
   - Digite a **mensagem de saudação** no chat
   - Clica em **Enviar**
   - Aguarda o delay configurado
   - Repete para o próximo cliente
3. Pressione `F9` novamente para **parar** o loop

### Configurações do Papervines

Clique em **"Configurações"** na aba Papervines para ajustar:

| Configuração | Descrição | Padrão |
|---|---|---|
| Saudação | Mensagem enviada para cada cliente | "Olá! Tudo bem? Meu nome é..." |
| Delay entre clientes | Tempo de espera entre cada cliente (ms) | 1500ms |
| Tecla de início | Hotkey para iniciar/parar o loop | F9 |

### Log em tempo real

A aba Papervines exibe um log com timestamps de cada ação executada, facilitando o acompanhamento e a identificação de problemas:

```
[10:32:01] Iniciando loop Papervines...
[10:32:02] Cliente 1: João Silva — Saudação enviada ✓
[10:32:04] Cliente 2: Maria Souza — Saudação enviada ✓
[10:32:06] Fila vazia. Loop encerrado.
```

---

## 6. Configurando Hotkeys (Binds)

Você pode criar, editar e remover hotkeys clicando em qualquer bind existente ou em **"+ Nova Bind"**.

### Campos de uma bind

| Campo | Descrição |
|---|---|
| **Nome** | Nome exibido no botão (ex: "Sinal Fraco") |
| **Tecla** | Hotkey do teclado (ex: F5, F8, Ctrl+1) |
| **Ativada** | Liga/desliga a hotkey sem excluí-la |
| **Cor** | Cor do botão na interface |
| **Tipo(s) de ação** | O que acontece ao pressionar a tecla |
| **Mensagem** | Texto enviado no chat (suporta `{numero}` e `{mes}`) |
| **Filtro Tipo SGP** | Filtra o tipo de ocorrência (ex: `sus`, `finan`) |
| **Filtro Origem SGP** | Filtra a origem (ex: `whatsapp`, `telefone`) |
| **Auto Cadastrar** | Cadastra a ocorrência automaticamente após preencher |
| **Auto Cadastrar OS** | Abre e cadastra a OS automaticamente |

### Variáveis disponíveis na mensagem

| Variável | O que substitui |
|---|---|
| `{numero}` | Número do contrato lido diretamente do SGP |
| `{mes}` | Mês atual por extenso (ex: "abril") |
| `{link}` | Link colado da área de transferência (usado no link de pagamento) |

### Exemplo: criar uma nova bind

1. Clique em **"+ Nova Bind"**
2. Preencha:
   - **Nome:** "Sem Sinal FTTx"
   - **Tecla:** F8
   - **Tipo:** Texto + Número
   - **Mensagem:** `SEM SINAL FTTx\n\n{numero}`
3. Clique em **Salvar**

A nova hotkey já estará ativa imediatamente.

---

## 7. Tipos de Ação

O SGPilot suporta 4 tipos de ação por bind (podem ser combinados):

### 1. Enviar texto no chat (`text`)

Envia uma mensagem fixa no chat do Papervines ou SGP, sem ler nenhum dado da página.

**Ideal para:** saudações, respostas padrão, mensagens fixas.

```
Mensagem: "Obrigado por entrar em contato! Estou verificando sua situação."
```

### 2. Texto + Número (`text_ocr`)

Lê o número do contrato diretamente do HTML do SGP e insere na mensagem usando `{numero}`.

**Ideal para:** registros de ocorrência que precisam do número do cliente.

```
Mensagem: "OFF LOSI/LOBI\n\n{numero}"
→ Resultado: "OFF LOSI/LOBI\n\n123456"
```

### 3. Automação SGP (`sgp_ocorrencia`)

Preenche automaticamente o formulário de nova ocorrência no SGP:
- Tipo de ocorrência (financeiro, técnico, etc.)
- Origem (whatsapp, telefone, etc.)
- Conteúdo/descrição
- Opcionalmente cadastra e abre OS

**Ideal para:** fluxos de ocorrência que se repetem muito (ex: comprovante de pagamento).

### 4. Link de pagamento (`link_pagamento`)

Fluxo em 2 etapas:
1. **Etapa 1:** O SGPilot aguarda você copiar o link de pagamento (Ctrl+C)
2. **Etapa 2:** Pressiona a hotkey novamente para enviar a mensagem com o link inserido automaticamente

**Ideal para:** envio de links de boleto/pix no chat.

---

## 8. Gerando o Executável `.exe`

Para distribuir o SGPilot sem precisar do Python instalado:

### Via script (recomendado)

Dê duplo clique em **`build.bat`**.

### Manualmente

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name SGPilot --icon sgpilot.ico ^
  --add-data "logo.png;." ^
  --add-data "logo2.png;." ^
  --collect-all customtkinter ^
  main.py
```

O arquivo `dist/SGPilot.exe` será gerado. Para distribuir, copie também:
- `logo.png`, `logo2.png`, `logo3.png`
- `icons/` (pasta inteira)
- `chrome_debug.bat`
- `instalar.bat` (se o destino não tiver Python)

> O `config.json` será criado automaticamente na primeira execução no computador de destino.

---

## 9. Solução de Problemas

### "Não foi possível conectar ao Chrome"

**Causa:** O Chrome não foi iniciado em modo debug, ou a porta está diferente.

**Solução:**
1. Feche todo o Chrome
2. Execute `chrome_debug.bat`
3. Abra o SGPilot e clique em "Conectar ao Chrome"

Verifique se a porta no `config.json` (`sgp.debug_port`) é `9222`.

---

### "Hotkey não funciona"

**Causa:** O SGPilot não tem permissão para capturar teclas globais, ou a tecla está sendo capturada por outro programa.

**Solução:**
1. Execute o SGPilot como **Administrador** (clique direito → Executar como administrador)
2. Verifique se nenhum outro programa está usando a mesma hotkey

---

### "O formulário SGP não é preenchido corretamente"

**Causa:** A página do SGP pode ter carregado antes do Selenium estar pronto, ou a estrutura HTML mudou.

**Solução:**
1. Aumente o `delay_ms` nas configurações (ex: de 150 para 300)
2. Certifique-se de estar na página de ocorrência correta antes de pressionar a hotkey

---

### "O loop Papervines para sozinho"

**Causa:** Não há mais clientes na fila ou ocorreu um erro ao localizar o botão.

**Solução:**
1. Verifique o log na aba Papervines para identificar o erro
2. Aumente o `delay_entre_clientes_ms` para dar mais tempo de carregamento

---

### Log de erros

O SGPilot grava um log completo no arquivo **`sgp_auto.log`** (na mesma pasta do executável). Em caso de erro, consulte esse arquivo para diagnóstico.

---

## Atalhos rápidos

| Ação | Como fazer |
|---|---|
| Conectar ao Chrome | Clique em "Conectar ao Chrome" no topo |
| Acionar uma automação | Pressione a hotkey (F1, F2, etc.) |
| Iniciar loop Papervines | Pressione F9 (ou a tecla configurada) |
| Parar loop Papervines | Pressione F9 novamente |
| Editar uma bind | Clique no botão da bind |
| Criar nova bind | Clique em "+ Nova Bind" |
| Ver configurações | Clique em "Configurações" |

---

*Desenvolvido por Vitor Glennon — SGPilot v5.1.0*
