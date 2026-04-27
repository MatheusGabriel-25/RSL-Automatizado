# 📚 busca_rsl.py — Automação de Revisão Sistemática da Literatura (RSL)

Script Python para automatizar a coleta, enriquecimento e auditoria de metadados bibliográficos de uma biblioteca do **Zotero**, gerando uma planilha Excel estruturada e pronta para o fluxo PRISMA de uma Revisão Sistemática da Literatura (RSL).

---

## 📋 Índice

1. [O que o script faz](#-o-que-o-script-faz)
2. [Pré-requisitos](#-pré-requisitos)
3. [Passo 1 — Instalar o Zotero](#-passo-1--instalar-o-zotero)
4. [Passo 2 — Instalar o Zotero Connector (extensão do navegador)](#-passo-2--instalar-o-zotero-connector-extensão-do-navegador)
5. [Passo 3 — Criar a API Key do Zotero](#-passo-3--criar-a-api-key-do-zotero)
6. [Passo 4 — Obter seu User ID do Zotero](#-passo-4--obter-seu-user-id-do-zotero)
7. [Passo 5 — Configurar o script](#-passo-5--configurar-o-script)
8. [Passo 6 — Instalar o Python](#-passo-6--instalar-o-python)
9. [Passo 7 — Executar o script](#-passo-7--executar-o-script)
10. [Como o script funciona internamente](#-como-o-script-funciona-internamente)
11. [Estrutura da planilha gerada](#-estrutura-da-planilha-gerada)
12. [Perguntas frequentes e problemas comuns](#-perguntas-frequentes-e-problemas-comuns)

---

## 🚀 O que o script faz

O `busca_rsl.py` é uma ferramenta de automação que realiza as seguintes etapas automaticamente:

1. **Conecta na sua biblioteca do Zotero** via API e baixa todos os artigos cadastrados.
2. **Enriquece os metadados** de cada artigo em até 4 camadas (Crossref → Semantic Scholar → OpenAlex → Web Scraping / PDF direto), preenchendo automaticamente os campos que estiverem vazios: resumo, DOI, autores e ano.
3. **Realiza uma Auditoria 360°** (opcional) que valida DOIs, cruza autores, confere anos e verifica se os links estão ativos.
4. **Gera uma planilha Excel** (`.xlsx`) já formatada com todas as abas do fluxo PRISMA, pronta para preencher manualmente as etapas de triagem.

---

## 📦 Pré-requisitos

Antes de executar o script, você precisa ter:

| Requisito | Descrição |
|---|---|
| **Zotero Desktop** | Gerenciador de referências bibliográficas (gratuito) |
| **Zotero Connector** | Extensão do navegador para salvar artigos no Zotero |
| **Conta no zotero.org** | Necessária para sincronizar a biblioteca na nuvem e usar a API |
| **API Key do Zotero** | Chave de acesso que autoriza o script a ler sua biblioteca |
| **User ID do Zotero** | Número de identificação único da sua conta |
| **Python 3.8+** | Linguagem de programação em que o script foi escrito |

> **Atenção:** O script instala automaticamente todas as bibliotecas Python necessárias na primeira execução. Você **não** precisa instalar nada manualmente além do Python.

---

## 🦊 Passo 1 — Instalar o Zotero

1. Acesse o site oficial: **[https://www.zotero.org/download/](https://www.zotero.org/download/)**
2. Clique em **"Download"** para o seu sistema operacional (Windows, macOS ou Linux).
3. Execute o instalador e siga as instruções na tela.
4. Ao abrir o Zotero pela primeira vez, **crie uma conta gratuita** em [https://www.zotero.org/user/register](https://www.zotero.org/user/register) e faça login dentro do programa (`Editar > Preferências > Sincronização`).

> A conta é essencial porque o script acessa sua biblioteca pela **nuvem (API web)**, não pelo arquivo local.

---

## 🔌 Passo 2 — Instalar o Zotero Connector (extensão do navegador)

O **Zotero Connector** é uma extensão que adiciona um botão ao seu navegador. Com ele, você pode salvar artigos de qualquer página da web (Google Scholar, IEEE, Springer, etc.) diretamente na sua biblioteca do Zotero com um único clique.

### Como instalar:

1. Ainda na página [https://www.zotero.org/download/](https://www.zotero.org/download/), clique em **"Install Connector"** para o seu navegador (Chrome, Firefox ou Edge).
2. Você será redirecionado para a loja de extensões do navegador. Clique em **"Adicionar ao navegador"** (ou "Add to Chrome" / "Add to Firefox").
3. Depois de instalado, um ícone do Zotero aparecerá na barra de ferramentas do navegador.

### Como usar para salvar artigos:

1. Acesse a página de um artigo (ex: no Google Scholar).
2. Clique no ícone do Zotero Connector na barra do navegador.
3. Selecione a pasta da sua biblioteca onde quer salvar e clique em **"Done"**.
4. O artigo aparecerá automaticamente na sua biblioteca do Zotero.

> **Dica de RSL:** Para pesquisa sistemática, crie uma **coleção** (pasta) no Zotero específica para seu tema. O script busca **todos os itens de nível superior** da biblioteca, então organize bem antes de rodar.

### Como adicionar um Ranking manual (campo Extra):

O script lê o campo **"Extra"** de cada item do Zotero para definir a ordem (`RANK`) na planilha. Para definir manualmente a posição de um artigo:

1. Clique no artigo no Zotero.
2. No painel direito, procure o campo **"Extra"**.
3. Digite **apenas o número** de classificação desejado no início do campo. Exemplo: `1` ou `01`.

Se nenhum artigo tiver número no campo Extra, o script numera automaticamente na ordem em que os itens aparecem.

---

## 🔑 Passo 3 — Criar a API Key do Zotero

A **API Key** é uma senha de acesso que você gera no site do Zotero. Ela permite que o script leia sua biblioteca sem precisar do seu login e senha.

### Como criar:

1. Acesse: **[https://www.zotero.org/settings/keys](https://www.zotero.org/settings/keys)**
2. Faça login na sua conta do Zotero, caso necessário.
3. Clique em **"Create new private key"**.
4. Dê um nome descritivo para a chave (ex: `Script RSL`).
5. Em **"Permissions"**, marque as seguintes opções:
   - ✅ **Allow library access** (acesso de leitura à biblioteca pessoal)
   - ✅ **Read Only** (somente leitura — o script não precisa escrever)
6. Clique em **"Save Key"**.
7. **Copie a chave gerada** (ela só aparece uma vez!). É uma sequência de letras e números, similar a: `AbCdEfGh1IjKlMnOp2`

> **Guarde bem essa chave.** Você vai precisar dela no Passo 5.

---

## 🪪 Passo 4 — Obter seu User ID do Zotero

O **User ID** é o número único que identifica sua conta. O script precisa dele para saber de qual biblioteca baixar os dados.

### Como encontrar:

1. Acesse: **[https://www.zotero.org/settings/keys](https://www.zotero.org/settings/keys)**
2. No topo da página, você verá uma mensagem como:
   > *"Your userID for use in API calls is **XXXXXXXX**"*
3. Esse número `XXXXXXXX` é o seu **User ID**. Anote-o.

---

## ⚙️ Passo 5 — Configurar o script

Abra o arquivo `busca_rsl.py` em um editor de texto (Bloco de Notas, VS Code, etc.) e localize o bloco de configuração:

```python
# ==========================================
# Configurações da automação da Revisão Sistemática da Literatura
# ==========================================
ID_ZOTERO = 'SEU_USER_ID_AQUI'
CHAVE_API = 'SUA_API_KEY_AQUI'
SEU_EMAIL = 'seu.email@exemplo.com'
```

Substitua os valores:

| Campo | O que colocar |
|---|---|
| `ID_ZOTERO` | Seu **User ID** do Zotero (obtido no Passo 4) |
| `CHAVE_API` | Sua **API Key** do Zotero (obtida no Passo 3) |
| `SEU_EMAIL` | Seu **e-mail real** (usado para se identificar nas APIs acadêmicas e evitar bloqueios) |

### Exemplo de configuração preenchida:

```python
ID_ZOTERO = '12345678'
CHAVE_API = 'AbCdEfGh1IjKlMnOp2'
SEU_EMAIL = 'fulano@universidade.edu.br'
```

> **Por que o e-mail?** APIs acadêmicas como o Crossref usam o conceito de *"Polite Pool"*: quem se identifica com um e-mail real recebe prioridade e raramente é bloqueado. Não precisa ser uma senha — é só para identificação.

### Configuração de dígitos do Ranking (opcional):

```python
DIGITOS_RANK = 2
```

Esse número define quantos dígitos o campo `RANK` terá na planilha. Com `2`, o ranking sai como `01`, `02`, `03`... Com `3`, sairia `001`, `002`, etc.

---

## 🐍 Passo 6 — Instalar o Python

Se ainda não tiver o Python instalado:

1. Acesse: **[https://www.python.org/downloads/](https://www.python.org/downloads/)**
2. Baixe a versão mais recente do **Python 3** (3.8 ou superior).
3. Execute o instalador.
4. **IMPORTANTE:** Marque a opção **"Add Python to PATH"** antes de clicar em "Install Now".

### Verificar se o Python está instalado corretamente:

Abra o terminal (Prompt de Comando no Windows ou Terminal no macOS/Linux) e digite:

```bash
python --version
```

Deve aparecer algo como `Python 3.11.x`. Se funcionar, você está pronto.

---

## ▶️ Passo 7 — Executar o script

1. Abra o terminal na pasta onde o `busca_rsl.py` está salvo.
   - **Windows:** Clique com o botão direito dentro da pasta e escolha "Abrir no Terminal" ou "Abrir janela de comando aqui".
2. Execute o script com o comando:

```bash
python busca_rsl.py
```

3. O script irá:
   - Verificar e instalar automaticamente as bibliotecas necessárias (somente na primeira vez).
   - Pedir que você **digite seu nome** — ele será usado no nome do arquivo Excel gerado.
   - Conectar ao Zotero e processar cada artigo da biblioteca, mostrando o progresso no terminal.
   - Ao final, perguntar se você deseja rodar a **Auditoria Profunda** (veja abaixo).
   - Salvar a planilha `.xlsx` na mesma pasta do script.

### Exemplo de saída no terminal:

```
======================================================
 🚀 Fazendo a Revisão Sistemática da Literatura All_Dataset
======================================================
Digite seu nome para o relatório: João Silva

📖 [1/30] Artigo: Deep Learning for Smart Home Automation...
+---------------+--------------------+------------------------------+
| CAMPO         | STATUS             | FONTE                        |
+---------------+--------------------+------------------------------+
| DOI           | ✅ Encontrado      | Zotero                       |
| Resumo        | ✅ Encontrado      | Crossref                     |
| Título        | ✅ Encontrado      | Zotero                       |
| Autores       | ✅ Encontrado      | Zotero                       |
| Ano           | ✅ Encontrado      | Zotero                       |
| Link          | ✅ Encontrado      | Gerado via DOI               |
+---------------+--------------------+------------------------------+
```

---

## 🔬 Como o script funciona internamente

### Instalação automática de dependências

Na primeira execução, o script detecta quais bibliotecas Python estão faltando e as instala automaticamente via `pip`, sem que você precise fazer nada manualmente.

As bibliotecas utilizadas são:

| Biblioteca | Para que serve |
|---|---|
| `pyzotero` | Conectar e baixar dados da biblioteca do Zotero via API |
| `pandas` | Organizar os dados em formato de tabela |
| `requests` | Fazer chamadas HTTP para as APIs e páginas web |
| `urllib3` | Comunicação HTTP em nível mais baixo |
| `pdfplumber` | Ler e extrair texto de arquivos PDF |
| `beautifulsoup4` | Extrair dados de páginas HTML (web scraping) |
| `openpyxl` | Criar e formatar a planilha Excel de saída |

---

### Motor de Enriquecimento Hierárquico (4 camadas)

Para cada artigo, o script tenta preencher os campos vazios passando por até 4 fontes em cascata. Ele para na primeira que funcionar para cada campo:

```
┌─────────────────────────────────────────────────────────┐
│ Camada 0: Zotero                                        │
│ → Lê os dados já cadastrados: título, resumo, DOI,      │
│   autores, ano e link.                                  │
└────────────────────┬────────────────────────────────────┘
                     │ (se algum campo estiver vazio)
                     ▼
┌─────────────────────────────────────────────────────────┐
│ Camada 1: Crossref API                                  │
│ → Busca pelo título. Aceita somente se a similaridade   │
│   do título retornado for > 75%.                        │
│ → Preenche: DOI, Resumo, Ano, Autores.                  │
└────────────────────┬────────────────────────────────────┘
                     │ (se ainda faltar)
                     ▼
┌─────────────────────────────────────────────────────────┐
│ Camada 2: Semantic Scholar API                          │
│ → Busca pelo título (limitado a 100 caracteres).        │
│ → Preenche: Resumo, Ano, Autores.                       │
└────────────────────┬────────────────────────────────────┘
                     │ (se ainda faltar)
                     ▼
┌─────────────────────────────────────────────────────────┐
│ Camada 3: OpenAlex API                                  │
│ → Busca pelo DOI (mais preciso). Decodifica o           │
│   "abstract_inverted_index" próprio do OpenAlex.        │
│ → Preenche: Resumo, Autores.                            │
└────────────────────┬────────────────────────────────────┘
                     │ (se ainda faltar)
                     ▼
┌─────────────────────────────────────────────────────────┐
│ Camada 4: Web Scraping / Leitor de PDF                  │
│ → Acessa o link do artigo ou doi.org/{DOI} diretamente. │
│ → Se a resposta for um PDF: usa pdfplumber + regex      │
│   para isolar o resumo (Abstract/Resumo/Summary).       │
│ → Se for HTML: usa BeautifulSoup para ler meta tags     │
│   (citation_title, dc.description, og:description etc.) │
│ → Suporte especial para: IEEE, ProQuest, Figshare.      │
│ → Preenche: Resumo, Título, Autores, Ano, DOI.          │
└─────────────────────────────────────────────────────────┘
```

> **Anti-bloqueio:** O script espera 1 segundo entre cada artigo e se identifica com seu e-mail nas APIs, seguindo as boas práticas de uso (*Polite Pool*).

---

### Auditoria Profunda 360° (opcional)

Ao final da coleta, o script pergunta se você quer rodar a auditoria. Ela realiza as seguintes verificações em cada artigo:

| Verificação | O que faz |
|---|---|
| **DOI** | Confirma o DOI via Crossref, Semantic Scholar, OpenAlex e Google Scholar. Se o artigo não tiver DOI, tenta resgatar automaticamente. |
| **Título vs. DOI** | Cruza o título cadastrado com o título oficial da referência para detectar DOIs trocados. |
| **Autores** | Compara os autores do Zotero com os autores oficiais da API para detectar erros. |
| **Ano** | Compara o ano cadastrado com o ano oficial registrado nas bases de dados. |
| **Link** | Verifica se o link está ativo (código HTTP). Distingue entre links quebrados e links protegidos por anti-bot. Quando disponível, sugere o link de PDF em Acesso Aberto. |
| **Resumo** | Avalia se o resumo tem tamanho razoável, se não contém textos de erro (CAPTCHA, login required) e se contém palavras relevantes do título. |

Os resultados da auditoria aparecem na coluna **"AVISOS DE AUDITORIA"** da planilha.

---

## 📊 Estrutura da planilha gerada

O arquivo Excel gerado (`RSL_Dashboard_NomeDoUsuario.xlsx`) contém **6 abas**, seguindo o fluxo metodológico PRISMA:

| Aba | Descrição |
|---|---|
| **All_Dataset** | Todos os artigos coletados com os metadados enriquecidos e os avisos de auditoria. |
| **Remoção da Literatura Cinza** | Aba vazia para o pesquisador filtrar manualmente teses, relatórios e literatura não indexada. |
| **Remoção de estudos duplicados** | Aba vazia para remover manualmente artigos duplicados após a revisão inicial. |
| **1 - Busca (Screening) - Título** | Triagem por título, com colunas extras: "Status (Incluído/Excluído)" e "Motivo Exclusão". |
| **2 - Leitura do Texto completo** | Triagem por leitura integral, com as mesmas colunas de status e motivo. |
| **Estudos Finais** | Artigos aprovados para análise aprofundada, com colunas extras para tipo de estudo, método, resultados principais, etc. |

### Colunas principais (aba All_Dataset):

| Coluna | Descrição |
|---|---|
| `RANK` | Classificação/ordem do artigo (lida do campo Extra do Zotero) |
| `DOI` | Identificador digital do objeto (Digital Object Identifier) |
| `AUTORES` | Lista de autores separados por `;` |
| `TÍTULO` | Título completo do artigo |
| `RESUMO` | Abstract/resumo completo |
| `ANO DE PUBLICAÇÃO` | Ano da publicação |
| `BASE DE DADOS` | Base de origem (fixado como "Google Scholar") |
| `LINK` | URL de acesso ao artigo |
| `AVISOS DE AUDITORIA` | Resultado detalhado da auditoria 360° |

---

## ❓ Perguntas frequentes e problemas comuns

### O script trava ou não conecta no Zotero
- Verifique se o `ID_ZOTERO` e a `CHAVE_API` foram preenchidos corretamente no script.
- Confirme que sua conta do Zotero está sincronizada e que os artigos aparecem em [https://www.zotero.org/](https://www.zotero.org/).

### Muitos artigos ficaram sem resumo
- É normal para artigos de editoras pagas (Elsevier, Springer, IEEE) que bloqueiam o acesso sem assinatura.
- Esses artigos serão marcados na planilha com `N/A` no campo Resumo.
- A coluna de auditoria indicará se há um PDF em Acesso Aberto disponível como alternativa.

### O script foi bloqueado pelo Google Scholar
- O Google Scholar bloqueia robôs periodicamente. O script já espera 3 segundos antes de acessar o Google, mas bloqueios podem ocorrer em lotes grandes.
- Se ocorrer bloqueio, o resultado na coluna de auditoria será: `⚠️ DOI Não Validado (Google Scholar Bloqueou IP)`.
- Tente rodar novamente depois de algumas horas.

### O arquivo Excel já existe e o script não sobrescreve
- O script detecta automaticamente que o arquivo existe e salva com um número incrementado no nome (`RSL_Dashboard_Nome_1.xlsx`, `_2.xlsx`, etc.).

### Como adicionar mais artigos e rodar novamente?
- Adicione os novos artigos no Zotero normalmente pelo Connector.
- Sincronize a biblioteca (`Arquivo > Sincronizar`).
- Execute o script novamente. Ele baixará todos os artigos do zero e gerará um novo arquivo.

### Posso usar o script com uma biblioteca de grupo do Zotero?
- O script está configurado para bibliotecas pessoais (`'user'`). Para adaptar para grupos do Zotero, é necessário modificar a linha de configuração da conexão de `'user'` para `'group'` e ajustar o ID para o ID do grupo.

---

## 🛡️ Boas práticas de segurança

- **Nunca compartilhe** sua `CHAVE_API` do Zotero publicamente (não faça upload do script com a chave preenchida em repositórios públicos como GitHub).
- Você pode revogar e gerar uma nova API Key a qualquer momento em [https://www.zotero.org/settings/keys](https://www.zotero.org/settings/keys).

---

## 📎 Resumo rápido de instalação

```
1. Baixar e instalar o Zotero → zotero.org/download
2. Criar conta em → zotero.org/user/register
3. Instalar o Zotero Connector no navegador → zotero.org/download
4. Obter o User ID em → zotero.org/settings/keys
5. Criar a API Key em → zotero.org/settings/keys
6. Preencher ID_ZOTERO, CHAVE_API e SEU_EMAIL no início do script
7. Instalar o Python 3.8+ → python.org/downloads
8. Executar: python busca_rsl.py
```
