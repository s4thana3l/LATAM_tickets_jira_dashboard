# Projeto Dashboard Tickets LATAM - Stanley 1913

## Visão geral

Este projeto automatiza a extração de tickets do Jira, organiza os dados em uma estrutura tabular e gera um dashboard visual em Excel.

A solução foi pensada para uso interno, com autenticação via API token disponível gratuitamente, export em Excel/CSV e apresentação resumida para gestão.

O fluxo atual do projeto inclui duas partes principais:

1. `jira_ticket_export.py` — coleta os dados do Jira e salva em arquivos locais
2. `jira_dashboard.py` — lê esses dados e monta abas resumidas com gráficos

---

## Objetivo do projeto

O objetivo principal é transformar a resposta JSON do Jira em informação útil e visualmente legível para:

- acompanhamento de tickets
- revisão de status
- acompanhamento por responsável
- relatórios mensais
- automação e integração com outras ferramentas

---

## Estrutura do projeto

- `jira_ticket_export.py` — extração dos dados do Jira
- `jira_dashboard.py` — dashboard em Excel
- `.env` — variáveis sensíveis do ambiente
- `.env.example` — exemplo de configuração
- `requirements.txt` — dependências do Python, importante instalá-las
- `README.md` — visão rápida do projeto
- `jira_tickets.xlsx` — arquivo com a base e o dashboard
- `jira_tickets.csv` — export de apoio em CSV

---

## Fluxo funcional do projeto

### 1) Leitura do ambiente

O primeiro passo é carregar informações sensíveis do arquivo `.env`.

Variáveis esperadas:

```env
JIRA_URL=https://seu-dominio.atlassian.net
JIRA_EMAIL=seu-email@empresa.com
JIRA_TOKEN=seu-api-token
JIRA_TARGET_TICKETS=500
```

Essas variáveis são lidas com `os.getenv()` e `python-dotenv`.

Se faltarem email ou token, o script interrompe a execução para evitar autenticação inválida.

### 2) Autenticação com o Jira

A autenticação usada é HTTP Basic com e-mail e API token.

```python
auth = (EMAIL, TOKEN)
```

Essa autenticação é enviada em cada request para a API do Jira.

### 3) Montagem do filtro JQL

A busca no Jira é feita com JQL, que define o intervalo de tempo e o conjunto de tickets a serem processados.

A lógica atual usa:

- projeto
- período do mês atual
- responsáveis específicos
- ordenação por data de criação

Os assignees são preservados como parte do filtro para manter o filtro direcionado a nós, que atendemos os tickets de LATAM.

### 4) Chamada à API

A rota usada é:

```python
https://{dominio}.atlassian.net/rest/api/3/search/jql
```

O payload enviado é algo como:

```python
payload = {
    "jql": jql,
    "fields": ["summary", "status", "assignee", "created", "resolution"],
    "maxResults": 100
}
```

Importante: o ambiente Jira atual pode rejeitar alguns padrões de paginação por offset, então a solução adotada é dividir a busca por janelas de data em vez de usar `startAt` puro.

### 5) Recebimento do JSON

A resposta da API retorna um JSON com chave `issues`, que contém a lista de tickets retornados.

A biblioteca `requests` faz a chamada HTTP e a resposta é convertida em dicionário com `.json()`.

### 6) Normalização dos dados

Os campos da API vêm aninhados e precisam ser extraídos.

Exemplo:

- `status` → objeto com `name`
- `assignee` → objeto com `displayName`
- `resolution` → objeto com `name`

O código transforma isso em uma estrutura simples, adequada para tabela e análise.

### 7) Criação do DataFrame

A estrutura final é transformada em um DataFrame do pandas usando as colunas:

```python
["key", "summary", "status", "assignee", "created", "resolution"]
```

Isso permite:

- exportação em Excel
- exportação em CSV
- agregação por status e responsável
- construção do dashboard

### 8) Export para Excel e CSV

O script cria ou atualiza o Excel e salva também uma cópia em CSV.

A planilha do mês atual fica em uma aba específica, por exemplo:

- `2026-08`

Isso ajuda a manter histórico por mês sem apagar os dados anteriores.

### 9) Dashboard no Excel

O dashboard lê o Excel gerado e monta abas de resumo:

- Resumo
- Status
- Assignee
- Gráficos

As abas são montadas em um único workbook para facilitar a leitura e a apresentação.

---

## Explicando cada parte do código

### `jira_ticket_export.py`

#### `load_jira_credentials()`

Lê as variáveis do `.env` e valida a presença do email e do token.

Se faltarem dados, o script para com uma mensagem clara.

#### `build_date_windows(start_date, end_date, chunk_days=7)`

Divide o período em janelas de datas.

Esse passo foi necessário porque o ambiente Jira atual se mostrou incompatível com certos usos de `startAt`, então a solução mais estável foi fazer a busca por blocos de tempo.

#### `build_jql(window_start, window_end, assignee_ids=None)`

Monta o JQL para cada janela de tempo.

Ela constrói uma query com:

- intervalo de criação
- projeto
- assignees
- ordenação por data de criação

#### `request_search(...)`

Executa a chamada HTTP para a API do Jira.

Responsabilidades:

- montar o payload JSON
- enviar a request
- tratar código HTTP
- manter a resposta em formato útil para o restante do fluxo

#### `normalize_issue(issue, assignee_ids)`

Recebe um item bruto da API e transforma em um dicionário limpo.

Ela também valida se o ticket pertence ao conjunto de assignees relevantes.

#### `collect_rows(...)`

É o bloco principal da extração.

Ela percorre cada janela de data, busca a resposta da API e acumula os tickets válidos em uma lista.

#### `apply_sheet_style(ws)`

Ajusta a aparência visual da aba Excel.

Responsabilidades:

- fonte em negrito para cabeçalho
- cor azul escura
- centralização do texto
- bordas finas
- largura ajustada das colunas
- congelamento do painel

#### `save_excel_with_month_sheet(df, excel_path, month_sheet)`

Salva a base do mês atual em uma aba específica do workbook.

Também evita duplicação de registros e mantém o histórico por período.

#### `main()`

Ponto de entrada do script.

Esse bloco orquestra:

- autenticação
- busca
- processamento
- exportação
- feedback visual no terminal

---

### `jira_dashboard.py`

#### `load_source_data()`

Lê o Excel do mês atual ou usa o CSV como fallback.

Essa função permite que o dashboard funcione mesmo quando o arquivo Excel ainda não foi gerado ou quando a aba do mês não existe.

#### `normalize_dataframe(raw_df)`

Padroniza os dados para uso no dashboard.

Ela garante que as colunas relevantes existam e que o DataFrame fique pronto para cálculos e gráficos.

#### `build_summary_tables(df)`

Cria os resumos necessários para gestão.

Ele gera:

- overview geral
- resumo por status
- resumo por responsável
- percentuais para facilitar leitura

#### `apply_dashboard_style(ws)`

Aplica a formatação visual das abas do dashboard.

Esse estilo mantém a apresentação mais profissional e mais fácil de ler.

#### `write_dashboard_sheets(df, workbook_path)`

Escreve as abas do dashboard no Excel existente.

Aba final:

- Resumo
- Status
- Assignee
- Gráficos

#### `main()`

Ponto de entrada do dashboard.

Ela chama:

- carregamento do arquivo fonte
- normalização dos dados
- escrita das abas de resumo
- atualização final do workbook

---

## Fluxo real em sequência

O projeto funciona assim, por ordem:

1. o usuário preenche o `.env`
2. executa `python jira_ticket_export.py`
3. o script conecta ao Jira
4. busca os tickets do mês atual
5. realiza a normalização dos dados
6. salva `jira_tickets.xlsx` e `jira_tickets.csv`
7. executa `python jira_dashboard.py`
8. o dashboard lê os dados exportados
9. monta as abas de resumo e os gráficos
10. apresenta o relatório pronto para leitura

---

## Vantagens da solução atual

- mantém credenciais fora do código
- separa responsabilidades em dois scripts
- facilita manutenção e evolução
- mantém visual mais profissional no Excel
- funciona bem para apresentação interna e relatórios gerenciais

---

## Melhorias possíveis no futuro

- adicionar tratativas de erro mais detalhadas
- criar logs do processo
- automatizar a execução mensal
- incluir dashboards mais elaborados por cliente, projeto ou área
- criar relatório em HTML ou em Power BI

---

## Conclusão

O projeto já está funcional para o uso principal: extrair tickets do Jira, exportar dados em planilha e gerar um dashboard de leitura rápida para gestão.

A arquitetura atual está organizada por responsabilidade, o que facilita continuar evoluindo sem quebrar o fluxo principal.

Em vez de um único arquivo, pode-se separar por:

- projeto
- status
- responsável
- data



## Resumo curto

A aplicação conecta no Jira, filtra tickets de interesse, transforma os dados em linhas e colunas e exporta para Excel/CSV. Isso cria um material pronto para automações e relatórios, principalmente quando combinado com Power Automate, Excel Online, SharePoint ou Power BI.
