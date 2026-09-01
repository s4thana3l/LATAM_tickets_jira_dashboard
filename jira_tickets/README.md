# Jira Ticket Export + Dashboard - Stanley 1913

Este projeto conecta ao Jira pela API REST dele, coleta os tickets do mês vigente, salva os dados em Excel/CSV e monta um dashboard básico para uma leitura mais gráfica.

O objetivo é lidar com mutitos dados brutos vindos diretamente do Jira de uma forma simplificada e de fácil leitura. Sem a necessidade de extrairmos estes dados manualmente. 

## Visão geral

O fluxo principal funciona em duas etapas:

1. Extrair os tickets do Jira
2. Gerar um dashboard em Excel com resumo por status, responsável e gráficos

Em resumo, o script:
- autentica no Jira com e-mail corporativo + token de API do Jira
- executa um filtro JQL, semelhante a uma query SQL padrão
- lê os campos principais do ticket, como status, título e responsável
- transforma o JSON em tabela
- salva a base em Excel e CSV
- monta abas com resumo visual para 

## Quando usar este projeto

Esse projeto é útil quando você precisa:
- extrair dados do Jira sem depender de licenças premium
- acompanhar volume, status e responsáveis
- apresentar informações de operação de forma mais objetiva
- manter um histórico mensal por aba/planilha

## Estrutura do projeto

- `jira_ticket_export.py` — script principal de extração
- `jira_dashboard.py` — script que lê o Excel e cria o dashboard
- `.env` — arquivo local com credenciais e dados sensíveis
- `.env.example` — modelo para configurar o ambiente
- `requirements.txt` — dependências do Python
- `README.md` — documentação resumida do projeto
- `DOCUMENTACAO_FLUXO.md` — explicação detalhada do fluxo, código e lógica
- `jira_tickets.xlsx` — arquivo principal, em Excel com base + dashboard
- `jira_tickets.csv` — export em CSV para apoio e backup

## Funcionalidades principais

- autenticação com e-mail e API token do Jira
- busca por filtro JQL com data e projeto
- leitura de campos como chave, resumo, status, responsável e resolução
- normalização de dados aninhados da API
- export em Excel e CSV
- criação de aba por mês para manter histórico
- dashboard com abas:
  - Resumo
  - Status
  - Assignee
  - Gráficos

## Requisitos

- Python 3.9+
- Pacotes principais:
  - `requests`
  - `pandas`
  - `openpyxl`
  - `python-dotenv`
  - `urllib3`

## Instalação

Na pasta do projeto, execute:

```bash
pip install -r requirements.txt
```

## Configuração do ambiente

Crie um arquivo `.env` na raiz do projeto com este conteúdo:

```env  
JIRA_URL=https://seu-dominio.atlassian.net
JIRA_EMAIL=seu-email@empresa.com
JIRA_TOKEN=seu-api-token
JIRA_TARGET_TICKETS=500
```

### O que cada variável faz

- `JIRA_URL`: URL base do Jira
- `JIRA_EMAIL`: e-mail do usuário com acesso ao Jira
- `JIRA_TOKEN`: token de API gerado pelo site de opções de sua conta Jira / Atlassian
- `JIRA_TARGET_TICKETS`: limite de registros que o script tenta coletar

## Como executar

### 1) Gerar os dados do Jira

```bash
python jira_ticket_export.py
```

Esse passo:
- conecta ao Jira
- executa o filtro de busca
- recebe o JSON da API
- limpa e organiza os campos
- salva a base em Excel e CSV

### 2) Gerar o dashboard

```bash
python jira_dashboard.py
```

Esse passo:
- lê o arquivo exportado
- organiza os dados para dashboard
- cria/atualiza as abas de resumo e gráficos
- deixa o Excel com visual pronto para apresentação

## Como o fluxo funciona

O script principal segue esta ordem:

1. lê o `.env`
2. autentica no Jira via HTTP Basic
3. monta o filtro JQL com período e projeto
4. chama a API do Jira em `/rest/api/3/search/jql`
5. recebe a resposta em JSON
6. normaliza os campos aninhados
7. transforma em DataFrame
8. exporta para Excel e CSV

Essa separação ajuda a manter o processo organizado em responsabilidades bem definidas:
- extração de dados
- transformação/limpeza
- exportação
- apresentação visual

## Como funciona o dashboard

O dashboard lê os dados exportados e monta um resumo mais fácil de acompanhar visualmente.

Ele cria abas como:
- `Resumo` — visão geral do volume e dos indicadores
- `Status` — contagem por situação dos tickets
- `Assignee` — visão por responsável
- `Gráficos` — representação visual para leitura rápida

Essa parte foi pensada para facilitar a leitura de quem vai apresentar o relatório sem precisar analisar a planilha inteira linha por linha.

## Boas práticas

- feche o arquivo Excel antes de rodar os scripts novamente
- mantenha as credenciais no `.env`, nunca no código
- use o projeto como base para relatórios mensais ou de acompanhamento
- consulte a documentação detalhada quando precisar entender cada regra do fluxo

## Documentação detalhada

Para ver a explicação completa do processo e do código, consulte:

- [DOCUMENTACAO_FLUXO.md](./DOCUMENTACAO_FLUXO.md)

## Observações finais

Este projeto foi construído para ser simples de usar, fácil de manter e útil para a nossa equipe de Operações. Não foi feito para ser executado de forma externa, apenas interna e localmente.
