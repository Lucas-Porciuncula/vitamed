# 🏥 VitalMed Analytics — Repositório de Automações & BI

> **5 meses de trabalho solo** como único profissional de dados da Grupo VitalMed.
> Pipelines de cancelamento, reativação, churn, automação de relatórios e extração de dados — tudo em produção.

---

## 📁 Estrutura do Repositório

```
.
├── CANCELADOS/
│   ├── CanceladosCLAUDE.py       # Pipeline principal de cancelamentos (Salvador)
│   ├── canceladosFeira.py        # Pipeline de cancelamentos (Feira de Santana)
│   ├── reativados.py             # Pipeline de reativações (Salvador)
│   ├── reativadosFeira.py        # Pipeline de reativações (Feira de Santana)
│   ├── churn.py                  # GUI + queries de churn reversível por período
│   ├── churnrelatorio.py         # Relatório semestral/mensal de churn por tempo de casa
│   ├── app.py                    # GUI Tkinter para geração de relatórios Excel
│   ├── dados.py                  # Queries SQL centralizadas (views + queries diretas)
│   └── templates.py              # Templates HTML dos e-mails (Zimbra-safe)
├── metasPorEquipe/
│   ├── aphTradicional.py         # Sistema de acompanhamento de metas APH por equipe
│   └── whatsapp2.py              # Robô de envio de relatório de metas via WhatsApp Web
├── ligar_gateway/
│   └── gateway.py                # Automação pyautogui para login no On-premises Data Gateway
├── zimbra/
│   ├── main.py                   # CLI para extração de CEPs de e-mails via IMAP
│   ├── varrer_dados.py           # Classe ZimbraCEPExtractor (IMAP + parsing + export)
│   └── outro_teste.py            # Utilitário de listagem de pastas IMAP
├── cancelados.bat                # Orquestrador: roda os 4 pipelines em sequência
├── email.bat                     # Launcher do robô de WhatsApp
└── .gitignore
```

---

## 🚀 Módulos

### `CANCELADOS/` — Pipeline de Churn & Reativação

O core do repositório. Roda diariamente via `cancelados.bat` e envia e-mails HTML automáticos para ~15 destinatários em Salvador e Feira de Santana.

**O que cada script faz:**

| Script | Fonte de dados | Saída |
|--------|---------------|-------|
| `CanceladosCLAUDE.py` | `VW_CANCELADOS_MESPLS` (SQL Server) | E-mail HTML com 8 análises de cancelamentos (SSA) |
| `canceladosFeira.py` | `VW_CANCELADOS_FSA_MESPLS` | E-mail HTML para Feira de Santana |
| `reativados.py` | `VW_REATIVADOS_MESPLS` | E-mail HTML de reativações (SSA) |
| `reativadosFeira.py` | `VW_REATIVADOS_FSA_MESPLS` | E-mail HTML de reativações (FSA) |

**Análises geradas automaticamente em cada relatório:**
- Cancelados/Reativados por faixa etária
- Cancelados/Reativados por tempo de casa
- Classificação de plano (estrelas)
- Motivo de bloqueio
- Top 10 vendedores
- Tipo de plano
- Estatísticas de valor (receita perdida / resgatada)
- Forma de pagamento

**KPIs no cabeçalho do e-mail:**
- Total de cancelamentos/reativações
- Idade média
- Tempo médio de casa
- Receita perdida / resgatada (R$)

---

### `CANCELADOS/churn.py` — Ferramenta de Análise de Churn Reversível

GUI Tkinter para geração de relatório Excel de churn por período customizado.

- Conecta via SQLAlchemy + pyodbc
- 3 abas: **Tempo de Casa**, **Faixa Etária**, **Classificação de Plano**
- Filtro por motivos reversíveis (dificuldade financeira, desinteresse, etc.)
- CTE com `ROLLUP` — total automático no SQL, sem pandas

```
Entrada: data início + data fim (formato AAAAMMDD)
Saída:   Churn_Completo_YYYYMMDD_YYYYMMDD.xlsx
```

---

### `CANCELADOS/churnrelatorio.py` — Série Histórica de Churn

Gera análise semestral ou mensal automaticamente a partir de uma lista de períodos.

- `MODO = "SEMESTRE"` → agrupa meses em S1/S2 automaticamente
- `MODO = "PERIODO"` → uma aba por mês
- Cada aba = resultado da query de tempo de casa para aquele intervalo

---

### `CANCELADOS/app.py` — Gerador de Relatório Excel com GUI

Interface minimalista (dark mode) para o time operacional gerar relatórios CSV/Excel sem precisar abrir o SQL.

- Toggle cidade: **Salvador** / **Feira de Santana**
- Filtro por período
- Saída: `cancelados_{cidade}_{data_inicio}_{data_fim}.xlsx`

---

### `metasPorEquipe/` — Acompanhamento de Metas APH

Sistema de acompanhamento diário de metas de vidas APH por equipe de vendas.

**`aphTradicional.py`**
- Carrega metas do Excel (`BASE_VENDEDORES.xlsx` em rede interna)
- Busca propostas aprovadas no SQL Server (`VENDAS_ONLINE_COMERCIAL`)
- Normaliza período de apuração (ciclo 26→25)
- Gera DataFrame com: `EQUIPE | Vidas APH | Meta | % da Meta`

**`whatsapp2.py`**
- Usa Selenium + ChromeDriver para enviar o relatório via WhatsApp Web
- Perfil Chrome persistente (sem QR code a cada execução)
- Formatação em negrito com `*texto*` nativa do WhatsApp
- Lista de grupos configurável no topo do arquivo

---

### `ligar_gateway/gateway.py` — Automação do On-premises Data Gateway

Script pyautogui que liga e autentica o Gateway de dados local automaticamente.

- Abre o atalho do Gateway via `os.startfile`
- Localiza botões por screenshot com `confidence=0.9`
- Preenche e-mail e clica em avançar automaticamente
- Usado para garantir que o Power BI refresh funcione sem intervenção manual

---

### `zimbra/` — Extrator de CEPs via IMAP

Ferramenta CLI para extrair CEPs de e-mails do servidor Zimbra.

- Conecta via IMAP SSL (porta 993)
- Busca e-mails com a palavra "CEP" no assunto ou corpo
- Extrai CEPs em múltiplos formatos: `12345-678`, `12345678`, `CEP: 12345678`
- Exporta para Excel ou CSV com timestamp

```
Opções de busca:
  1. Últimos 100 e-mails
  2. Últimos 500 e-mails
  3. Últimos 1000 e-mails
  4. Últimos 6 meses
  5. Período customizado
  6. TODOS os e-mails
```

---

## ⚙️ Setup

### Pré-requisitos

```
Python 3.13+
ODBC Driver 17 for SQL Server
Chrome instalado (para whatsapp2.py)
```

### Instalação

```bash
pip install pandas pyodbc numpy python-dotenv openpyxl sqlalchemy \
            selenium webdriver-manager pyperclip beautifulsoup4
```

### Variáveis de ambiente (`.env`)

```env
# SQL Server — banco principal
SERVER2=<servidor>
DB_NAME2=<database>
DB_USER=<usuario>
DB_PASS=<senha>

# SQL Server — banco de propostas/metas
SERVER=<servidor>
DB_NAME=<database>
DB_PASS6=<senha>

# E-mail Zimbra
SENHA_EMAIL=<senha>
```

---

## ▶️ Execução

### Rodar todos os pipelines de uma vez

```bat
cancelados.bat
```

Executa em sequência:
1. Cancelados SSA
2. Reativados SSA
3. Cancelados FSA
4. Reativados FSA

### Rodar individualmente

```bash
python CANCELADOS/CanceladosCLAUDE.py
python CANCELADOS/reativados.py
python CANCELADOS/canceladosFeira.py
python CANCELADOS/reativadosFeira.py
```

### Ferramentas com GUI

```bash
python CANCELADOS/app.py        # Relatório Excel
python CANCELADOS/churn.py      # Análise de churn reversível
```

### Relatório de metas no WhatsApp

```bash
email.bat
# ou
python metasPorEquipe/whatsapp2.py
```

### Extrator de CEPs

```bash
python zimbra/main.py
```

---

## 🏗️ Arquitetura de Dados

```
SQL Server (Protheus)
    │
    ├── VW_CANCELADOS_MESPLS          → CanceladosCLAUDE.py
    ├── VW_CANCELADOS_FSA_MESPLS      → canceladosFeira.py
    ├── VW_REATIVADOS_MESPLS          → reativados.py
    ├── VW_REATIVADOS_FSA_MESPLS      → reativadosFeira.py
    ├── VENDAS_ONLINE_COMERCIAL        → aphTradicional.py
    └── BA1010, BCA010, BI3010, ...   → QUERY_PURA_PARA_RELATORIO (queries diretas)

Rede interna (\\192.168.1.26)
    └── BASE_VENDEDORES.xlsx           → aphTradicional.py (metas + equipes)

IMAP Zimbra
    └── imap.emailzimbraonline.com     → zimbra/varrer_dados.py
```

---

## 📧 Fluxo dos E-mails

```
SQL Server
    └─► pd.read_sql()
            └─► transformações pandas (faixas, agrupamentos, KPIs)
                    └─► df_to_html() — tabelas HTML com barras de progresso
                            └─► templates.HTML_ZIMBRA / HTML_REATIVADOS
                                    └─► smtplib STARTTLS → smtp.emailzimbraonline.com:587
```

Os templates são **Zimbra-safe**: apenas `<table>` inline, sem CSS externo, sem `<div>` para layout principal.

---

## 📊 Contexto de Negócio

Esses scripts nasceram da necessidade de **visibilidade diária sobre churn** numa operadora de saúde com dois centros de operação (Salvador e Feira de Santana). Antes, os relatórios eram manuais e chegavam com dias de atraso.

Com este pipeline automatizado:
- **Tempo de geração**: de horas para ~2 minutos
- **Frequência**: diária, via agendador de tarefas do Windows
- **Alcance**: 15 destinatários por e-mail, incluindo coordenação e diretoria
- **Receita monitorada**: perdida (cancelamentos) e resgatada (reativações) em tempo real

---

## 👨‍💻 Autor

**Lucas Amorim** — Analista de Dados  
Grupo VitalMed · Salvador, Bahia  
`laporciuncula@vitalmed.com.br`

---

*Repositório interno. Dados sensíveis protegidos via `.env` e `.gitignore`.*
