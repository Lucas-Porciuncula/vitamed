# ==============================================================================
# 1. IMPORTS
# ==============================================================================
import pandas as pd
import pyodbc
import numpy as np
import os
import smtplib

from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dados import QUERY_CANCELADOS_FEIRA
from templates import HTML_ZIMBRA  

# ==============================================================================
# 2. CONFIGURAÇÕES GERAIS
# ==============================================================================
load_dotenv()

BINS = [0, 10, 20, 30, 40, 50, 60, 70, np.inf]
LABELS = ['0 a 9', '10 a 19', '20 a 29', '30 a 39', '40 a 49', '50 a 59', '60 a 69', '70+']
BINSMESES = [0, 7, 13, 25, np.inf]
LABELSMESES = ['1 a 6 meses', '6 a 12 meses', '1 a 2 anos', '2+']


# ==============================================================================
# 3. CONEXÃO COM BANCO
# ==============================================================================
def get_connection():
    """
    Cria conexão com SQL Server usando variáveis de ambiente (.env)
    """
    return pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('SERVER2')};"
        f"DATABASE={os.getenv('DB_NAME2')};"
        f"UID={os.getenv('DB_USER')};"
        f"PWD={os.getenv('DB_PASS')};"
    )


# ==============================================================================
# 4. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================
def processar_agrupamento(df, coluna_grupo, nome_metrica):
    """
    Agrupa, conta registros e adiciona linha de TOTAL GERAL
    """
    agrupado = (
        df
        .groupby(coluna_grupo, observed=False)
        .size()
        .reset_index(name=nome_metrica)
    )

    total = agrupado[nome_metrica].sum()
    linha_total = pd.DataFrame([[' TOTAL GERAL', total]], columns=agrupado.columns)
    agrupado = pd.concat([agrupado, linha_total], ignore_index=True)

    return agrupado


def processar_agrupamento_com_percentual(df, coluna_grupo, nome_metrica):
    """
    Agrupa, conta registros, calcula percentual e adiciona linha de TOTAL GERAL
    """
    agrupado = (
        df
        .groupby(coluna_grupo, observed=False)
        .size()
        .reset_index(name=nome_metrica)
    )
    
    total = agrupado[nome_metrica].sum()
    agrupado['%'] = (agrupado[nome_metrica] / total * 100).round(1).astype(str) + '%'
    
    linha_total = pd.DataFrame([[' TOTAL GERAL', total, '100.0%']], columns=agrupado.columns)
    agrupado = pd.concat([agrupado, linha_total], ignore_index=True)
    
    return agrupado


def calcular_top_vendedores(df, top_n=10):
    """
    Retorna os top N vendedores com mais cancelamentos
    """
    vendedores = (
        df
        .groupby('NomeVendedor')
        .size()
        .reset_index(name='Qtd')
        .sort_values('Qtd', ascending=False)
        .head(top_n)
    )
    
    total = vendedores['Qtd'].sum()
    vendedores['%'] = (vendedores['Qtd'] / df.shape[0] * 100).round(1).astype(str) + '%'
    
    return vendedores


def calcular_estatisticas_valor(df):
    """
    Calcula estatísticas sobre valores de mensalidade
    """
    df_com_valor = df[df['ValorMensalidade'].notna() & (df['ValorMensalidade'] > 0)]
    
    if len(df_com_valor) == 0:
        return None
    
    stats = pd.DataFrame({
        'Métrica': ['Valor Médio', 'Valor Mediano', 'Valor Mínimo', 'Valor Máximo', 'Receita Perdida (Mensal)'],
        'Valor': [
            f"R$ {df_com_valor['ValorMensalidade'].mean():.2f}",
            f"R$ {df_com_valor['ValorMensalidade'].median():.2f}",
            f"R$ {df_com_valor['ValorMensalidade'].min():.2f}",
            f"R$ {df_com_valor['ValorMensalidade'].max():.2f}",
            f"R$ {df_com_valor['ValorMensalidade'].sum():.2f}"
        ]
    })
    
    return stats


def analisar_planos(df):
    """
    Analisa distribuição por tipo de plano
    """
    planos = (
        df
        .groupby('DescricaoPlano')
        .size()
        .reset_index(name='Qtd')
        .sort_values('Qtd', ascending=False)
    )
    
    total = planos['Qtd'].sum()
    planos['%'] = (planos['Qtd'] / total * 100).round(1).astype(str) + '%'
    
    return planos

def analisar_estrelas(df):
    """
    Analisa distribuição por classificação de plano (estrelas)
    """
    estrelas = (
        df
        .groupby('BA1_VTCLAS')
        .size()
        .reset_index(name='Qtd')
        .sort_values('Qtd', ascending=False)
    )
    
    total = estrelas['Qtd'].sum()
    estrelas['%'] = (estrelas['Qtd'] / total * 100).round(1).astype(str) + '%'
    
    return estrelas

def analisar_forma_pagamento(df):
    """
    Analisa distribuição por forma de pagamento
    """
    pagamento = (
        df[df['TipoPagamento'].notna()]
        .groupby('TipoPagamento')
        .size()
        .reset_index(name='Qtd')
        .sort_values('Qtd', ascending=False)
    )
    
    if len(pagamento) > 0:
        total = pagamento['Qtd'].sum()
        pagamento['%'] = (pagamento['Qtd'] / total * 100).round(1).astype(str) + '%'
    
    return pagamento


# ==============================================================================
# 5. FUNÇÕES DE HTML (FRONT-END DO EMAIL)
# ==============================================================================
def df_to_html(df, cor_header='#1a3c5e', cor_barra='#2e6da4', com_barra=True):
    """
    Converte DataFrame para HTML estilizado (email-safe).
    Se a coluna '%' existir e com_barra=True, adiciona coluna visual de barra.

    Args:
        df          : DataFrame a converter
        cor_header  : Cor do cabeçalho ('#1a3c5e' cancelamentos, '#198754' reativações)
        cor_barra   : Cor da barra de progresso
        com_barra   : Se True e coluna '%' existir, adiciona coluna de barras
    """

    TH = (
        f'background:{cor_header};color:#ffffff;padding:10px 12px;'
        f'text-align:left;font-size:11px;text-transform:uppercase;'
        f'letter-spacing:0.5px;font-weight:600;border:none;'
    )
    TD      = 'padding:9px 12px;color:#2d3e50;border-bottom:1px solid #f0f4f8;border:none;border-bottom:1px solid #f0f4f8;'
    TD_NUM  = TD + 'text-align:right;'
    TD_PCT  = TD + 'text-align:right;color:#7a8fa6;'
    TD_BAR  = TD + 'width:120px;'

    tem_pct = '%' in df.columns and com_barra

    # ── Cabeçalho ──────────────────────────────────────────────────────────────
    thead = '<thead><tr>'
    for col in df.columns:
        alinhamento = 'text-align:right;' if col in ('Qtd', 'Qtd Cancelados', 'Qtd Resgatada') else ''
        thead += f'<th style="{TH}{alinhamento}">{col}</th>'
    if tem_pct:
        thead += f'<th style="{TH}width:120px;"></th>'  # coluna vazia p/ barra
    thead += '</tr></thead>'

    # ── Corpo ───────────────────────────────────────────────────────────────────
    tbody = '<tbody>'
    for _, row in df.iterrows():
        eh_total = str(row.iloc[0]).strip().upper().startswith('TOTAL')
        peso = 'font-weight:700;' if eh_total else ''

        tbody += '<tr>'
        for col in df.columns:
            val = row[col]
            if col == '%':
                tbody += f'<td style="{TD_PCT}{peso}">{val}</td>'
            elif col in ('Qtd', 'Qtd Cancelados', 'Qtd Resgatada'):
                tbody += f'<td style="{TD_NUM}{peso}">{val}</td>'
            else:
                tbody += f'<td style="{TD}{peso}">{val}</td>'

        # Barra visual
        if tem_pct:
            pct_val = row['%']
            try:
                largura = float(str(pct_val).replace('%', '').strip())
                # limita a 100 para não quebrar layout
                largura = min(largura, 100)
            except ValueError:
                largura = 0

            if eh_total or largura == 0:
                tbody += f'<td style="{TD_BAR}"></td>'
            else:
                tbody += (
                    f'<td style="{TD_BAR}">'
                    f'<div style="background:#e8edf2;border-radius:4px;height:6px;width:100%;">'
                    f'<div style="background:{cor_barra};border-radius:4px;height:6px;width:{largura}%;"></div>'
                    f'</div>'
                    f'</td>'
                )
        tbody += '</tr>'
    tbody += '</tbody>'

    # ── Tabela final ────────────────────────────────────────────────────────────
    return (
        f'<table style="width:100%;border-collapse:collapse;font-size:13px;">'
        f'{thead}{tbody}'
        f'</table>'
    )


def criar_card_kpi(titulo, valor, cor_fundo='#e9f1ff', cor_texto='#0b5ed7'):
    """
    Cria um card KPI estilizado
    """
    return f"""
    <div style="background:{cor_fundo};padding:15px;border-radius:6px;text-align:center;margin:10px 0;">
        <h4 style="margin:0;color:{cor_texto};font-size:14px;">{titulo}</h4>
        <p style="font-size:24px;margin:5px 0 0;font-weight:bold;color:#000;">
            {valor}
        </p>
    </div>
    """


# ==============================================================================
# 6. ENVIO DE EMAIL
# ==============================================================================
def enviar_email_report(dfs_dict, metricas_dict, destinatarios, com_copia=None, com_copia_oculta=None,  cor_header='#1a3c5e'):
    """
    Monta HTML final e envia o email para múltiplos destinatários
    
    Args:
        dfs_dict: Dicionário com os DataFrames a serem enviados
        metricas_dict: Dicionário com as métricas (KPIs)
        destinatarios: String ou lista de emails principais (To)
        com_copia: String ou lista de emails em cópia (Cc) - opcional
        com_copia_oculta: String ou lista de emails em cópia oculta (Bcc) - opcional
    """
    meu_email = "laporciuncula@vitalmed.com.br"
    minha_senha = os.getenv("SENHA_EMAIL")

    smtp_server = "smtp.emailzimbraonline.com"
    porta = 587

    # -----------------------
    # Processa destinatários
    # -----------------------
    def processar_emails(emails):
        """Converte string ou lista em lista de emails"""
        if isinstance(emails, str):
            return [email.strip() for email in emails.split(',')]
        elif isinstance(emails, list):
            return [email.strip() for email in emails]
        return []
    
    lista_to = processar_emails(destinatarios)
    lista_cc = processar_emails(com_copia) if com_copia else []
    lista_bcc = processar_emails(com_copia_oculta) if com_copia_oculta else []
    
    # Lista completa para envio (To + Cc + Bcc)
    todos_destinatarios = lista_to + lista_cc + lista_bcc

    # -----------------------
    # Montagem dos KPIs
    # -----------------------
    kpis_html = '<table width="100%" cellpadding="5" cellspacing="0"><tr>'
    
    col_width = 100 // len(metricas_dict)
    for titulo, valor in metricas_dict.items():
        kpis_html += f'<td width="{col_width}%">{criar_card_kpi(titulo, valor)}</td>'
    
    kpis_html += '</tr></table>'

    # -----------------------
    # Montagem das Tabelas
    # -----------------------
    tabelas_html = ""

    for titulo, df in dfs_dict.items():
        if df is not None and not df.empty:
            tabelas_html += f"""
            <h3 style="margin-top:30px;color:#0b5ed7;border-bottom:2px solid #0b5ed7;padding-bottom:8px;">
                {titulo}
            </h3>
            {df_to_html(df)}
            """

    # -----------------------
    # Injeta no Template
    # -----------------------
    html_body = (
        HTML_ZIMBRA
        .replace("{{DATA}}", pd.Timestamp.now().strftime('%d/%m/%Y'))
        .replace("{{TOTAL_CANCELADOS}}", str(metricas_dict.get('Total de Cancelamentos', 0)))
        .replace("{{KPIS}}", kpis_html)
        .replace("{{TABELAS}}", tabelas_html)
    )

    # -----------------------
    # Criação do Email
    # -----------------------
    msg = MIMEMultipart()
    msg['From'] = meu_email
    msg['To'] = ', '.join(lista_to)
    if lista_cc:
        msg['Cc'] = ', '.join(lista_cc)
    # Bcc não vai no header (por isso é oculto)
    msg['Subject'] = f"📊 Relatório de Cancelamentos FSA VitalMed - {pd.Timestamp.now().strftime('%d/%m/%Y')}"

    msg.attach(MIMEText(html_body, 'html'))

    # -----------------------
    # Envio
    # -----------------------
    try:
        server = smtplib.SMTP(smtp_server, porta)
        server.starttls()
        server.login(meu_email, minha_senha)
        server.send_message(msg, to_addrs=todos_destinatarios)
        server.quit()
        print(f"✅ E-mail enviado com sucesso para {len(todos_destinatarios)} destinatário(s)!")
        print(f"   📧 Para: {', '.join(lista_to)}")
        if lista_cc:
            print(f"   📄 Cc: {', '.join(lista_cc)}")
        if lista_bcc:
            print(f"   🔒 Bcc: {len(lista_bcc)} destinatário(s) oculto(s)")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")


# ==============================================================================
# 7. PIPELINE DE EXECUÇÃO
# ==============================================================================
if __name__ == "__main__":
    print("=" * 50)
    print("🔄 Iniciando processamento de cancelamentos...")
    print("=" * 50)
    
    # Conexão e extração
    conn = get_connection()
    print("✅ Conexão estabelecida com o banco de dados")
    
    df_cancelados = pd.read_sql(QUERY_CANCELADOS_FEIRA, conn)
    conn.close()
    print(f"✅ {len(df_cancelados)} registros extraídos")

    # -----------------------
    # Transformações de Dados
    # -----------------------
    print("🔄 Processando dados...")
    
    # Limpeza e padronização
    df_cancelados['NomeVendedor'] = df_cancelados['NomeVendedor'].str.strip()
    df_cancelados['Idade'] = df_cancelados['Idade'].clip(lower=0)
    
    # Garantir tipos corretos
    df_cancelados['TempoContratacao_Anos'] = pd.to_numeric(
        df_cancelados['TempoContratacao_Anos'], 
        errors='coerce'
    ).fillna(0)
    
    df_cancelados['ValorMensalidade'] = pd.to_numeric(
        df_cancelados['ValorMensalidade'], 
        errors='coerce'
    )

    # Criar faixas
    df_cancelados['Faixa Idade'] = pd.cut(
        df_cancelados['Idade'], 
        bins=BINS, 
        labels=LABELS, 
        right=False
    )
    
    df_cancelados['Faixa Tempo de Casa'] = pd.cut(
        df_cancelados['TempoContratacao_Anos'], 
        bins=BINSMESES, 
        labels=LABELSMESES, 
        right=False
    )

    # -----------------------
    # Cálculo de Métricas
    # -----------------------
    total_cancelados = df_cancelados.shape[0]
    idade_media = df_cancelados['Idade'].mean()
    tempo_medio = df_cancelados['TempoContratacao_Anos'].mean() / 12
    
    # Receita perdida
    valor_total_perdido = df_cancelados['ValorMensalidade'].sum()
    
    metricas = {
        'Total de Cancelamentos': f"{total_cancelados:,}".replace(',', '.'),
        'Idade Média': f"{idade_media:.0f} anos",
        'Tempo Médio': f"{tempo_medio:.0f} anos",
        'Receita Perdida': f"R$ {valor_total_perdido:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    }

    # -----------------------
    # Agrupamentos e Análises
    # -----------------------
    print("📊 Gerando relatórios...")
    
    agrup_idade = processar_agrupamento_com_percentual(
        df_cancelados, 'Faixa Idade', 'Qtd Cancelados'
    )
    
    agrup_tempo = processar_agrupamento_com_percentual(
        df_cancelados, 'Faixa Tempo de Casa', 'Qtd Cancelados'
    )
    
    agrup_motivo = processar_agrupamento_com_percentual(
        df_cancelados, 'MotivoBloqueio', 'Qtd Cancelados'
    )
    
    top_vendedores = calcular_top_vendedores(df_cancelados, top_n=10)
    
    stats_valor = calcular_estatisticas_valor(df_cancelados)
    
    analise_planos = analisar_planos(df_cancelados)
    
    analise_pagamento = analisar_forma_pagamento(df_cancelados)

    # -----------------------
    # Montagem dos Relatórios
    # -----------------------
    relatorios = {
        '📊 Cancelados por Faixa Etária': agrup_idade,
        '⏰ Cancelados por Tempo de Casa': agrup_tempo,
        '⭐ Cancelados por Classificação de Plano': analisar_estrelas(df_cancelados),
        '🚫 Cancelados por Motivo de Bloqueio': agrup_motivo,
        '👥 Top 10 Vendedores com Mais Cancelamentos': top_vendedores,
        '📋 Cancelados por Tipo de Plano': analise_planos,
    }
    
    # Adicionar análises opcionais se houver dados
    if stats_valor is not None:
        relatorios['💰 Estatísticas de Valor'] = stats_valor
    
    if not analise_pagamento.empty:
        relatorios['💳 Cancelados por Forma de Pagamento'] = analise_pagamento

    # -----------------------
    # Envio do Email
    # -----------------------
    print("📧 Enviando relatório por e-mail...")
    
    # ========================================
    # OPÇÕES DE DESTINATÁRIOS
    # ========================================
    
    # OPÇÃO 1: Um único destinatário
    # enviar_email_report(
    #     dfs_dict=relatorios,
    #     metricas_dict=metricas,
    #     destinatarios="adrianasilva@vitalmed.com.br"
    # )
    
    # OPÇÃO 2: Lista de destinatários (To)
    # enviar_email_report(
    #     dfs_dict=relatorios,
    #     metricas_dict=metricas,
    #     destinatarios=[
    #         "adrianasilva@vitalmed.com.br",
    #         "gerencia@vitalmed.com.br",
    #         "diretoria@vitalmed.com.br"
    #     ]
    # )
    
    # OPÇÃO 3: String separada por vírgulas (To)
    # enviar_email_report(
    #     dfs_dict=relatorios,
    #     metricas_dict=metricas,
    #     destinatarios="adrianasilva@vitalmed.com.br, gerencia@vitalmed.com.br"
    # )
    
    # OPÇÃO 4: Com destinatários em cópia (Cc) e cópia oculta (Bcc)
    enviar_email_report(
        dfs_dict=relatorios,
        metricas_dict=metricas,
        destinatarios="adrianasilva@vitalmed.com.br",
        com_copia=[
            "ccoliveira@vitalmed.com.br",
            "unmelo@vitalmed.com.br",
            "ssmarques@vitalmed.com.br",
            "easantos@vitalmed.com.br",
            "marisangela.jesus@vitalmed.com.br",
            "laporciuncula@vitalmed.com.br",
            "lcsobral@vitalmed.com.br",
            "ebsouza@vitalmed.com.br",
            "jclacerda@vitalmed.com.br",
            "admvendas@vitalmed.com.br",
            "gestaocrc@vitalmed.com.br",
            "fasampaio@vitalmed.com.br",
            "jbalmeida@vitalmed.com.br",
            "mesilva@vitalmed.com.br",
            "pssilva@vitalmed.com.br"
        ],
        com_copia_oculta=[],
        cor_header='#1a3c5e'
    )
    
    print("=" * 50)
    print("✅ Processo finalizado com sucesso!")
    print("=" * 50)