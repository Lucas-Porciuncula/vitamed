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

from dados import QUERY, QUERY_REATIVADOS
from templates import HTML_ZIMBRA, HTML_REATIVADOS  
from CanceladosCLAUDE import *


# ==============================================================================
# 2. funções auxiliares 
# ==============================================================================
def calcular_estatisticas_valor_reativados(df):
    """
    Calcula estatísticas sobre valores de mensalidade
    """
    df_com_valor = df[df['ValorTotal'].notna() & (df['ValorTotal'] > 0)]
    
    if len(df_com_valor) == 0:
        return None
    
    stats = pd.DataFrame({        'Métrica': ['Valor Médio', 'Valor Mediano', 'Valor Mínimo', 'Valor Máximo', 'Receita Resgatada (Mensal)'],
        'Valor': [
            f"R$ {df_com_valor['ValorTotal'].mean():.2f}",
            f"R$ {df_com_valor['ValorTotal'].median():.2f}",
            f"R$ {df_com_valor['ValorTotal'].min():.2f}",
            f"R$ {df_com_valor['ValorTotal'].max():.2f}",
            f"R$ {df_com_valor['ValorTotal'].sum():.2f}"
        ]
    })
def enviar_email_report(dfs_dict, metricas_dict, destinatarios, com_copia=None, com_copia_oculta=None):
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
        HTML_REATIVADOS
        .replace("{{DATA}}", pd.Timestamp.now().strftime('%d/%m/%Y'))
        .replace("{{TOTAL_REATIVADOS}}", str(metricas_dict.get('Total de Reativados', 0)))
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
    msg['Subject'] = f"📊 Relatório de Reativações VitalMed - {pd.Timestamp.now().strftime('%d/%m/%Y')}"

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
# 2. Conectar 
# ==============================================================================
if __name__ == "__main__":
    print("=" * 50)
    print("🔄 Iniciando processamento de reativações...")
    print("=" * 50)
    
    # Conexão e extração
    conn = get_connection()
    print("✅ Conexão estabelecida com o banco de dados")

    df_reativados = pd.read_sql_query(QUERY_REATIVADOS, conn)
    conn.close()
    print(f"✅ {len(df_reativados)} registros extraídos")

    # -----------------------
    # Transformações de Dados
    # -----------------------
    print("🔄 Processando dados...")

    # Limpeza e padronização
    df_reativados['NomeVendedor'] = df_reativados['NomeVendedor'].str.strip()
    df_reativados['Idade'] = df_reativados['Idade'].clip(lower=0)

    # Garantir tipos corretos
    df_reativados['AnosDesdeInclusao'] = pd.to_numeric(
        df_reativados['AnosDesdeInclusao'], 
        errors='coerce'
    ).fillna(0)

    df_reativados['ValorTotal'] = pd.to_numeric(
        df_reativados['ValorTotal'], 
        errors='coerce'
    )

    # Criar faixas
    df_reativados['Faixa Idade'] = pd.cut(
        df_reativados['Idade'], 
        bins=BINS, 
        labels=LABELS, 
        right=False
    )

    df_reativados['Faixa Tempo de Casa'] = pd.cut(
        df_reativados['AnosDesdeInclusao'], 
        bins=BINSMESES, 
        labels=LABELSMESES, 
        right=False
    )

    # -----------------------
    # Cálculo de Métricas
    # -----------------------
    total_reativados = df_reativados.shape[0]
    idade_media = df_reativados['Idade'].mean()
    tempo_medio = df_reativados['AnosDesdeInclusao'].mean() / 12
    
    # Receita Resgatada
    valor_resgatado = df_reativados['ValorTotal'].sum()
    
    metricas = {
        'Total de Reativados': f"{total_reativados:,}".replace(',', '.'),
        'Idade Média': f"{idade_media:.0f} anos",
        'Tempo Médio': f"{tempo_medio:.0f} anos",
        'Receita Resgatada': f"R$ {valor_resgatado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    }
    # -----------------------
    # Agrupamentos e Análises
    # -----------------------
    print("📊 Gerando relatórios...")
    
    agrup_idade = processar_agrupamento_com_percentual(
        df_reativados, 'Faixa Idade', 'Qtd Resgatada'
    )
    
    agrup_tempo = processar_agrupamento_com_percentual(
        df_reativados, 'Faixa Tempo de Casa', 'Qtd Resgatada'
    )
    
    top_vendedores = calcular_top_vendedores(df_reativados, top_n=10)
    
    stats_valor = calcular_estatisticas_valor_reativados(df_reativados)
    
    analise_planos = analisar_planos(df_reativados)
    
    analise_pagamento = analisar_forma_pagamento(df_reativados)

    # -----------------------
    # Montagem dos Relatórios
    # -----------------------
    relatorios = {
        '📊 Reativados por Faixa Etária': agrup_idade,
        '⏰ Reativados por Tempo de Casa Desde Inclusão': agrup_tempo,
        '👥 Top 10 Vendedores com Mais Reativações': top_vendedores,
        '📋 Reativados por Tipo de Plano': analise_planos,
    }
    
    # Adicionar análises opcionais se houver dados
    if stats_valor is not None:
        relatorios['💰 Estatísticas de Valor'] = stats_valor
    
    if not analise_pagamento.empty:
        relatorios['💳 Reativados por Forma de Pagamento'] = analise_pagamento

    # -----------------------
    # Envio do Email
    # -----------------------
    print("📧 Enviando relatório por e-mail...")
    enviar_email_report(
        dfs_dict=relatorios,
        metricas_dict=metricas,
        destinatarios="laporciuncula@vitalmed.com.br",  # Destinatário principal
        com_copia = [
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
    "gestaocrc@vitalmed.com.br"
],  # Cc
        com_copia_oculta=[""]  # Bcc (não aparece para outros)
    )
    
    print("=" * 50)
    print("✅ Processo finalizado com sucesso!")
    print("=" * 50)

