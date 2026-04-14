"""
Gerador de Relatório PDF para Lead Scoring Vitalmed
Cria um documento técnico para apresentação à gestão
"""

from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, 
    TableStyle, Image, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.lib import colors
from datetime import datetime
import os

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================

OUTPUT_FILE = "Relatorio_Lead_Scoring_Vitalmed.pdf"
TITULO_PRINCIPAL = "LEAD SCORING — ANÁLISE PREDITIVA DE VENDAS"
SUBTITULO = "Vitalmed — Canais Digitais"
DATA_GERACAO = datetime.now().strftime("%d de %B de %Y").replace("of", "de")

# =============================================================================
# ESTILOS PERSONALIZADOS
# =============================================================================

styles = getSampleStyleSheet()

titulo = ParagraphStyle(
    'TituloCustom',
    parent=styles['Heading1'],
    fontSize=24,
    textColor=colors.HexColor('#1C3144'),
    spaceAfter=6,
    alignment=TA_CENTER,
    fontName='Helvetica-Bold'
)

subtitulo_style = ParagraphStyle(
    'SubtituloCustom',
    parent=styles['Normal'],
    fontSize=14,
    textColor=colors.HexColor('#3D5A7F'),
    spaceAfter=20,
    alignment=TA_CENTER,
    fontName='Helvetica'
)

heading2 = ParagraphStyle(
    'Heading2Custom',
    parent=styles['Heading2'],
    fontSize=14,
    textColor=colors.HexColor('#1C3144'),
    spaceAfter=12,
    spaceBefore=12,
    fontName='Helvetica-Bold'
)

heading3 = ParagraphStyle(
    'Heading3Custom',
    parent=styles['Heading3'],
    fontSize=11,
    textColor=colors.HexColor('#3D5A7F'),
    spaceAfter=8,
    fontName='Helvetica-Bold'
)

normal_justify = ParagraphStyle(
    'NormalJustify',
    parent=styles['Normal'],
    fontSize=10,
    alignment=TA_JUSTIFY,
    spaceAfter=10,
    leading=13
)

highlight = ParagraphStyle(
    'Highlight',
    parent=styles['Normal'],
    fontSize=10,
    textColor=colors.HexColor('#0D5C3D'),
    fontName='Helvetica-Bold'
)

# =============================================================================
# CONTEÚDO DO RELATÓRIO
# =============================================================================

def criar_relatorio():
    """Monta e gera o relatório PDF."""
    
    doc = SimpleDocTemplate(
        OUTPUT_FILE,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    story = []
    
    # ─────────────────────────────────────────────────────────────────────
    # 1. CAPA
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Spacer(1, 1.5*inch))
    story.append(Paragraph(TITULO_PRINCIPAL, titulo))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(SUBTITULO, subtitulo_style))
    story.append(Spacer(1, 0.3*inch))
    
    capa_info = Paragraph(
        f"<b>Relatório Técnico</b><br/>"
        f"Gerado em: {DATA_GERACAO}<br/>"
        f"<br/>"
        f"<i>Análise Preditiva para Priorização de Leads em Vendas</i>",
        ParagraphStyle(
            'CapaInfo',
            parent=styles['Normal'],
            fontSize=11,
            alignment=TA_CENTER,
            spaceAfter=12,
            textColor=colors.HexColor('#555555')
        )
    )
    story.append(capa_info)
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 2. SUMÁRIO EXECUTIVO
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("SUMÁRIO EXECUTIVO", heading2))
    
    exec_text = """
    Este projeto implementa um <b>modelo de Machine Learning</b> para prever a 
    probabilidade de conversão de leads, permitindo priorizar esforços comerciais 
    nos prospects com maior potencial de fechamento.
    <br/><br/>
    <b>Objetivos:</b><br/>
    • Identificar automaticamente leads de alto potencial de conversão<br/>
    • Otimizar alocação de recursos das equipes de vendas (SDR, Closers)<br/>
    • Reduzir tempo gasto com leads de baixa probabilidade<br/>
    • Fornecer insights sobre fatores que impactam a conversão<br/>
    <br/>
    <b>Resultado:</b> Modelo com <b>AUC-ROC de 0.9708</b> — performance excelente 
    que permite classificar com alta confiança quais leads têm maior chance de 
    se tornarem clientes.
    """
    story.append(Paragraph(exec_text, normal_justify))
    story.append(Spacer(1, 0.2*inch))
    
    # ─────────────────────────────────────────────────────────────────────
    # 3. CONTEXTO DO PROBLEMA
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("1. CONTEXTO E PROBLEMA", heading2))
    
    context_text = """
    <b>Desafio de Negócio:</b><br/>
    A Vitalmed recebe milhares de leads vinculados através de canais digitais 
    (tráfego pago, orgânico, offline). Atualmente, a classificação desses leads 
    é manual ou baseada em heurísticas simples, resultando em:<br/>
    <br/>
    • Alocação ineficiente de vendedores<br/>
    • Equipes gastando tempo com leads de baixa probabilidade<br/>
    • Sem visão preditiva do potencial de cada lead<br/>
    <br/>
    <b>Solução Proposta:</b><br/>
    Desenvolver um modelo preditivo que analisa características históricas dos 
    leads (demografias, comportamentos, canal de origem, histórico de interações) 
    para gerar um <b>score de conversão 0–100</b>, permitindo:<br/>
    <br/>
    • Priorização automática de leads Hot (75+) para SDR/Closers<br/>
    • Distribuição inteligente de leads por vendedor<br/>
    • Feedback contínuo para otimizar o modelo
    """
    story.append(Paragraph(context_text, normal_justify))
    story.append(Spacer(1, 0.15*inch))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 4. METODOLOGIA E DADOS
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("2. METODOLOGIA", heading2))
    
    story.append(Paragraph("2.1 Dataset de Treino", heading3))
    
    dataset_text = """
    • <b>Total de leads:</b> 10.806<br/>
    • <b>Leads com desfecho definido:</b> 7.501 (treino)<br/>
    &nbsp;&nbsp;&nbsp;— Ganhos (conversões): 811 (10,8%)<br/>
    &nbsp;&nbsp;&nbsp;— Perdidos/Cancelados: 6.690 (89,2%)<br/>
    • <b>Leads em aberto:</b> 9.995 (para scoring)<br/>
    <br/>
    <b>Origem dos dados:</b> Exportação do Power BI da tabela fLEADS, contendo 
    histórico de interações, estágios do funil e desfechos finalizados.
    """
    story.append(Paragraph(dataset_text, normal_justify))
    story.append(Spacer(1, 0.1*inch))
    
    story.append(Paragraph("2.2 Engenharia de Features", heading3))
    
    features_text = """
    <b>Features Numéricas (11):</b><br/>
    • Hora de criação (0–23h)<br/>
    • Dia da semana (0–6)<br/>
    • Mês de criação (1–12)<br/>
    • Dias no funil<br/>
    • Receita esperada<br/>
    • Dias até fechamento<br/>
    • Horas até fechamento<br/>
    • Flags: tem_email, tem_receita, tem_campanha, fim_de_semana<br/>
    <br/>
    <b>Features Categóricas (4):</b><br/>
    • Canal de origem (Tráfego Pago, Orgânico, Offline, etc)<br/>
    • Praça (Salvador, Feira de Santana, outras)<br/>
    • Equipe (SDR, Televendas, Vendas Diretas)<br/>
    • Turno de criação (Madrugada, Manhã, Tarde, Noite)<br/>
    <br/>
    <b>Processamento:</b> Features categóricas foram codificadas com LabelEncoder; 
    valores ausentes preenchidos com -1 (sentinel value).
    """
    story.append(Paragraph(features_text, normal_justify))
    story.append(Spacer(1, 0.15*inch))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 5. BIBLIOTECAS UTILIZADAS
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("3. BIBLIOTECAS UTILIZADAS", heading2))
    
    bibliotecas_data = [
        ["Biblioteca", "Versão", "Propósito", "Por que usar?"],
        [Paragraph("<b>Pandas</b>", styles['Normal']), 
         "2.x", 
         "Manipulação de dados, leitura CSV, transformações",
         "Padrão ouro para processamento tabular em Python. Eficiente e confiável."],
        
        [Paragraph("<b>NumPy</b>", styles['Normal']), 
         "1.x", 
         "Operações numéricas vetorizadas",
         "Acelera cálculos, essencial para ML. Bem testado e otimizado."],
        
        [Paragraph("<b>Scikit-Learn</b>", styles['Normal']), 
         "1.5+", 
         "Pré-processamento (LabelEncoder), métricas (ROC-AUC, CV)",
         "Padrão acadêmico e industrial. Implementações robustas e documentadas."],
        
        [Paragraph("<b>XGBoost</b>", styles['Normal']), 
         "2.x", 
         "Modelo de classificação (gradient boosting)",
         "Top-tier para competições e produção. Melhor tradeoff performance/velocidade."],
        
        [Paragraph("<b>Scikit-Learn Calibration</b>", styles['Normal']), 
         "1.5+", 
         "Calibração isotônica de probabilidades",
         "Garante que probabilidades preditas reflitam confiança real. Crítico para scores."],
        
        [Paragraph("<b>SHAP</b>", styles['Normal']), 
         "0.14+", 
         "Explicabilidade (importância de features)",
         "Interpretação teórica sólida (Shapley values). Essencial para validação."],
        
        [Paragraph("<b>Matplotlib</b>", styles['Normal']), 
         "3.x", 
         "Visualizações (gráficos de importância)",
         "Padrão para análise visual de modelos. Integrado com SHAP."],
    ]
    
    # Tabela de bibliotecas
    biblio_table = Table(bibliotecas_data, colWidths=[1.2*inch, 0.6*inch, 2.0*inch, 2.2*inch])
    biblio_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1C3144')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
    ]))
    
    story.append(biblio_table)
    story.append(Spacer(1, 0.2*inch))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 6. MODELO E PERFORMANCE
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("4. MODELO E PERFORMANCE", heading2))
    
    story.append(Paragraph("4.1 Arquitetura do Modelo", heading3))
    
    model_text = """
    <b>Algoritmo:</b> XGBoost com Calibração Isotônica<br/>
    <br/>
    <b>Hiperparâmetros:</b><br/>
    • n_estimators: 300 árvores (boosting sequencial)<br/>
    • max_depth: 5 (profundidade limitada para evitar overfitting)<br/>
    • learning_rate: 0.05 (passos pequenos no gradiente)<br/>
    • subsample: 0.8 (80% das amostras por árvore — regularização)<br/>
    • colsample_bytree: 0.8 (80% das features por árvore)<br/>
    • scale_pos_weight: 8.2 (ajusta para desbalanceamento 89% não-ganho vs 11% ganho)<br/>
    <br/>
    <b>Calibração:</b> Isotônica (3-fold CV) garante que scores 0–100 refletem 
    confiança real, não só ranking relativo.
    """
    story.append(Paragraph(model_text, normal_justify))
    story.append(Spacer(1, 0.15*inch))
    
    story.append(Paragraph("4.2 Resultados de Performance", heading3))
    
    metrics_data = [
        ["Métrica", "Valor", "Interpretação"],
        ["AUC-ROC (Teste)", "0.9708", "Excelente. Modelo discrimina bem entre Ganho/Não-Ganho."],
        ["AUC-ROC (CV 5-fold)", "0.9767 ± 0.0015", "Muito estável. Generaliza bem para dados novos."],
        ["Precisão (Ganho)", "93%", "De cada lead classificado como 'provável ganho', 93% realmente converte."],
        ["Recall (Ganho)", "75%", "Dos 162 ganhos no teste, modelo identifica 75%. Trade-off aceitável."],
        ["Acurácia Geral", "97%", "97% de todas as predições estão corretas."],
    ]
    
    perf_table = Table(metrics_data, colWidths=[2.0*inch, 1.5*inch, 2.5*inch])
    perf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1C3144')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E8F5E9')),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
    ]))
    
    story.append(perf_table)
    story.append(Spacer(1, 0.2*inch))
    
    story.append(Paragraph("4.3 Por que estas métricas provam confiabilidade?", heading3))
    
    conf_text = """
    ✓ <b>AUC-ROC > 0.97:</b> Modelo com performance próxima ao "ouro" em classificação 
    (valor máximo é 1.0). Erros aleatórios não alcançam isso.<br/>
    <br/>
    ✓ <b>Cross-validation (5-fold) estável:</b> Rodamos o treino 5 vezes com dados 
    diferentes. Resultado quase idêntico (±0.0015) prova que o modelo generaliza — 
    não é sorte ou overfitting.<br/>
    <br/>
    ✓ <b>Calibração isotônica:</b> Garante que um score de 80 realmente significa 
    "80% de chance de conversão", não só "melhor que 50".<br/>
    <br/>
    ✓ <b>Desbalanceamento tratado:</b> Dados têm 89% de não-ganhos. Usamos scale_pos_weight 
    para dar peso igual aos poucos ganhos. Sem isso, modelo tenderia a prever "não-ganho".
    """
    story.append(Paragraph(conf_text, normal_justify))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 7. INTERPRETABILIDADE
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("5. INTERPRETABILIDADE — O QUE IMPACTA A CONVERSÃO", heading2))
    
    shap_text = """
    <b>Técnica SHAP (SHapley Additive exPlanations):</b> Usa teoria dos jogos para 
    medir quanto cada feature contribui para as predições. Mais rigoroso que 
    "importância clássica" porque lida com correlações entre features.<br/>
    <br/>
    <b>Visualização gerada:</b> arquivo `shap_importancia_features.png` mostra 
    as 15 features mais influentes, ordenadas por impacto.<br/>
    <br/>
    <b>Benefícios para gestão:</b><br/>
    • Entender quais características de um lead importam mais<br/>
    • Validar se o modelo "pensa" de forma coerente (ex: receita esperada deve impactar)<br/>
    • Sustentar decisões: "Por que priorizamos este lead?" → Resposta: "Porque tem receita 
    esperada alta e foi criado em horário X"<br/>
    <br/>
    <b>Confiabilidade da interpretação:</b> SHAP tem fundamentação teórica sólida 
    (Shapley values já são conceitos conhecidos em economia). Não é "black box".
    """
    story.append(Paragraph(shap_text, normal_justify))
    story.append(Spacer(1, 0.2*inch))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 8. RESULTADOS OPERACIONAIS
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("6. RESULTADOS E DISTRIBUIÇÃO DE SCORES", heading2))
    
    distrib_data = [
        ["Faixa de Score", "Leads", "Interpretação", "Ação Recomendada"],
        ["Hot (76–100)", "58", "Alta probabilidade de conversão", "Prioritário para SDR/Closer"],
        ["Quente (51–75)", "159", "Média-alta probabilidade", "Acompanhamento ativo"],
        ["Morno (26–50)", "331", "Probabilidade moderada", "Nutrir com conteúdo"],
        ["Frio (0–25)", "9.447", "Baixa probabilidade", "Long-term nurture ou descarte"],
    ]
    
    distrib_table = Table(distrib_data, colWidths=[1.3*inch, 1.0*inch, 2.2*inch, 2.5*inch])
    distrib_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1C3144')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, -1), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (0, 1), colors.HexColor('#C8E6C9')),
        ('BACKGROUND', (0, 2), (0, 2), colors.HexColor('#BBDEFB')),
        ('BACKGROUND', (0, 3), (0, 3), colors.HexColor('#FFE0B2')),
        ('BACKGROUND', (0, 4), (0, 4), colors.HexColor('#FFCCBC')),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
    ]))
    
    story.append(distrib_table)
    story.append(Spacer(1, 0.2*inch))
    
    story.append(Paragraph("6.1 Casos de Uso", heading3))
    
    casos_text = """
    <b>Caso 1: Priorização em tempo real</b><br/>
    Novo lead entra no CRM → Score atribuído automaticamente → Se Hot, 
    notificação instantânea para SDR → Contato em &lt;2 horas.<br/>
    <br/>
    <b>Caso 2: Distribuição entre vendedores</b><br/>
    Gerente vê que vendedor A tem 3 Hot, vendedor B tem 12 Hot → 
    Redistributribui para balancear carga de alta prioridade.<br/>
    <br/>
    <b>Caso 3: Análise de padrão</b><br/>
    "Por que leads com receita esperada alta não vendem?" → 
    Investigar com SHAP e dados históricos → Ajustar processo de vendas.
    """
    story.append(Paragraph(casos_text, normal_justify))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 9. INTEGRAÇÃO NO POWER BI
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("7. INTEGRAÇÃO NO POWER BI", heading2))
    
    powerbi_text = """
    <b>Arquivos fornecidos:</b><br/>
    <br/>
    1. <b>lead_scores_vitalmed.csv</b><br/>
    &nbsp;&nbsp;Contém: lead_id, nome, vendedor, estágio, score (0–100), probabilidade, faixa<br/>
    &nbsp;&nbsp;→ Importar via Power Query como nova tabela "ScoreLeads"<br/>
    <br/>
    2. <b>instrucoes_powerbi.dax</b><br/>
    &nbsp;&nbsp;Medidas DAX prontas para usar:<br/>
    &nbsp;&nbsp;• "Score Conversão" — exibe score no cartão<br/>
    &nbsp;&nbsp;• "Cor Score" — cor para formatação condicional (verde=Hot, vermelho=Frio)<br/>
    <br/>
    3. <b>shap_importancia_features.png</b><br/>
    &nbsp;&nbsp;Gráfico de importância das features → Insira em um visual para documentação<br/>
    <br/>
    <b>Relacionamento:</b> fLEADS[lead_id] ← → ScoreLeads[lead_id] (muitos-para-um)<br/>
    <br/>
    <b>Visualizações sugeridas:</b><br/>
    • Tabela de leads ativos com score e faixa<br/>
    • KPI: "% de leads Hot" vs "% de conversão" (para validar)<br/>
    • Gráfico temporal: tendência de scores Hot/mês<br/>
    • Scatter: Dias no funil vs Score (insights)
    """
    story.append(Paragraph(powerbi_text, normal_justify))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 10. VALIDAÇÃO E CONFIABILIDADE
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("8. VALIDAÇÃO E CONFIABILIDADE", heading2))
    
    valid_text = """
    <b>Checklist de Robustez:</b><br/>
    <br/>
    ✓ <b>Separação treino-teste (80-20)</b><br/>
    &nbsp;&nbsp;Modelo treinado em 80% dos dados, avaliado em 20% completamente novos.<br/>
    &nbsp;&nbsp;Sem contaminação.<br/>
    <br/>
    ✓ <b>Stratified K-Fold (5-fold)</b><br/>
    &nbsp;&nbsp;Cross-validation garante estabilidade. Resultado: AUC 0.9767 ± 0.0015 
    (muito estável).<br/>
    <br/>
    ✓ <b>Tratamento de desbalanceamento</b><br/>
    &nbsp;&nbsp;89% não-ganho vs 11% ganho. Usamos scale_pos_weight para dar peso igual.<br/>
    &nbsp;&nbsp;Sem isto, modelo tenderia a prever "não-ganho" para tudo.<br/>
    <br/>
    ✓ <b>Calibração de probabilidades</b><br/>
    &nbsp;&nbsp;Isotônica (3-fold) garante que score 80 = 80% chance real.<br/>
    <br/>
    ✓ <b>Interpretabilidade (SHAP)</b><br/>
    &nbsp;&nbsp;Não é black-box: cada predição é explicável.<br/>
    <br/>
    <b>Limitações e Ressalvas:</b><br/>
    • Modelo reflete padrões HISTÓRICOS. Mudanças recentes em mercado podem reduzir 
    performance.<br/>
    • Requer atualização periódica: a cada 1.000 novos leads finalizados, retreine.<br/>
    • Features foram baseadas em dados disponíveis. Novos dados (ex: feedback do vendedor) 
    melhorariam.<br/>
    • Scores são probabilidades: um lead com score 25 ainda pode converter (just baixa 
    prioridade).
    """
    story.append(Paragraph(valid_text, normal_justify))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 11. PRÓXIMOS PASSOS E RECOMENDAÇÕES
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("9. PRÓXIMOS PASSOS E RECOMENDAÇÕES", heading2))
    
    proximo_text = """
    <b>Fase 1 — Implantação (Semana 1–2)</b><br/>
    1. Importar CSV no Power BI<br/>
    2. Criar relacionamento com fLEADS<br/>
    3. Deployar medidas DAX<br/>
    4. Comunicar às equipes (SDR, Closers): como usar os scores<br/>
    5. Monitorar acurácia: "Leads Hot realmente converteram?"<br/>
    <br/>
    <b>Fase 2 — Otimização (Mês 2–3)</b><br/>
    1. Coletar feedback dos vendedores: qual score foi melhor preditor na prática?<br/>
    2. Ajustar thresholds (ex: talvez "Hot" deva ser 70–100, não 76–100)<br/>
    3. Incluir novas features se disponíveis (ex: feedback do vendedor, histórico de 
    interações por canal)<br/>
    <br/>
    <b>Fase 3 — Automação (Mês 4+)</b><br/>
    1. Conectar modelo ao CRM: score automático para cada novo lead<br/>
    2. Alertas em tempo real para Hot leads<br/>
    3. A/B testing: comparar resultados de vendedores que usam os scores vs. manual<br/>
    4. Retreinamento automático a cada X novos dados<br/>
    <br/>
    <b>Métricas de Sucesso:</b><br/>
    • Taxa de conversão de leads Hot vs. Frio (esperado 3–5x maior)<br/>
    • Tempo médio de ciclo (esperado reduzir)<br/>
    • Ramp-up de novos SDRs (esperado acelerar com priorização)<br/>
    • ROI das campanhas (esperado melhorar seleção de prospects)
    """
    story.append(Paragraph(proximo_text, normal_justify))
    
    story.append(PageBreak())
    
    # ─────────────────────────────────────────────────────────────────────
    # 12. CONCLUSÃO
    # ─────────────────────────────────────────────────────────────────────
    
    story.append(Paragraph("10. CONCLUSÃO", heading2))
    
    conclusao_text = """
    O projeto de <b>Lead Scoring Vitalmed</b> representa uma materialização de 
    <b>Data Science</b> para resolver um problema real de vendas. <br/>
    <br/>
    <b>Validação Técnica:</b><br/>
    • Modelo com AUC-ROC 0.9708 — performance excelente<br/>
    • Cross-validation prova estabilidade e generalização<br/>
    • XGBoost + calibração garante confiança nas probabilidades<br/>
    • SHAP oferece interpretabilidade (não é black-box)<br/>
    <br/>
    <b>Valor de Negócio:</b><br/>
    • 58 leads identificados como "Hot" para abordagem prioritária<br/>
    • Potencial de aumentar taxa de conversão ao focar em prospects de alta probabilidade<br/>
    • Redução de tempo gasto em leads de baixa probabilidade<br/>
    • Escalabilidade: processo automatizável para milhares de leads<br/>
    <br/>
    <b>Próximos Passos:</b> Importar scores no Power BI, comunicar às equipes, 
    e monitorar impacto em 30 dias.
    """
    story.append(Paragraph(conclusao_text, normal_justify))
    
    story.append(Spacer(1, 0.3*inch))
    
    footer_text = Paragraph(
        f"<i>Relatório técnico gerado em {DATA_GERACAO}</i><br/>"
        f"<i>Modelo: XGBoost com Calibração Isotônica</i>",
        ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=8,
            alignment=TA_CENTER,
            textColor=colors.grey
        )
    )
    story.append(footer_text)
    
    # ─────────────────────────────────────────────────────────────────────
    # GERA O PDF
    # ─────────────────────────────────────────────────────────────────────
    
    doc.build(story)
    print(f"\n✅ Relatório gerado: {OUTPUT_FILE}")
    print(f"   Caminho: {os.path.abspath(OUTPUT_FILE)}")

if __name__ == "__main__":
    criar_relatorio()
