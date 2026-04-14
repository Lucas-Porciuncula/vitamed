"""
=============================================================================
LEAD SCORING — PROBABILIDADE DE CONVERSÃO
Projeto: Vitalmed — Canais Digitais
=============================================================================

COMO FUNCIONA:
  1. Extrai os dados do Power BI via DAX (mcp ou exportação manual)
  2. Engenharia de features com os campos disponíveis em fLEADS
  3. Treina um XGBoost (classificação binária: Ganho vs. Não Ganho)
  4. Gera score 0–100 para cada lead ativo
  5. Exporta tabela CSV pronta para importar no Power BI como nova tabela

DEPENDÊNCIAS:
  pip install pandas scikit-learn xgboost shap matplotlib openpyxl

ENTRADA ESPERADA:
  Um CSV com os dados de fLEADS (exportado do Power BI via Power Query ou
  copiado da visualização de tabela). Colunas mínimas necessárias:
    lead_id, vendedor, Estágio, Data de criação, Hora de criação,
    Data de fechamento, Canais, Praças, Equipe de Vendas,
    Receita Esperada, motivo da perda, campanhas, Dias

  Ou rode direto via Python + análise SSAS se tiver o pacote pyadomd.
=============================================================================
"""

import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, StratifiedKFold, cross_val_score
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import (
    classification_report, roc_auc_score,
    confusion_matrix, ConfusionMatrixDisplay
)
from sklearn.pipeline import Pipeline
from sklearn.calibration import CalibratedClassifierCV
import xgboost as xgb
import shap
import matplotlib.pyplot as plt
import warnings
import os

warnings.filterwarnings("ignore")

# =============================================================================
# 1. CARREGAMENTO DOS DADOS
# =============================================================================

def carregar_dados(caminho_csv: str) -> pd.DataFrame:
    """
    Carrega o CSV exportado do Power BI.
    Se quiser usar conexão direta ao Power BI Desktop via SSAS,
    substitua por conexão pyadomd/adodbapi (veja comentário no final).
    """
    df = pd.read_csv(
        caminho_csv,
        sep=";",
        encoding="latin-1",
        decimal=",",
        parse_dates=["Data de criação", "Data de fechamento", "date_last_stage_update"],
        dayfirst=True,
    )
    print(f"[OK] Dados carregados: {len(df):,} leads")
    return df


# =============================================================================
# 2. ENGENHARIA DE FEATURES
# =============================================================================

def extrair_hora(col_hora):
    """Extrai a hora de uma coluna de tempo (vem como string HH:MM:SS)."""
    try:
        return pd.to_datetime(col_hora, format="%H:%M:%S", errors="coerce").dt.hour
    except Exception:
        return pd.Series([np.nan] * len(col_hora))


def feature_engineering(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    print(f"\n[INFO] Estagios unicos no CSV:")
    print(df["Estágio"].value_counts())
    print()

    # --- Target: 1 = Ganho, 0 = Não Ganho (leads com desfecho definido) ---
    # Ganho se: Estágio == "Ganho" OU "Receita Esperada" > 0 e "Data de fechamento" está preenchida
    # Perdido se: tem "motivo da perda" preenchido OU Estágio contém palavras-chave de perda
    tem_motivo_perda = (~df["motivo da perda"].isna()) & (df["motivo da perda"] != "")
    eh_ganho = df["Estágio"] == "Ganho"
    eh_perdido = tem_motivo_perda | df["Estágio"].isin(["Perdido", "Desistiu", "Descartado"])
    
    # Classifica fechados (Ganho ou Perdido)
    df["tem_desfecho"] = eh_ganho | eh_perdido
    df["target"] = eh_ganho.astype(int)

    # Mantém só leads com desfecho para treino (remove leads ainda em aberto
    # que serão usados apenas para scoring)
    estagios_abertos = ["Prospecção", "Contato Pré Vendas", "Negociação", "Em Aprovação", "Sem contato"]
    df["lead_aberto"] = df["Estágio"].isin(estagios_abertos)

    # --- Feature: hora do dia de criação ---
    df["hora_criacao"] = extrair_hora(df["Hora de criação"])
    df["turno"] = pd.cut(
        df["hora_criacao"],
        bins=[-1, 6, 12, 18, 24],
        labels=["madrugada", "manhã", "tarde", "noite"]
    )

    # --- Feature: dia da semana ---
    df["dia_semana"] = pd.to_datetime(df["Data de criação"], dayfirst=True, errors="coerce").dt.dayofweek
    df["fim_de_semana"] = (df["dia_semana"] >= 5).astype(int)

    # --- Feature: mês de criação ---
    df["mes_criacao"] = pd.to_datetime(df["Data de criação"], dayfirst=True, errors="coerce").dt.month

    # --- Feature: dias no funil ---
    df["dias_funil"] = pd.to_numeric(df["Dias"], errors="coerce").fillna(0)

    # --- Feature: tem receita esperada preenchida ---
    df["receita_esperada"] = pd.to_numeric(
        df["Receita Esperada"].astype(str).str.replace(",", "."), errors="coerce"
    ).fillna(0)
    df["tem_receita_esperada"] = (df["receita_esperada"] > 0).astype(int)

    # --- Feature: tem e-mail preenchido ---
    df["tem_email"] = (~df["E-mail"].isna() & (df["E-mail"] != "")).astype(int)

    # --- Feature: canal de origem ---
    canais_map = {
        "Tráfego Pago": "trafego_pago",
        "Orgânico": "organico",
        "nenhum": "desconhecido",
        "Offline": "offline",
    }
    df["canal_norm"] = df["Canais"].map(canais_map).fillna("outro")

    # --- Feature: praça ---
    df["praca_norm"] = df["Praças"].fillna("desconhecida").str.strip().str.lower()

    # --- Feature: equipe SDR vs Closer ---
    df["equipe_tipo"] = df["Equipe de Vendas"].apply(
        lambda x: "sdr" if "Pré-Vendas" in str(x) or "Suporte" in str(x)
        else ("televendas" if "Televendas" in str(x)
              else ("vendas_diretas" if "Vendas Diretas" in str(x)
                    else "outras"))
    )

    # --- Feature: campanha preenchida ---
    df["tem_campanha"] = (~df["campanhas"].isna() & (df["campanhas"] != "")).astype(int)

    # --- Feature: tempo até primeiro fechamento (proxy de velocidade) ---
    data_cria = pd.to_datetime(df["Data de criação"], dayfirst=True, errors="coerce")
    data_fecha = pd.to_datetime(df["Data de fechamento"], dayfirst=True, errors="coerce")
    df["dias_ate_fechamento"] = (data_fecha - data_cria).dt.days.clip(lower=0).fillna(-1)

    # --- Feature: horas até fechamento ---
    df["horas_fechamento"] = pd.to_numeric(
        df["Horas até fechamento"].astype(str).str.replace(",", "."), errors="coerce"
    ).fillna(-1)

    print(f"[OK] Features criadas com sucesso")
    return df


# =============================================================================
# 3. PREPARAÇÃO DO DATASET DE TREINO
# =============================================================================

FEATURES_CATEGORICAS = ["canal_norm", "praca_norm", "equipe_tipo", "turno"]

FEATURES_NUMERICAS = [
    "hora_criacao", "dia_semana", "fim_de_semana", "mes_criacao",
    "dias_funil", "receita_esperada", "tem_receita_esperada",
    "tem_email", "tem_campanha", "dias_ate_fechamento", "horas_fechamento"
]

TODAS_FEATURES = FEATURES_CATEGORICAS + FEATURES_NUMERICAS


def preparar_dataset(df: pd.DataFrame):
    """
    Separa treino (leads com desfecho) e scoring (leads em aberto).
    Codifica categorias e retorna X, y, encoders.
    """
    encoders = {}
    df_enc = df.copy()

    for col in FEATURES_CATEGORICAS:
        le = LabelEncoder()
        df_enc[col] = df_enc[col].astype(str)
        le.fit(df_enc[col])
        df_enc[col + "_enc"] = le.transform(df_enc[col])
        encoders[col] = le

    features_enc = [c + "_enc" for c in FEATURES_CATEGORICAS] + FEATURES_NUMERICAS

    # Dataset de treino: apenas leads com desfecho definido (Ganho OU Perdido)
    df_treino = df_enc[df_enc["tem_desfecho"]].copy()
    # Dataset de scoring: leads ainda em aberto
    df_aberto = df_enc[df_enc["lead_aberto"]].copy()

    X = df_treino[features_enc].fillna(-1)
    y = df_treino["target"]

    X_score = df_aberto[features_enc].fillna(-1)

    # VALIDAÇÃO: garantir que temos ambas as classes
    n_ganho = (y == 1).sum()
    n_perdido = (y == 0).sum()
    
    print(f"[OK] Treino: {len(df_treino):,} leads finalizados")
    print(f"    [+] Ganhos: {n_ganho:,} ({100*n_ganho/len(y) if len(y)>0 else 0:.1f}%)")
    print(f"    [-] Perdidos: {n_perdido:,} ({100*n_perdido/len(y) if len(y)>0 else 0:.1f}%)")
    print(f"[OK] Para scoring: {len(df_aberto):,} leads em aberto")
    
    if len(y) == 0:
        raise ValueError("[ERRO] Nenhum lead com desfecho definido (Ganho ou Perdido) encontrado no CSV. Verifique os estagios esperados.")
    
    if n_ganho == 0 or n_perdido == 0:
        print(f"\n[AVISO] Dataset desbalanceado - faltam exemplos de {'Perdido' if n_ganho > 0 else 'Ganho'}.")
        print("Isto pode afetar a qualidade do modelo.\n")

    return X, y, X_score, df_aberto, features_enc, encoders


# =============================================================================
# 4. TREINAMENTO DO MODELO
# =============================================================================

def treinar_modelo(X, y):
    """
    Treina XGBoost com calibração de probabilidade.
    Usa scale_pos_weight para lidar com desbalanceamento (poucos Ganhos).
    """
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )

    # Calcula peso para desbalanceamento
    ratio = (y == 0).sum() / (y == 1).sum()
    print(f"[OK] Desbalanceamento - ratio negativo/positivo: {ratio:.1f}x")

    modelo_base = xgb.XGBClassifier(
        n_estimators=300,
        max_depth=5,
        learning_rate=0.05,
        subsample=0.8,
        colsample_bytree=0.8,
        scale_pos_weight=ratio,
        use_label_encoder=False,
        eval_metric="auc",
        random_state=42,
        n_jobs=-1,
    )

    # Calibração isotônica para scores bem calibrados (não só ranking)
    modelo = CalibratedClassifierCV(modelo_base, method="isotonic", cv=3)
    modelo.fit(X_train, y_train)

    # Avaliação
    y_pred = modelo.predict(X_test)
    y_prob = modelo.predict_proba(X_test)[:, 1]

    auc = roc_auc_score(y_test, y_prob)
    print(f"\n[METRICS] AUC-ROC no teste: {auc:.4f}")
    print("\n[REPORT] Classificacao:")
    print(classification_report(y_test, y_pred, target_names=["Não Ganho", "Ganho"]))

    # Cross-validation
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    scores_cv = cross_val_score(modelo_base, X, y, cv=cv, scoring="roc_auc", n_jobs=-1)
    print(f"\n[METRICS] Cross-Val AUC (5-fold): {scores_cv.mean():.4f} +/- {scores_cv.std():.4f}")

    return modelo, X_test, y_test, y_prob


# =============================================================================
# 5. INTERPRETABILIDADE (SHAP)
# =============================================================================

def analisar_features(modelo, X, feature_names, output_dir="."):
    """
    Gera gráfico SHAP de importância das features.
    Usa o estimador base do CalibratedClassifierCV.
    """
    print("\n[INFO] Calculando importancia SHAP...")

    # Extrai o estimador XGBoost de dentro da calibração
    estimador_xgb = modelo.calibrated_classifiers_[0].estimator

    explainer = shap.TreeExplainer(estimador_xgb)
    shap_values = explainer.shap_values(X)

    plt.figure(figsize=(10, 6))
    shap.summary_plot(
        shap_values, X,
        feature_names=feature_names,
        plot_type="bar",
        show=False,
        max_display=15
    )
    plt.title("Importância das Features — Lead Scoring Vitalmed", fontsize=13)
    plt.tight_layout()
    path = os.path.join(output_dir, "shap_importancia_features.png")
    plt.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"[OK] Grafico SHAP salvo em: {path}")


# =============================================================================
# 6. SCORING DOS LEADS EM ABERTO
# =============================================================================

def gerar_scores(modelo, X_score, df_aberto, output_dir=".") -> pd.DataFrame:
    """
    Aplica o modelo aos leads em aberto e exporta CSV com scores.
    """
    probabilidades = modelo.predict_proba(X_score.fillna(-1))[:, 1]
    scores = (probabilidades * 100).round(1)

    resultado = df_aberto[
        ["lead_id", "Nome do contato", "vendedor", "Estágio",
         "Canais", "Praças", "Equipe de Vendas", "Data de criação",
         "Receita Esperada", "campanhas"]
    ].copy()

    resultado["score_conversao"] = scores
    resultado["probabilidade_pct"] = probabilidades.round(4)

    # Classificação em faixas para uso no Power BI
    resultado["faixa_score"] = pd.cut(
        scores,
        bins=[0, 25, 50, 75, 101],
        labels=["Frio (0-25)", "Morno (26-50)", "Quente (51-75)", "Hot (76-100)"],
        right=False
    )

    resultado = resultado.sort_values("score_conversao", ascending=False)

    path_csv = os.path.join(output_dir, "lead_scores_vitalmed.csv")
    resultado.to_csv(path_csv, index=False, sep=";", decimal=",", encoding="utf-8-sig")
    print(f"\n[OK] Scores exportados: {path_csv}")
    print(f"\n[STATS] Distribuicao dos scores:")
    print(resultado["faixa_score"].value_counts().sort_index())

    return resultado


# =============================================================================
# 7. MEDIDA DAX PARA IMPORTAR NO POWER BI
# =============================================================================

DAX_PARA_POWERBI = """
-- ============================================================
-- COMO USAR OS SCORES NO POWER BI
-- ============================================================
-- 1. No Power Query (Power BI Desktop):
--    Home > Obter Dados > CSV > selecione "lead_scores_vitalmed.csv"
--    Renomeie a tabela para: ScoreLeads
--
-- 2. Crie o relacionamento:
--    fLEADS[lead_id]  →  ScoreLeads[lead_id]  (muitos-para-um)
--
-- 3. Use esta medida DAX para exibir o score no cartão/tabela:

Score Conversão =
VAR _score = RELATED(ScoreLeads[score_conversao])
RETURN
    IF(ISBLANK(_score), "—", FORMAT(_score, "0.0") & " pts")

-- 4. Medida de cor para semáforo (formatação condicional):

Cor Score =
VAR _s = RELATED(ScoreLeads[score_conversao])
RETURN
    SWITCH(
        TRUE(),
        _s >= 76, "#1D9E75",   -- Verde: Hot
        _s >= 51, "#378ADD",   -- Azul: Quente
        _s >= 26, "#EF9F27",   -- Laranja: Morno
        "#D85A30"               -- Vermelho: Frio
    )
"""


# =============================================================================
# 8. PIPELINE PRINCIPAL
# =============================================================================

def main(caminho_csv: str, output_dir: str = "."):
    print("=" * 60)
    print("  LEAD SCORING — VITALMED  ")
    print("=" * 60)

    os.makedirs(output_dir, exist_ok=True)

    # Carrega e processa
    df_raw = carregar_dados(caminho_csv)
    df = feature_engineering(df_raw)

    # Prepara datasets
    X, y, X_score, df_aberto, features_enc, encoders = preparar_dataset(df)

    # Treina
    modelo, X_test, y_test, y_prob = treinar_modelo(X, y)

    # Interpretabilidade
    analisar_features(modelo, X, features_enc, output_dir)

    # Gera scores para leads em aberto
    df_scores = gerar_scores(modelo, X_score, df_aberto, output_dir)

    # Salva instruções DAX
    with open(os.path.join(output_dir, "instrucoes_powerbi.dax"), "w", encoding="utf-8") as f:
        f.write(DAX_PARA_POWERBI)
    print(f"[OK] Instrucoes DAX salvas em: instrucoes_powerbi.dax")

    print("\n[PROXIMOS] Passos seguintes:")
    print("  1. Importe 'lead_scores_vitalmed.csv' no Power BI")
    print("  2. Crie relacionamento com fLEADS pelo campo lead_id")
    print("  3. Use as medidas DAX do arquivo instrucoes_powerbi.dax")
    print("  4. Adicione coluna 'score_conversao' na tabela de Vendedores")
    print("  5. Filtre por 'faixa_score = Hot' para lista de prioridade SDR")

    return modelo, df_scores


# =============================================================================
# EXECUÇÃO
# =============================================================================

if __name__ == "__main__":
    # Substitua pelo caminho do seu CSV exportado do Power BI
    CAMINHO_CSV = r"C:\Users\laporciuncula\Downloads\Pasta1.csv"
    OUTPUT_DIR = "output_lead_scoring"

    modelo, df_scores = main(CAMINHO_CSV, OUTPUT_DIR)


# =============================================================================
# NOTA: CONEXÃO DIRETA AO POWER BI DESKTOP (sem exportar CSV)
# =============================================================================
# Se preferir conectar diretamente ao modelo aberto no Power BI Desktop,
# instale: pip install pyadomd
#
# from pyadomd import Pyadomd
# conn_str = "Provider=MSOLAP;Data Source=localhost:55767;"
# dax = """
# EVALUATE
#   SELECTCOLUMNS(
#     fLEADS,
#     "lead_id", fLEADS[lead_id],
#     "estagio", fLEADS[Estágio],
#     ... demais colunas
#   )
# """
# with Pyadomd(conn_str) as conn:
#     df = pd.DataFrame(conn.cursor().execute(dax).fetchall())
# =============================================================================