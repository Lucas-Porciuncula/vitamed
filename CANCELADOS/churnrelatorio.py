import pandas as pd
from dotenv import load_dotenv
from churn import query_tempo_casa, gerar_engine

load_dotenv()

# Mantenha apenas a fonte da verdade: os períodos mensais
PERIODOS = [
    ("20241226", "20250125", "Jan-2025"),
    ("20250126", "20250225", "Fev-2025"),
    ("20250226", "20250325", "Mar-2025"),
    ("20250326", "20250425", "Abr-2025"),
    ("20250426", "20250525", "Mai-2025"),
    ("20250526", "20250625", "Jun-2025"),
    ("20250626", "20250725", "Jul-2025"),
    ("20250726", "20250825", "Ago-2025"),
    ("20250826", "20250925", "Set-2025"),
    ("20250926", "20251025", "Out-2025"),
    ("20251026", "20251125", "Nov-2025"),
    ("20251126", "20251225", "Dez-2025"),
    ("20251226", "20260125", "Jan-2026"),
    ("20260126", "20260225", "Fev-2026"),
    ("20260226", "20260325", "Mar-2026"),
]

MODO = "SEMESTRE"  # "PERIODO" ou "SEMESTRE"

def agrupar_por_semestre(periodos):
    """Agrupa os períodos mensais em semestres automaticamente."""
    df_temp = pd.DataFrame(periodos, columns=['ini', 'fim', 'label'])
    
    # Extrai mês e ano do label (Ex: Jan-2025)
    # Mapeamento simples: Jan-Jun = S1, Jul-Dez = S2
    meses_map = {
        'Jan': 'S1', 'Fev': 'S1', 'Mar': 'S1', 'Abr': 'S1', 'Mai': 'S1', 'Jun': 'S1',
        'Jul': 'S2', 'Ago': 'S2', 'Set': 'S2', 'Out': 'S2', 'Nov': 'S2', 'Dez': 'S2'
    }
    
    df_temp['mes_pt'] = df_temp['label'].str.split('-').str[0]
    df_temp['ano'] = df_temp['label'].str.split('-').str[1]
    df_temp['semestre'] = df_temp['mes_pt'].map(meses_map)
    df_temp['grupo'] = df_temp['semestre'] + "-" + df_temp['ano']
    
    semestres_agrupados = []
    for grupo, data in df_temp.groupby('grupo', sort=False):
        # Pega a data inicial do primeiro mês e a final do último mês do bloco
        data_ini = data['ini'].iloc[0]
        data_fim = data['fim'].iloc[-1]
        semestres_agrupados.append((data_ini, data_fim, grupo))
        
    return semestres_agrupados

def main():
    engine = gerar_engine()
    
    # Lógica matadora: se for semestre, ele calcula sozinho com base nos períodos
    if MODO == "SEMESTRE":
        base = agrupar_por_semestre(PERIODOS)
    else:
        base = PERIODOS

    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"Churn_TempoCasa_{MODO}_{timestamp}.xlsx"

    with pd.ExcelWriter(nome_arquivo, engine="openpyxl") as writer:
        for data_ini, data_fim, label in base:
            print(f"▶ Processando {MODO}: {label} ({data_ini} a {data_fim})")
            
            df = pd.read_sql(query_tempo_casa, engine, params=(data_ini, data_fim))
            df.to_excel(writer, sheet_name=label[:31], index=False) # Excel limita sheet a 31 chars
            
            print(f"   ✔ {len(df)} registros")

    print(f"\n✅ Arquivo gerado: {nome_arquivo}")
    engine.dispose()

if __name__ == "__main__":
    main()