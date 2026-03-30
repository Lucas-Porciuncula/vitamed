import pandas as pd
import pyodbc
from dotenv import load_dotenv
import os
from datetime import datetime


class SistemaAcompanhamentoMetas:
    """
    Sistema para acompanhamento de metas de vendas por equipe.
    Processa dados de propostas e compara com metas estabelecidas.ai 
    """
    
    def __init__(self):
        self.conn = None
        self.METAS = None
        self.DADOS_EQUIPE = None
        self.fPROPOSTAS = None
        
        # ============================================================
        # 🔴 CONFIGURAÇÃO AUTOMÁTICA DE PERÍODO (26 a 25)
        # ============================================================
        hoje = pd.Timestamp.today().normalize()
        self.DATA_INICIO = pd.Timestamp('2026-02-26')
        self.DATA_FIM = pd.Timestamp('2026-03-30')    # Configuração fixa para testes

        ##if hoje.day >= 26:
        ##    self.DATA_INICIO = hoje.replace(day=26)
        ##    self.DATA_FIM = (self.DATA_INICIO + pd.DateOffset(months=1)).replace(day=25)
        ##else:
        ##    self.DATA_FIM = hoje.replace(day=25)
        ##    self.DATA_INICIO = (self.DATA_FIM - pd.DateOffset(months=1)).replace(day=26)

        self.DATA_FIM = hoje  # DataFim sempre é hoje
        print("Data início:", self.DATA_INICIO.strftime("%d/%m/%Y"))
        print("Data fim:", self.DATA_FIM.strftime("%d/%m/%Y"))
        # ============================================================
    # ============================================================
        
        self.FILIAL_FILTRO = 'Salvador'
        self.PRODUTO_FILTRO = 'APH TRADICIONAL'
        self.STATUS_FILTRO = 'Aprovado'
    
    def conectar_banco(self):
        """Estabelece conexão com o banco de dados SQL Server."""
        load_dotenv()
        
        server = os.getenv("SERVER")
        database = os.getenv("DB_NAME")
        user = os.getenv("DB_USER")
        password = os.getenv("DB_PASS6")
        
        self.conn = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={user};"
            f"PWD={password};"
        )
        
        print("=" * 50)
        print("✓ Conexão bem-sucedida ao banco de dados.")
        print("=" * 50)
    
    def carregar_bases(self):
        """Carrega as bases de metas, equipes e propostas."""
        caminho_base = r'\\192.168.1.26\adm de vendas\Analytics\Relatórios\Acompanhamento Diário\Bases\BASE_VENDEDORES.xlsx'
        
        # Carrega base de metas
        self.METAS = pd.read_excel(
            caminho_base,
            sheet_name='EQUIPES',
            usecols='A:E'
        )
        
        # Carrega dados de equipe
        self.DADOS_EQUIPE = pd.read_excel(
            caminho_base,
            sheet_name='VENDEDOR-EQUIPE',
            usecols='A:I'
        )
        
        # Carrega propostas do banco
        self.fPROPOSTAS = pd.read_sql(
            "SELECT * FROM VENDAS_ONLINE_COMERCIAL",
            self.conn
        )
        
        print("✓ Bases carregadas com sucesso.")
    
    @staticmethod
    def normalizar_mes_ano(valor):
        """
        Normaliza diferentes formatos de data para o primeiro dia do mês.
        Aceita: datetime, ISO, 'jan/2024', etc.
        """
        if pd.isna(valor):
            return pd.NaT
        
        meses = {
            'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4,
            'mai': 5, 'jun': 6, 'jul': 7, 'ago': 8,
            'set': 9, 'out': 10, 'nov': 11, 'dez': 12
        }
        
        if isinstance(valor, pd.Timestamp):
            return valor.to_period('M').to_timestamp()
        
        valor = str(valor).strip().lower()
        
        # Tenta formato ISO
        try:
            return pd.to_datetime(valor).to_period('M').to_timestamp()
        except:
            pass
        
        # Tenta formato 'jan/2024'
        try:
            mes, ano = valor.split('/')
            mes_num = meses.get(mes[:3])
            if mes_num:
                return pd.Timestamp(year=int(ano), month=mes_num, day=1)
        except:
            pass
        
        return pd.NaT
    
    def processar_metas(self):
        """Processa e filtra as metas para a filial e último mês disponível."""
        # Normaliza datas
        self.METAS['MES_ANO_DT'] = self.METAS['MÊS/ANO'].apply(self.normalizar_mes_ano)
        
        # Filtra por filial
        self.METAS = self.METAS[self.METAS['FILIAL'] == self.FILIAL_FILTRO]
        
        # Identifica último mês disponível
        ultimo_mes = self.METAS['MES_ANO_DT'].max().to_period('M') 
        
        # Filtra somente o último mês
        METAS_ULTIMO_MES = self.METAS[
            self.METAS['MES_ANO_DT'].dt.to_period('M') == ultimo_mes
        ]
        
        # Agrega meta por equipe
        METAS_FINAL = (
            METAS_ULTIMO_MES
            .groupby('EQUIPE', as_index=False)
            .agg({'META': 'sum'})
        )
        
        print("\n✓ Metas processadas (último mês):")
        print(METAS_FINAL.head())
        
        return METAS_FINAL
    
    @staticmethod
    def pad_cod_vendedor(x):
        """Padroniza código do vendedor para 6 dígitos."""
        try:
            return f"{int(float(x)):06d}"
        except:
            return None
    
    def processar_propostas(self):
        """
        Processa propostas e mescla com dados de equipe.
        Aplica filtros de filial, status e produto.
        """
        # Padroniza códigos de vendedor
        self.fPROPOSTAS['COD VENDEDOR'] = self.fPROPOSTAS['COD VENDEDOR'].apply(self.pad_cod_vendedor)
        self.DADOS_EQUIPE['COD VENDEDOR'] = self.DADOS_EQUIPE['COD VENDEDOR'].apply(self.pad_cod_vendedor)

        self.fPROPOSTAS['FILIAL'] = self.fPROPOSTAS['FILIAL'].str.strip().str.title()
        
        # Mescla propostas com equipes
        mescla = pd.merge(
            self.fPROPOSTAS,
            self.DADOS_EQUIPE,
            how='left',
            on='COD VENDEDOR'
        )
        
        # Aplica filtros
        mescla = mescla[
            (mescla['FILIAL_x'] == self.FILIAL_FILTRO) &
            (mescla['STATUS'] == self.STATUS_FILTRO) &
            (mescla['PRODUTO'] == self.PRODUTO_FILTRO)
        ]
        
        # Seleciona colunas relevantes
        mescla = mescla[['EQUIPE', 'QTD_VIDAS', 'DATA_PROPOSTA']]
        
        # Converte data
        mescla['DATA_PROPOSTA'] = pd.to_datetime(mescla['DATA_PROPOSTA'])
        
        return mescla
    

    def diagnosticar_vendedor(self, cod_raw):
        """
        Diagnóstico completo para um código de vendedor específico.
        Chame após carregar_bases() e antes de processar_propostas().
        """
        cod = self.pad_cod_vendedor(cod_raw)
        print(f"\n🔍 DIAGNÓSTICO — Vendedor {cod_raw} → padded: '{cod}'")
        print("=" * 60)

        # 1. Verifica se o código existe na planilha de equipes
        equipe_match = self.DADOS_EQUIPE[self.DADOS_EQUIPE['COD VENDEDOR'].apply(self.pad_cod_vendedor) == cod]
        if equipe_match.empty:
            print("❌ Código NÃO encontrado em DADOS_EQUIPE (planilha VENDEDOR-EQUIPE)")
            print("   Valores próximos na planilha:")
            print(self.DADOS_EQUIPE['COD VENDEDOR'].dropna().head(10).tolist())
        else:
            print(f"✅ Encontrado em DADOS_EQUIPE:")
            print(equipe_match.to_string(index=False))

        # 2. Verifica propostas brutas no banco (sem nenhum filtro)
        props_brutas = self.fPROPOSTAS[
            self.fPROPOSTAS['COD VENDEDOR'].apply(self.pad_cod_vendedor) == cod
        ]
        print(f"\n📋 Propostas brutas no banco para este vendedor: {len(props_brutas)}")
        if not props_brutas.empty:
            cols = ['COD VENDEDOR', 'FILIAL', 'STATUS', 'PRODUTO', 'DATA_PROPOSTA', 'QTD_VIDAS']
            print(props_brutas[[c for c in cols if c in props_brutas.columns]].to_string(index=False))

        # 3. Verifica cada filtro individualmente
        if not props_brutas.empty:
            print("\n🔎 Checagem dos filtros:")
            print(f"   FILIAL esperada : '{self.FILIAL_FILTRO}'")
            print(f"   STATUS esperado : '{self.STATUS_FILTRO}'")
            print(f"   PRODUTO esperado: '{self.PRODUTO_FILTRO}'")
            print(f"   Valores reais   :")
            print(props_brutas[['FILIAL', 'STATUS', 'PRODUTO']].drop_duplicates().to_string(index=False))

            # 4. Verifica período
            props_brutas = props_brutas.copy()
            props_brutas['DATA_PROPOSTA'] = pd.to_datetime(props_brutas['DATA_PROPOSTA'])
            no_periodo = props_brutas[
                props_brutas['DATA_PROPOSTA'].between(self.DATA_INICIO, self.DATA_FIM)
            ]
            print(f"\n📅 Propostas dentro do período ({self.DATA_INICIO.date()} a {self.DATA_FIM.date()}): {len(no_periodo)}")

        print("=" * 60)


    def filtrar_por_periodo(self, df):
        """
        Filtra dataframe pelo período configurado (DATA_INICIO e DATA_FIM).
        
        ============================================================
        🔴 ATENÇÃO: O período é definido nas variáveis da classe:
           - self.DATA_INICIO
           - self.DATA_FIM
        ============================================================
        """
        print(f"\n📅 Filtrando período: {self.DATA_INICIO.date()} até {self.DATA_FIM.date()}")
        
        df_filtrado = df[
            df['DATA_PROPOSTA'].between(self.DATA_INICIO, self.DATA_FIM)
        ]
        
        print(f"✓ Registros no período: {len(df_filtrado)}")
        print(f"✓ Soma no período: {df_filtrado['QTD_VIDAS'].sum()}")
        
        return df_filtrado
    
    def agregar_vidas_por_equipe(self, df):
        """Agrega quantidade de vidas por equipe."""
        resultado = (
            df
            .groupby('EQUIPE', as_index=False)
            .agg({'QTD_VIDAS': 'sum'})
        )
        print("\n✓ Vidas agregadas por equipe:")
        print(resultado.head())
        
        return resultado
    
    def gerar_relatorio_final(self, vidas_por_equipe, metas):
        """Gera relatório final com vidas, metas e percentual de atingimento."""
        
        
        final = pd.merge(
            vidas_por_equipe,
            metas,
            how='left',
            on='EQUIPE'
        )
        
        final = final.rename(columns={
            'QTD_VIDAS': 'Vidas APH',
            'META': 'Meta'
        })
        
        # Calcula percentual da meta
        final['% da Meta'] = (
            (final['Vidas APH'] / final['Meta'])
            .fillna(0)
            .mul(100)
            .round(1)
            .astype(str) + '%'
        )
        
        # Ordena colunas
        final = final[['EQUIPE', 'Vidas APH', 'Meta', '% da Meta']]
        
        return final
    
    def executar(self):
        """Executa todo o fluxo de processamento."""
        print("\n🚀 Iniciando processamento...\n")
        
        # 1. Conecta ao banco
        self.conectar_banco()
        
        # 2. Carrega bases
        self.carregar_bases()

        self.diagnosticar_vendedor(937000)
        
        # 3. Processa metas
        metas_final = self.processar_metas()
        
        # 4. Processa propostas
        propostas = self.processar_propostas()
        
        # 5. Filtra por período configurado
        propostas_periodo = self.filtrar_por_periodo(propostas)
        
        # 6. Agrega vidas por equipe
        vidas_por_equipe = self.agregar_vidas_por_equipe(propostas_periodo)
        
        # 7. Gera relatório final
        relatorio_final = self.gerar_relatorio_final(vidas_por_equipe, metas_final)
        
        print("\n" + "=" * 50)
        print("📊 RELATÓRIO FINAL")
        print("=" * 50)
        print(relatorio_final.to_string(index=False))
        print("=" * 50)
        
        return relatorio_final
    
    def fechar_conexao(self):
        """Fecha a conexão com o banco de dados."""
        if self.conn:
            self.conn.close()
            print("\n✓ Conexão fechada.")

