QUERY = """

SELECT * FROM VW_CANCELADOS_MESPLS
"""

QUERY_REATIVADOS = """

SELECT * FROM VW_REATIVADOS_MESPLS
"""

QUERY_CANCELADOS_FEIRA = """

SELECT * FROM VW_CANCELADOS_FSA_MESPLS

"""

QUERY_REATIVADOS_FEIRA = """


SELECT * FROM VW_REATIVADOS_FSA_MESPLS

"""

QUERY_PURA_PARA_RELATORIO = """


WITH DatasFechamento AS (
    SELECT
        CAST(? AS DATE) AS DataInicio,
        CAST(? AS DATE) AS DataFim
),
UltimoBloqueio AS (
    SELECT 
        BCA_MATRIC,
        BCA_TIPREG,
        MAX(R_E_C_N_O_) AS UltimoRecno
    FROM BCA010
    CROSS JOIN DatasFechamento D
    WHERE D_E_L_E_T_ = ''
      AND BCA_TIPO = '0'
      AND BCA_DATA BETWEEN D.DataInicio AND D.DataFim
    GROUP BY BCA_MATRIC, BCA_TIPREG
),
DadosPrincipais AS (
    SELECT DISTINCT
        -- Identificação do Beneficiário
        BA1.BA1_CODMAT,
        BA1.BA1_NOMUSR,
        BA1.BA1_CPFUSR,
        BA1.BA1_TELEFO,
        
        -- Endereço
        BA1.BA1_DISTRA,
        BA1.BA1_BAIRRO,
        
        -- Dados do Plano
       CASE 
    WHEN BI3_VTMOD = 1 THEN 'Tradicional'
    WHEN BI3_VTMOD = 2 THEN 'Participativo'
    WHEN BI3_VTMOD = 3 THEN 
        CASE 
            WHEN BI3_CODIGO = '0008' THEN 'OMT Vitalfone'
            ELSE 'Prot. Total'
        END
    WHEN BI3_VTMOD = 4 THEN 'Área Protegida'
    WHEN BI3_VTMOD = 5 THEN 'Vitalmed Online'
    WHEN BI3_VTMOD = 6 THEN 'APH A.P. Resid.'
    WHEN BI3_VTMOD = 7 THEN 'Vitalmed Online PJ'
    WHEN BI3_VTMOD = 8 THEN 'APH Cli. Assistidos'
    WHEN BI3_VTMOD = 9 THEN 'Func V Online'
    ELSE 'Outros'
END AS DescricaoPlano,

        BA1.BA1_VTCLAS,
        
        -- Dados do Contrato
        CASE 
            WHEN BA3.BA3_COBNIV = '1' THEN BA3.BA3_CODCLI
            WHEN BQC.BQC_COBNIV = '1' THEN BQC.BQC_CODCLI
            WHEN BT5.BT5_COBNIV = '1' THEN BT5.BT5_CODCLI
        END AS CodigoCliente,
        
        CASE 
            WHEN BT5.BT5_COBNIV = '1' THEN BT5.BT5_NOME
            ELSE BQC.BQC_DESCRI
        END AS DescricaoContrato,
        
        -- Dados do Vendedor
        SA3.A3_NOME AS NomeVendedor,
        
        -- Dados do Bloqueio
        BG3.BG3_DESBLO AS MotivoBloqueio,
        BCA.BCA_USUOPE AS UsuarioBloqueio,
        CONVERT(DATE, BCA.BCA_DATA) AS DataBloqueio,
        
        -- Datas formatadas
        CONVERT(DATE, BA1.BA1_DATNAS) AS DataNascimento,
        CONVERT(DATE, BA1.BA1_DATINC) AS DataInclusao,
        CONVERT(DATE, BA1.BA1_DATBLO) AS DataBloqueioRegistro,
        
        -- Cálculos de Tempo CORRIGIDOS
        -- Tempo de Casa: da inclusão até o bloqueio (em meses)
        DATEDIFF(MONTH, 
            CONVERT(DATE, BA1.BA1_DATINC), 
            COALESCE(
    NULLIF(CONVERT(DATE, NULLIF(LTRIM(RTRIM(BA1.BA1_DATBLO)), '')), '1900-01-01'),
    GETDATE()
)
        ) AS TempoContratacaoAno,
        
        -- Idade no momento do bloqueio
        DATEDIFF(YEAR, 
            CONVERT(DATE, BA1.BA1_DATNAS), 
            COALESCE(
    NULLIF(CONVERT(DATE, NULLIF(LTRIM(RTRIM(BA1.BA1_DATBLO)), '')), '1900-01-01'),
    GETDATE()
)
        ) AS IdadeAtual,
        
        -- Valores
        BDK.BDK_VALOR AS ValorMensalidade,
        BQL.BQL_DESCRI AS TipoPagamento,
        
        -- Contador de Atendimentos
        TA.QTDATE AS QuantidadeAtendimentos,
        
        -- Ocorrência mais recente
        (
            SELECT TOP 1 SE12.E1_DESCOC
            FROM SE1010 SE12
            WHERE SE12.D_E_L_E_T_ = ''
              AND SE12.E1_CLIENTE = BA3.BA3_CODCLI
              AND SE12.E1_CODINT = BA3.BA3_CODINT
              AND SE12.E1_CODEMP = BA3.BA3_CODEMP
              AND SE12.E1_MATRIC = BA3.BA3_MATRIC
              AND SE12.E1_VENCTO < GETDATE()
            ORDER BY SE12.E1_EMISSAO DESC
        ) AS UltimaOcorrencia

    FROM BA1010 BA1
    
    -- Joins Obrigatórios
    INNER JOIN BA3010 BA3 ON (
        BA3.BA3_CODINT = BA1.BA1_CODINT
        AND BA3.BA3_CODEMP = BA1.BA1_CODEMP
        AND BA3.BA3_MATRIC = BA1.BA1_MATRIC
        AND BA3.D_E_L_E_T_ = ''
    )
    
    INNER JOIN SA3010 SA3 ON (
        SA3.A3_COD = BA1.BA1_CODVEN
        AND SA3.D_E_L_E_T_ = ''
    )
    
    INNER JOIN UltimoBloqueio UB ON (
        UB.BCA_MATRIC = BA1.BA1_CODINT + BA1.BA1_CODEMP + BA1.BA1_MATRIC
        AND UB.BCA_TIPREG = BA1.BA1_TIPREG
    )
    
    INNER JOIN BCA010 BCA ON (
        BCA.R_E_C_N_O_ = UB.UltimoRecno
        AND BCA.D_E_L_E_T_ = ''
    )
    
    INNER JOIN BG3010 BG3 ON (
        BG3.BG3_CODBLO = BCA.BCA_MOTBLO
        AND BG3.D_E_L_E_T_ = ''
    )
    
    INNER JOIN BI3010 BI3 ON (
        BI3.BI3_CODIGO = BA1.BA1_CODPLA
        AND BI3.D_E_L_E_T_ = ''
    )
    
    -- Joins Opcionais
    LEFT JOIN BQC010 BQC ON (
        BQC.BQC_CODINT = BA1.BA1_CODINT
        AND BQC.BQC_CODEMP = BA1.BA1_CODEMP
        AND BQC.BQC_NUMCON = BA1.BA1_CONEMP
        AND BQC.BQC_VERCON = BA1.BA1_VERCON
        AND BQC.BQC_SUBCON = BA1.BA1_SUBCON
        AND BQC.D_E_L_E_T_ = ''
    )
    
    LEFT JOIN BT5010 BT5 ON (
        BT5.BT5_CODINT = BA1.BA1_CODINT
        AND BT5.BT5_CODIGO = BA1.BA1_CODEMP
        AND BT5.BT5_NUMCON = BA1.BA1_CONEMP
        AND BT5.BT5_VERSAO = BA1.BA1_VERCON
        AND BT5.D_E_L_E_T_ = ''
    )
    
    LEFT JOIN BDK010 BDK ON (
        BDK.BDK_CODINT = BA1.BA1_CODINT
        AND BDK.BDK_FILIAL = BA1.BA1_FILIAL
        AND BDK.BDK_CODEMP = BA1.BA1_CODEMP
        AND BDK.BDK_MATRIC = BA1.BA1_MATRIC
        AND BDK.BDK_TIPREG = BA1.BA1_TIPREG
        AND BDK.D_E_L_E_T_ = ''
        AND DATEDIFF(YEAR, CONVERT(DATE, BA1.BA1_DATNAS), GETDATE()) 
            BETWEEN BDK.BDK_IDAINI AND BDK.BDK_IDAFIN
    )
    
    LEFT JOIN TABLE_ATEND TA ON (
        TA.MATANT = BA1.BA1_CODMAT
    )
    
    LEFT JOIN BQL010 BQL ON (
        BQL.BQL_CODIGO = BA3.BA3_TIPPAG
        AND BQL.D_E_L_E_T_ = ''
    )
    
    -- Filtros
    WHERE BA1.D_E_L_E_T_ = ''
      AND BA1.BA1_CODEMP <> '0003'
      AND BI3.BI3_VTMOD IN ('1', '2')
	  --AND SA3.A3_NOME <> 'VITALMED IMPORTADOS/PJ'
)

-- Query Final com dados do cliente
SELECT 
    DP.BA1_CODMAT AS CodigoMatricula,
    DP.CodigoCliente,
    DP.BA1_NOMUSR AS NomeBeneficiario,
    DP.BA1_CPFUSR AS CPF,
    DP.BA1_DISTRA AS Logradouro,
    DP.BA1_BAIRRO AS Bairro,
    DP.MotivoBloqueio,
    SA1.A1_TEL AS TelefoneCliente,
    DP.DescricaoContrato,
    DP.NomeVendedor,
    DP.DataNascimento,
    DP.QuantidadeAtendimentos,
    DP.TempoContratacaoAno AS TempoContratacao_Anos,
    DP.BA1_TELEFO AS TelefoneBeneficiario,
    DP.IdadeAtual AS Idade,
    DP.UsuarioBloqueio,
    DP.DataInclusao,
    DP.DataBloqueio,
    DP.DescricaoPlano,
    CASE LTRIM(RTRIM(DP.BA1_VTCLAS))
        WHEN '1'   THEN '1 - Classe 1'
        WHEN '2'   THEN '2 - Classe 2'
        WHEN '3'   THEN '3 - Classe 3'
        WHEN '4'   THEN '4 - Classe 4'
        WHEN '5'   THEN '5 - Classe 5'
        WHEN 'IND' THEN 'Individual'
        WHEN 'PRE' THEN 'Premium'
        WHEN 'VER' THEN 'VIP'
        ELSE DP.BA1_VTCLAS
    END AS ClassificacaoPlano,
    DP.ValorMensalidade,
    DP.TipoPagamento,
    DP.UltimaOcorrencia
    
FROM DadosPrincipais DP

LEFT JOIN SA1010 SA1 ON (
    SA1.A1_COD = DP.CodigoCliente
    AND SA1.D_E_L_E_T_ = ''
)



"""

QUERY_PURA_PARA_RELATORIO_FEIRA = """
WITH DatasFechamento AS (
    SELECT
        CAST(? AS DATE) AS DataInicio,
        CAST(? AS DATE) AS DataFim
),
UltimoBloqueio AS (
    SELECT 
        BCA_MATRIC,
        BCA_TIPREG,
        MAX(R_E_C_N_O_) AS UltimoRecno
    FROM BCA070
    CROSS JOIN DatasFechamento D
    WHERE D_E_L_E_T_ = ''
      AND BCA_TIPO = '0'
      AND BCA_DATA BETWEEN D.DataInicio AND D.DataFim
    GROUP BY BCA_MATRIC, BCA_TIPREG
),
DadosPrincipais AS (
    SELECT DISTINCT
        -- Identificação do Beneficiário
        BA1.BA1_CODMAT,
        BA1.BA1_NOMUSR,
        BA1.BA1_CPFUSR,
        BA1.BA1_TELEFO,
        
        -- Endereço
        BA1.BA1_DISTRA,
        BA1.BA1_BAIRRO,
        
        -- Dados do Plano
        CASE 
            WHEN BI3.BI3_VTMOD = 1 THEN 'Tradicional'
            WHEN BI3.BI3_VTMOD = 2 THEN 'Participativo'
            WHEN BI3.BI3_VTMOD = 3 THEN 
                CASE 
                    WHEN BI3.BI3_CODIGO = '0008' THEN 'OMT Vitalfone'
                    ELSE 'Prot. Total'
                END
            WHEN BI3.BI3_VTMOD = 4 THEN 'Área Protegida'
            WHEN BI3.BI3_VTMOD = 5 THEN 'Vitalmed Online'
            WHEN BI3.BI3_VTMOD = 6 THEN 'APH A.P. Resid.'
            WHEN BI3.BI3_VTMOD = 7 THEN 'Vitalmed Online PJ'
            WHEN BI3.BI3_VTMOD = 8 THEN 'APH Cli. Assistidos'
            WHEN BI3.BI3_VTMOD = 9 THEN 'Func V Online'
            ELSE 'Outros'
        END AS DescricaoPlano,

        BA1.BA1_VTCLAS,
        
        -- Dados do Contrato
        CASE 
            WHEN BA3.BA3_COBNIV = '1' THEN BA3.BA3_CODCLI
            WHEN BQC.BQC_COBNIV = '1' THEN BQC.BQC_CODCLI
            WHEN BT5.BT5_COBNIV = '1' THEN BT5.BT5_CODCLI
        END AS CodigoCliente,
        
        CASE 
            WHEN BT5.BT5_COBNIV = '1' THEN BT5.BT5_NOME
            ELSE BQC.BQC_DESCRI
        END AS DescricaoContrato,
        
        -- Dados do Vendedor
        SA3.A3_NOME AS NomeVendedor,
        
        -- Dados do Bloqueio
        BG3.BG3_DESBLO AS MotivoBloqueio,
        BCA.BCA_USUOPE AS UsuarioBloqueio,
        
        -- Datas formatadas com proteção contra valores inválidos
        TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATNAS),''), 5, 0, '-'), 8, 0, '-')) AS DataNascimento,
        TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATINC),''), 5, 0, '-'), 8, 0, '-')) AS DataInclusao,
        TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BCA.BCA_DATA),''), 5, 0, '-'), 8, 0, '-')) AS DataBloqueio,
        
        -- Tempo de Contratação em Anos
		DATEDIFF(MONTH, 
			TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATINC),''), 5, 0, '-'), 8, 0, '-')), 
				COALESCE(
				NULLIF(
					TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATBLO),''), 5, 0, '-'), 8, 0, '-')),
				'1900-01-01'
			),
        GETDATE()
    )
        ) AS TempoContratacao_Anos,
        
        -- Idade no momento do bloqueio
		DATEDIFF(YEAR, 
		    TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATNAS),''), 5, 0, '-'), 8, 0, '-')), 
		    COALESCE(
		        NULLIF(
		            TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATBLO),''), 5, 0, '-'), 8, 0, '-')),
		            '1900-01-01'
		        ),
		        GETDATE()
		    )
		) AS IdadeAtual,
        
        -- Valores
        BDK.BDK_VALOR AS ValorMensalidade,
        BQL.BQL_DESCRI AS TipoPagamento,
        
        -- Contador de Atendimentos
        TA.QTDATE AS QuantidadeAtendimentos,
        
        -- Ocorrência mais recente
        (
            SELECT TOP 1 SE12.E1_DESCOC
            FROM SE1070 SE12
            WHERE SE12.D_E_L_E_T_ = ''
              AND SE12.E1_CLIENTE = BA3.BA3_CODCLI
              AND SE12.E1_CODINT = BA3.BA3_CODINT
              AND SE12.E1_CODEMP = BA3.BA3_CODEMP
              AND SE12.E1_MATRIC = BA3.BA3_MATRIC
              AND SE12.E1_VENCTO < CONVERT(DATETIME, GETDATE(), 103)
            ORDER BY SE12.E1_EMISSAO DESC
        ) AS UltimaOcorrencia

    FROM BA1070 BA1
    
    -- Joins Obrigatórios
    INNER JOIN BA3070 BA3 ON (
        BA3.BA3_CODINT + BA3.BA3_CODEMP + BA3.BA3_MATRIC = 
        BA1.BA1_CODINT + BA1.BA1_CODEMP + BA1.BA1_MATRIC
        AND BA3.D_E_L_E_T_ = ''
    )
    
    INNER JOIN SA3070 SA3 ON (
        SA3.A3_COD = BA1.BA1_CODVEN
        AND SA3.D_E_L_E_T_ = ''
    )
    
    INNER JOIN UltimoBloqueio UB ON (
        UB.BCA_MATRIC = BA1.BA1_CODINT + BA1.BA1_CODEMP + BA1.BA1_MATRIC
        AND UB.BCA_TIPREG = BA1.BA1_TIPREG
    )
    
    INNER JOIN BCA070 BCA ON (
        BCA.R_E_C_N_O_ = UB.UltimoRecno
        AND BCA.D_E_L_E_T_ = ''
    )
    
    INNER JOIN BG3070 BG3 ON (
        BG3.BG3_CODBLO = BCA.BCA_MOTBLO
        AND BG3.D_E_L_E_T_ = ''
    )
    
    INNER JOIN BI3070 BI3 ON (
        BI3.BI3_CODIGO = BA1.BA1_CODPLA
        AND BI3.D_E_L_E_T_ = ''
    )
    
    -- Joins Opcionais
    LEFT JOIN BQC070 BQC ON (
        BQC.BQC_CODINT = BA1.BA1_CODINT
        AND BQC.BQC_CODEMP = BA1.BA1_CODEMP
        AND BQC.BQC_NUMCON = BA1.BA1_CONEMP
        AND BQC.BQC_VERCON = BA1.BA1_VERCON
        AND BQC.BQC_SUBCON = BA1.BA1_SUBCON
        AND BQC.D_E_L_E_T_ = ''
    )
    
    LEFT JOIN BT5070 BT5 ON (
        BT5.BT5_CODINT = BA1.BA1_CODINT
        AND BT5.BT5_CODIGO = BA1.BA1_CODEMP
        AND BT5.BT5_NUMCON = BA1.BA1_CONEMP
        AND BT5.BT5_VERSAO = BA1.BA1_VERCON
        AND BT5.D_E_L_E_T_ = ''
    )
    
    LEFT JOIN BDK070 BDK ON (
        BDK.BDK_CODINT = BA1.BA1_CODINT
        AND BDK.BDK_FILIAL = BA1.BA1_FILIAL
        AND BDK.BDK_CODEMP = BA1.BA1_CODEMP
        AND BDK.BDK_MATRIC = BA1.BA1_MATRIC
        AND BDK.BDK_TIPREG = BA1.BA1_TIPREG
        AND BDK.D_E_L_E_T_ = ''
        AND TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATNAS),''), 5, 0, '-'), 8, 0, '-')) IS NOT NULL
        AND DATEDIFF(YEAR,
                TRY_CONVERT(DATE, STUFF(STUFF(NULLIF(RTRIM(BA1.BA1_DATNAS),''), 5, 0, '-'), 8, 0, '-')),
                GETDATE()
            ) BETWEEN BDK.BDK_IDAINI AND BDK.BDK_IDAFIN
    )
    
    LEFT JOIN TABLE_ATEND_FS TA ON (
        TA.MATANT = BA1.BA1_CODMAT
    )
    
    LEFT JOIN BQL070 BQL ON (
        BQL.BQL_CODIGO = BA3.BA3_TIPPAG
        AND BQL.D_E_L_E_T_ = ''
    )
    
    -- Filtros
    WHERE BA1.D_E_L_E_T_ = ''
      AND BA1.BA1_CODEMP IN ('0001','0002','0003','0006','0007')
      AND BA3.BA3_CODPLA NOT IN ('0083','0094','0095','0096')
)

-- Query Final com dados do cliente
SELECT 
    DP.BA1_CODMAT AS CodigoMatricula,
    DP.CodigoCliente,
    DP.BA1_NOMUSR AS NomeBeneficiario,
    DP.BA1_CPFUSR AS CPF,
    DP.BA1_DISTRA AS Logradouro,
    DP.BA1_BAIRRO AS Bairro,
    DP.MotivoBloqueio,
    SA1.A1_TEL AS TelefoneCliente,
    DP.DescricaoContrato,
    DP.NomeVendedor,
    DP.DataNascimento,
    DP.QuantidadeAtendimentos,
    DP.TempoContratacao_Anos,
    DP.BA1_TELEFO AS TelefoneBeneficiario,
    DP.IdadeAtual AS Idade,
    DP.UsuarioBloqueio,
    DP.DataInclusao,
    DP.DataBloqueio,
    DP.DescricaoPlano,
    CASE LTRIM(RTRIM(DP.BA1_VTCLAS))
        WHEN '1'   THEN '1 *'
        WHEN '2'   THEN '2 *'
        WHEN '3'   THEN '3 *'
        WHEN '4'   THEN '4 *'
        WHEN '5'   THEN '5 *'
        WHEN 'IND' THEN '   '
        WHEN 'PRE' THEN '* P'
        WHEN 'VER' THEN '* V'
        ELSE DP.BA1_VTCLAS
    END AS BA1_VTCLAS,
    DP.ValorMensalidade,
    DP.TipoPagamento,
    DP.UltimaOcorrencia AS DESCOC
    
FROM DadosPrincipais DP

LEFT JOIN SA1070 SA1 ON (
    SA1.A1_COD = DP.CodigoCliente
    AND SA1.D_E_L_E_T_ = ''
)


"""