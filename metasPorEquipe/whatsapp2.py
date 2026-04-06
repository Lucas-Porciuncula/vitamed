"""
=============================================================================
ROBÔ DE RELATÓRIO APH AUTOMÁTICO PARA WHATSAPP
=============================================================================

O QUE ESSE CÓDIGO FAZ:
    1. Busca os dados das metas no sistema
    2. Transforma em um relatório bonito
    3. Envia automaticamente pro WhatsApp

COMO USAR:
    - Coloque os nomes dos grupos na lista GRUPOS_WHATSAPP
    - Rode o script
    - Pronto! 🚀
    
ATENÇÃO:
    - O WhatsApp Web precisa estar logado no perfil do Chrome
    - Os nomes dos grupos têm que ser EXATOS
=============================================================================
"""

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from aphTradicional import SistemaAcompanhamentoMetas
import pyperclip
import math


# =============================================================================
# FUNÇÃO AUXILIAR: GERAR BARRA DE PROGRESSO EMOCIONAL
# =============================================================================
def barra_progresso(percentual, tamanho=10):
    try:
        p = float(percentual)
    except:
        return ""

    preenchido = int((p / 100) * tamanho)
    vazio = tamanho - preenchido

    if p >= 100:
        cheio = "🟦"
    elif p >= 70:
        cheio = "🟩"
    else:
        cheio = "🟨"

    barra = cheio * preenchido + "⬜" * vazio
    return f"{barra} {p:.0f}%"



# =============================================================================
# PARTE 1: PEGAR OS DADOS DO SISTEMA
# =============================================================================

def buscar_dados_metas():
    """
    Conecta no sistema e busca os dados de hoje.
    
    Retorna:
        DataFrame com colunas: EQUIPE, Vidas APH, Meta, % da Meta
    """
    print("📥 Buscando dados do sistema...")
    
    sistema = SistemaAcompanhamentoMetas()
    
    try:
        dados = sistema.executar()
        print("✅ Dados coletados com sucesso!")
        return dados
        
    except Exception as erro:
        print(f"❌ Deu ruim ao buscar dados: {erro}")
        raise
        
    finally:
        sistema.fechar_conexao()


# =============================================================================
# PARTE 2: TRANSFORMAR DADOS EM MENSAGEM BONITA
# =============================================================================

def criar_relatorio(dados):
    """
    Pega o DataFrame e transforma numa mensagem formatada pro WhatsApp.

    Parâmetros:
        dados: DataFrame com os dados das equipes

    Retorna:
        String com a mensagem pronta pra enviar
    """

    def fmt_vidas(valor):
        """Converte float para int, trata nan."""
        try:
            if math.isnan(float(valor)):
                return "—"
            return str(int(float(valor)))
        except:
            return "—"

    def fmt_progresso(meta, vidas):
        """Monta o texto de progresso tratando nan e meta zero."""
        try:
            m = float(meta)
            v = float(vidas)
            if math.isnan(m):
                return "sem meta definida"
            saldo = m - v
            if saldo <= 0:
                superou = int(abs(saldo))
                return f"✅ Meta batida! Superou em {superou} vida(s)"
            else:
                return f"faltam {int(saldo)} vidas para alcançar a meta"
        except:
            return "—"

    def fmt_percentual(perc_raw, meta, vidas):
        """Formata percentual, trata inf e nan."""
        try:
            perc_limpo = str(perc_raw).replace('%', '').replace(',', '.').strip()
            percentual = float(perc_limpo)
            if math.isinf(percentual) or math.isnan(percentual):
                m = float(meta)
                v = float(vidas)
                if math.isnan(m) or m == 0:
                    return "—"
                return f"{(v / m * 100):.1f}%"
            return f"{percentual:.1f}%"
        except:
            return "—"

    mensagem = []

    # --- CABEÇALHO COM TOTAIS ---
    total_vidas = int(dados['Vidas APH'].sum())
    total_meta_raw = dados['Meta'].sum()
    barra = barra_progresso((total_vidas / total_meta_raw * 100) if total_meta_raw > 0 else 0)

    if math.isnan(float(total_meta_raw)):
        total_meta_txt = "—"
        percentual_texto = "—"
        progresso_geral = "sem meta definida"
    else:
        total_meta = int(float(total_meta_raw))
        total_meta_txt = str(total_meta)
        if total_meta > 0:
            percentual_texto = f"{(total_vidas / total_meta * 100):.1f}%"
        else:
            percentual_texto = "—"
        saldo_geral = total_meta - total_vidas
        if saldo_geral <= 0:
            progresso_geral = f"✅ Meta batida! Superou em {abs(saldo_geral)} vida(s)"
        else:
            progresso_geral = f"faltam {saldo_geral} vidas para alcançar a meta"

    mensagem.append("📊 *RESULTADO APH – HOJE*")
    mensagem.append("")
    mensagem.append("📌 *TOTAL GERAL*")
    mensagem.append(f"➕ Vidas: {total_vidas}")
    mensagem.append(f"🎯 Meta: {total_meta_txt}")
    mensagem.append(f"🎯 Progresso: {progresso_geral}\n")
    mensagem.append(f"📈 Atingido: {percentual_texto}")
    mensagem.append(barra)
    mensagem.append("")
    mensagem.append("━━━━━━━━━━━━━━━━━━")

   # --- DADOS DE CADA EQUIPE ---
    for _, equipe in dados.iterrows():
        vidas_txt = fmt_vidas(equipe['Vidas APH'])
        meta_txt  = fmt_vidas(equipe['Meta'])
        progresso = fmt_progresso(equipe['Meta'], equipe['Vidas APH'])
        perc_txt  = fmt_percentual(equipe['% da Meta'], equipe['Meta'], equipe['Vidas APH'])

        barra = barra_progresso(
    (equipe['Vidas APH'] / equipe['Meta'] * 100) if equipe['Meta'] > 0 else 0
)

        mensagem.append(
            f"🏷️ *{equipe['EQUIPE']}*\n"
            f"   ➕ Vidas: {vidas_txt}\n"
            f"   🎯 Meta: {meta_txt}\n"
            f"   🎯 Progresso: {progresso}\n"
            f"   📈 Atingido: {perc_txt}\n"
            f"   {barra}\n"
        )

    # --- RODAPÉ MOTIVACIONAL ---
    try:
        if not math.isnan(float(total_meta_raw)) and total_vidas >= int(float(total_meta_raw)):
            mensagem.append("🏆 *Meta geral atingida! Parabéns a todos!*")
        else:
            mensagem.append("🚀 *Bora bater a meta geral!*")
    except:
        mensagem.append("🚀 *Bora bater a meta geral!*")

    return "\n".join(mensagem)


# =============================================================================
# PARTE 3: ENVIAR PRO WHATSAPP
# =============================================================================

def enviar_pro_whatsapp(navegador, nome_do_grupo, mensagem):
    """
    Envia a mensagem para um grupo específico do WhatsApp.
    
    Parâmetros:
        navegador: O Chrome aberto pelo Selenium
        nome_do_grupo: Nome EXATO do grupo (igual tá no WhatsApp)
        mensagem: O texto que vai ser enviado
    """
    
    print(f"📤 Enviando para: {nome_do_grupo}")
    
    espera = WebDriverWait(navegador, 30)
    
    try:
        # PASSO 1: Procurar o grupo
        caixa_pesquisa = espera.until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div/div/div/div/div[3]/div/div[4]/div/div[1]/div/div/div/div/div/div/div[2]/input')
            )
        )
        caixa_pesquisa.click()
        caixa_pesquisa.clear()
        caixa_pesquisa.send_keys(nome_do_grupo)
        sleep(2)
        
        # PASSO 2: Abrir o grupo
        grupo = espera.until(
            EC.element_to_be_clickable(
                (By.XPATH, f'//span[@title="{nome_do_grupo}"]')
            )
        )
        grupo.click()
        sleep(2)
        
        # PASSO 3: Copiar a mensagem (mantém formatação)
        pyperclip.copy(mensagem)
        
        # PASSO 4: Colar e enviar
        caixa_texto = espera.until(
            EC.presence_of_element_located(
                (By.XPATH, '//footer//div[@contenteditable="true"]')
            )
        )
        caixa_texto.click()
        sleep(1)
        
        # Ctrl+V pra colar
        caixa_texto.send_keys(Keys.CONTROL + 'v')
        sleep(1)
        
        # Enter pra enviar
        caixa_texto.send_keys(Keys.ENTER)
        
        print(f"   ✅ Enviado!")
        sleep(1)
        
    except Exception as erro:
        print(f"   ❌ Erro: {erro}")


# =============================================================================
# PARTE 4: CONFIGURAR O CHROME
# =============================================================================

def abrir_whatsapp_web():
    """
    Abre o Chrome no WhatsApp Web usando o perfil salvo.
    
    Retorna:
        O navegador Chrome pronto pra usar
    """
    print("🌐 Abrindo WhatsApp Web...")
    
    # Configura o ChromeDriver
    servico = Service(ChromeDriverManager().install())
    opcoes = webdriver.ChromeOptions()
    
    # Usa o perfil salvo (evita escanear QR code toda vez)
    opcoes.add_argument(
        r"user-data-dir=C:\Users\laporciuncula\AppData\Local\Google\Chrome\SeleniumProfile"
    )
    
    # Abre o navegador
    navegador = webdriver.Chrome(service=servico, options=opcoes)
    navegador.get("https://web.whatsapp.com")
    
    # Espera o WhatsApp carregar completamente
    print("⏳ Aguardando sincronização...")
    espera = WebDriverWait(navegador, 60)
    espera.until(EC.presence_of_element_located((By.ID, 'side')))
    print("✅ WhatsApp pronto!")
    
    return navegador


# =============================================================================
# PARTE 5: EXECUTAR TUDO
# =============================================================================

def main():
    """
    Função principal que executa todo o processo.
    """
    
    # PASSO 1: Buscar dados
    dados = buscar_dados_metas()
    print(dados)
    
    # PASSO 2: Criar relatório
    relatorio = criar_relatorio(dados)
    
    # PASSO 3: Abrir WhatsApp
    navegador = abrir_whatsapp_web()
    
    # PASSO 4: Definir grupos que vão receber
    GRUPOS_WHATSAPP = [
         "Lucas Amorim",
        "VITALMED AÇÕES MKT"
       
        #"Charlene Oliveira",
        #"Vitalmed ADM 1"
        # Adicione mais grupos aqui se precisar:
        # "Equipe APH",
        # "Gestores",
    ]
    
    # PASSO 5: Enviar pra cada grupo
    for grupo in GRUPOS_WHATSAPP:
        enviar_pro_whatsapp(navegador, grupo, relatorio)
        sleep(5)  # Pausa entre envios
    
    # PASSO 6: Fechar navegador
    print("\n🎉 Tudo certo! Relatórios enviados.")
    navegador.quit()


# =============================================================================
# RODAR O CÓDIGO
# =============================================================================

if __name__ == "__main__":
    main()


# =============================================================================
# DICAS PRA QUEM VAI MEXER NISSO DEPOIS
# =============================================================================
"""
💡 PROBLEMAS COMUNS E SOLUÇÕES:

1️⃣ "Não encontrou o grupo"
   → Confere se o nome tá EXATAMENTE igual no WhatsApp
   → Cuidado com maiúsculas e espaços!

2️⃣ "Timeout / Demorou muito"
   → Internet ruim ou WhatsApp não carregou
   → Tenta aumentar o timeout de 30 pra 60 segundos

3️⃣ "Mensagem sem formatação"
   → Instala o pyperclip: pip install pyperclip

4️⃣ "Tem que escanear QR code toda vez"
   → O perfil do Chrome não tá salvando
   → Verifica o caminho em user-data-dir

📝 IDEIAS PRA MELHORAR:

    - Adicionar um arquivo de configuração (tipo Excel ou JSON)
    - Fazer um log mais completo dos envios
    - Agendar pra rodar automaticamente todo dia
    - Adicionar gráficos no relatório
    - Enviar pra contatos individuais também
"""