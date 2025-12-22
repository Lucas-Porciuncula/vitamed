from dotenv import load_dotenv
from varrer_dados import ZimbraCEPExtractor
import os
from datetime import datetime, timedelta


def main():
    load_dotenv()

    # =========================
    # CONFIGURAÇÕES
    # =========================
    EMAIL =  'laporciuncula@vitalmed.com.br'#"adrianasilva@vitalmed.com.br"
    SENHA = os.getenv  ("SENHA_EMAIL") #("SENHA_DRIKA")
    SERVIDOR_IMAP = "imap.emailzimbraonline.com"

    if not SENHA:
        print("✗ ERRO: variável de ambiente SENHA_EMAIL não encontrada")
        return

    # =========================
    # INICIALIZA EXTRATOR
    # =========================
    extrator = ZimbraCEPExtractor(EMAIL, SENHA, SERVIDOR_IMAP)

    # =========================
    # CONEXÃO
    # =========================
    if not extrator.conectar():
        print("✗ Não foi possível conectar ao servidor IMAP.")
        return

    # =========================
    # MENU DE BUSCA
    # =========================
    os.system('cls' if os.name == 'nt' else 'clear')
    print("\n" + "="*60)
    print("EXTRATOR DE CEPs - ZIMBRA EMAIL")
    print("="*60)
    print("\n--- OPÇÕES DE BUSCA ---")
    print("1. Últimos 100 emails (rápido - recomendado para teste)")
    print("2. Últimos 500 emails")
    print("3. Últimos 1000 emails")
    print("4. Todos os emails dos últimos 6 meses")
    print("5. Período personalizado")
    print("6. BUSCA TOTAL - Todos os emails (pode demorar muito!)")
    
    escolha = input("\nEscolha uma opção (1-6): ").strip()

    limite = None
    data_inicio = None
    buscar_todos = False

    if escolha == "1":
        limite = 100

    elif escolha == "2":
        limite = 500

    elif escolha == "3":
        limite = 1000

    elif escolha == "4":
        seis_meses_atras = datetime.now() - timedelta(days=8*30)
        data_inicio = seis_meses_atras.strftime("%d-%b-%Y")

    elif escolha == "5":
        data_input = input("Informe a data inicial (ex: 01-Jan-2025): ").strip()

        if data_input:
            data_inicio = data_input

        limite_input = input("Informe limite de emails (ou deixe vazio para todos): ").strip()
        limite = int(limite_input) if limite_input.isdigit() else None

    elif escolha == "6":
        confirma = input("⚠️  Isso vai processar TODOS os emails. Confirma? (s/n): ").strip().lower()
        
        if confirma == 's':
            buscar_todos = True
            limite_input = input("Quer definir um limite de segurança? (deixe vazio para sem limite): ").strip()
            limite = int(limite_input) if limite_input.isdigit() else None
        else:
            print("Operação cancelada. Usando padrão de 100 emails.")
            limite = 100
    else:
        print("Opção inválida. Usando padrão de 100 emails.")
        limite = 100

    # =========================
    # BUSCA EMAILS COM CEP
    # =========================
    print("\n" + "="*60)
    print("INICIANDO BUSCA")
    print("="*60)
    
    if buscar_todos:
        print("⚠️  MODO: Busca Total (processa todos os emails)")
    else:
        print("🔍 MODO: Busca Otimizada (apenas emails com palavra 'CEP')")
    
    if limite:
        print(f"📊 LIMITE: {limite} emails")
    if data_inicio:
        print(f"📅 DATA INICIAL: {data_inicio}")
    
    print("\nProcessando...\n")

    resultados = extrator.buscar_emails_com_cep(
        limite=limite,
        data_inicio=data_inicio,
        pastas=["INBOX"]
    )


    # =========================
    # RESUMO E EXPORTAÇÃO
    # =========================
    print("\n" + "="*60)
    print("RESUMO FINAL")
    print("="*60)
    print(f"✓ Emails com CEP encontrados: {len(resultados)}")

    if resultados:
        total_ceps = sum(r["quantidade"] for r in resultados)
        print(f"✓ Total de CEPs únicos extraídos: {total_ceps}")
        
        # Mostra preview dos primeiros resultados
        print("\n--- PREVIEW (primeiros 5 resultados) ---")
        for i, resultado in enumerate(resultados[:5], 1):
            print(f"\n{i}. {resultado['assunto'][:60]}...")
            print(f"   Remetente: {resultado['remetente'][:50]}")
            print(f"   CEPs: {resultado['ceps']}")
        
        if len(resultados) > 5:
            print(f"\n... e mais {len(resultados) - 5} emails com CEPs")
        
        # Exporta
        print("\n" + "="*60)
        formato = input("Exportar como Excel (e) ou CSV (c)? [e/c]: ").strip().lower()
        formato_export = 'excel' if formato != 'c' else 'csv'
        
        arquivo = extrator.exportar_resultados(resultados, formato=formato_export)
        
        if arquivo:
            print(f"\n🎉 Arquivo criado: {arquivo}")
    else:
        print("\n⚠️  Nenhum CEP válido encontrado.")
        print("\n💡 DICAS:")
        print("   • Tente aumentar o limite de emails")
        print("   • Use a opção 6 (Busca Total) para garantir")
        print("   • Verifique se os emails realmente contêm CEPs no formato XXXXX-XXX")

    extrator.desconectar()
    print("\n✓ Processo finalizado.\n")


if __name__ == "__main__":
    main()