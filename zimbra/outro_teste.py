from dotenv import load_dotenv
from zimbra.varrer_dados import ZimbraCEPExtractor
import os

def main():
    load_dotenv()

    EMAIL = "adrianasilva@vitalmed.com.br"
    SENHA = os.getenv("SENHA_DRIKA")
    SERVIDOR_IMAP = "imap.emailzimbraonline.com"

    if not SENHA:
        print("✗ ERRO: variável de ambiente SENHA_DRIKA não encontrada")
        return

    extrator = ZimbraCEPExtractor(EMAIL, SENHA, SERVIDOR_IMAP)

    if not extrator.conectar():
        print("✗ Não foi possível conectar ao servidor IMAP.")
        return

    # Lista todas as pastas
    status, pastas = extrator.mail.list()
    if status == "OK" and pastas:
        print("📁 Pastas disponíveis no email:")
        for p in pastas:
            partes = p.decode().split(' "/" ')
            if len(partes) == 2:
                print(f" - {partes[1].strip('\"')}")
    else:
        print("Não foi possível listar pastas ou nenhuma encontrada.")

    extrator.desconectar()
    print("\n✓ Processo finalizado.")

if __name__ == "__main__":
    main()
