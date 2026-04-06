import os
import shutil

# Origem (Downloads\exe)
origem = r"C:\Users\laporciuncula\Downloads\exe"

# Destino inteligente
destino_base = r"C:\Apps\Instaladores"

# Cria pasta destino se não existir
os.makedirs(destino_base, exist_ok=True)

for arquivo in os.listdir(origem):
    caminho_arquivo = os.path.join(origem, arquivo)

    if os.path.isfile(caminho_arquivo) and arquivo.lower().endswith(".exe"):
        destino_arquivo = os.path.join(destino_base, arquivo)

        # Evita sobrescrever arquivo com mesmo nome
        contador = 1
        nome, extensao = os.path.splitext(arquivo)

        while os.path.exists(destino_arquivo):
            novo_nome = f"{nome}_{contador}{extensao}"
            destino_arquivo = os.path.join(destino_base, novo_nome)
            contador += 1

        shutil.move(caminho_arquivo, destino_arquivo)
        print(f"Movido: {arquivo}")

print("Tudo organizado 🚀")