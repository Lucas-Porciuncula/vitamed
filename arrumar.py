import os
import shutil

# Caminho da pasta
pasta = r"C:\Users\laporciuncula\Documents"

# Percorrer todos os arquivos
for arquivo in os.listdir(pasta):
    caminho_arquivo = os.path.join(pasta, arquivo)

    # Verifica se é arquivo (ignora pastas)
    if os.path.isfile(caminho_arquivo):
        # Pega a extensão
        extensao = arquivo.split('.')[-1].lower()

        # Se não tiver extensão
        if '.' not in arquivo:
            extensao = "sem_extensao"

        # Cria a pasta da extensão
        pasta_destino = os.path.join(pasta, extensao)
        os.makedirs(pasta_destino, exist_ok=True)

        # Move o arquivo
        shutil.move(caminho_arquivo, os.path.join(pasta_destino, arquivo))

print("Organização concluída 🚀")