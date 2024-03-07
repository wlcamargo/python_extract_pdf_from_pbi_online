from pathlib import Path
import shutil
import os

#Digite o caminho completo do arquivo na pasta Downloads da tua máquina
origem_download = Path('Caminho do arquivo.pdf')

# APONTA PARA O CAMINHO QUE O ARQUIVO SERÁ COPIADO
destino_download = Path('Caminho do Destinho')

# COPIA DATASET DO LOCAL DE ORIGEM NO DESTINO
shutil.copy2(origem_download, destino_download)

# VERIFICA SE EXISTE ARQUIVO E SE EXISTIR APAGA
if os.path.exists(origem_download):
    os.remove(origem_download)

print("Arquivo transferido com sucesso!")
print (f'Local de origem {origem_download} movido para {destino_download}')
