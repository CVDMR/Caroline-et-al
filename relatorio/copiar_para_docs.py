"""Pós-render: copia o site Quarto (_saida) para docs/r, sem tocar no resto de docs/.

O Dropbox às vezes trava pastas recém-criadas; por isso tenta apagar algumas vezes e,
se não conseguir, sobrescreve os arquivos no lugar.
"""
import shutil
import time
from pathlib import Path

aqui = Path(__file__).resolve().parent
origem = aqui / "_saida"
destino = aqui.parent / "docs" / "r"

for tentativa in range(5):
    try:
        if destino.exists():
            shutil.rmtree(destino)
        break
    except PermissionError:
        time.sleep(1.5)
shutil.copytree(origem, destino, dirs_exist_ok=True)
print(f"Site copiado para {destino}")
