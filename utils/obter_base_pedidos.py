from pathlib import Path
import sys

from utils.usuario_atual import obter_usuario_atual

user = obter_usuario_atual()

def obter_base_pedidos():
    """
    Retorna o caminho do arquivo mais recente que contém 'FUP'
    no nome dentro da pasta APP-Acompanhamento de Pedidos.
    """

    pasta = Path(
        fr"C:\Users\{user}\Claro SA\APP-Acompanhamento de Pedidos - Documentos"
    )

    arquivos_fup = [
        arq
        for arq in pasta.iterdir()
        if arq.is_file() and "FUP" in arq.name.upper() and arq.name.lower().endswith(".xlsx")
    ]

    if not arquivos_fup:
        raise FileNotFoundError(
            "Nenhum arquivo contendo 'FUP' foi encontrado na pasta."
        )

    arquivo_mais_recente = max(
        arquivos_fup,
        key=lambda arq: arq.stat().st_mtime
    )

    return str(arquivo_mais_recente)
