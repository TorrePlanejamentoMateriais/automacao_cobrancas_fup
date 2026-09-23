import getpass
import os


def obter_usuario_atual():
    """
    Retorna o usuário atualmente conectado no Windows.

    Exemplo:
        Z701038
    """
    try:
        usuario = getpass.getuser()

        if usuario:
            return usuario.strip()

        usuario = os.environ.get("USERNAME", "")

        if usuario:
            return usuario.strip()

        return "USUARIO_NAO_IDENTIFICADO"

    except Exception:
        return "USUARIO_NAO_IDENTIFICADO"

