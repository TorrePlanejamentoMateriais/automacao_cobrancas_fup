"""
Módulo responsável pelo envio de e-mails pelo Microsoft Outlook.

Funcionalidades:
- cria uma mensagem no Outlook;
- adiciona destinatários;
- define assunto e corpo HTML;
- anexa uma planilha;
- incorpora imagens no corpo por meio de Content-ID, CID;
- permite abrir a mensagem para revisão;
- permite realizar o envio automático.
"""

from pathlib import Path
from typing import Optional, Union

import pythoncom
import win32com.client as win32


# ==========================================================
# CONSTANTES MAPI
# ==========================================================

# Content-ID utilizado pela imagem no HTML.
PR_ATTACH_CONTENT_ID = (
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
)

# Tipo MIME do anexo.
PR_ATTACH_MIME_TAG = (
    "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
)

# Indica que o anexo é uma imagem incorporada e não deve,
# preferencialmente, aparecer como anexo convencional.
PR_ATTACHMENT_HIDDEN = (
    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
)

# Localização do conteúdo incorporado.
PR_ATTACH_CONTENT_LOCATION = (
    "http://schemas.microsoft.com/mapi/proptag/0x3713001F"
)


# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================

def validar_arquivo(
    caminho_arquivo: Union[str, Path],
    descricao: str,
) -> Path:
    """
    Valida a existência de um arquivo e retorna seu caminho
    absoluto.

    Parameters
    ----------
    caminho_arquivo:
        Caminho do arquivo que será validado.

    descricao:
        Descrição utilizada na mensagem de erro.

    Returns
    -------
    Path
        Caminho absoluto do arquivo.
    """

    caminho = Path(
        caminho_arquivo
    ).expanduser().resolve()

    if not caminho.exists():
        raise FileNotFoundError(
            f"{descricao} não encontrado: {caminho}"
        )

    if not caminho.is_file():
        raise ValueError(
            f"O caminho informado para {descricao.lower()} "
            f"não é um arquivo: {caminho}"
        )

    return caminho


def incorporar_imagem_cid(
    mensagem,
    caminho_imagem: Union[str, Path],
    content_id: str,
    mime_type: str = "image/png",
):
    """
    Incorpora uma imagem no corpo HTML do e-mail.

    O Content-ID informado nesta função deve ser exatamente
    igual ao CID utilizado na tag HTML.

    Exemplo no HTML:
        cid:imagem_claro

    Exemplo nesta função:
        content_id="imagem_claro"

    Parameters
    ----------
    mensagem:
        Objeto MailItem do Outlook.

    caminho_imagem:
        Caminho completo da imagem.

    content_id:
        Identificador utilizado no atributo src da tag img.

    mime_type:
        Tipo MIME da imagem.

    Returns
    -------
    object
        Objeto Attachment criado pelo Outlook.
    """

    caminho_imagem = validar_arquivo(
        caminho_arquivo=caminho_imagem,
        descricao=(
            f"Imagem incorporada '{content_id}'"
        ),
    )

    content_id = str(
        content_id
    ).strip()

    if not content_id:
        raise ValueError(
            "O Content-ID da imagem não pode ficar vazio."
        )

    # O parâmetro 1 representa olByValue.
    anexo_imagem = mensagem.Attachments.Add(
        str(caminho_imagem),
        1,
        0,
        caminho_imagem.name,
    )

    property_accessor = (
        anexo_imagem.PropertyAccessor
    )

    property_accessor.SetProperty(
        PR_ATTACH_CONTENT_ID,
        content_id,
    )

    property_accessor.SetProperty(
        PR_ATTACH_MIME_TAG,
        mime_type,
    )

    property_accessor.SetProperty(
        PR_ATTACHMENT_HIDDEN,
        True,
    )

    property_accessor.SetProperty(
        PR_ATTACH_CONTENT_LOCATION,
        content_id,
    )

    return anexo_imagem


# ==========================================================
# FUNÇÃO PRINCIPAL
# ==========================================================

def enviar_email_com_anexo(
    destinatarios: str,
    assunto: str,
    corpo: str,
    caminho_anexo: Optional[
        Union[str, Path]
    ] = None,
    caminho_imagem_claro: Optional[
        Union[str, Path]
    ] = None,
    caminho_simbolo_pequeno: Optional[
        Union[str, Path]
    ] = None,
    exibir_antes_de_enviar: bool = True,
):
    """
    Cria e envia um e-mail pelo Outlook Classic.

    Parameters
    ----------
    destinatarios:
        Um ou mais destinatários separados por ponto e
        vírgula.

    assunto:
        Assunto da mensagem.

    corpo:
        Conteúdo HTML completo da mensagem.

    caminho_anexo:
        Caminho da planilha ou de outro arquivo convencional
        que será anexado.

    caminho_imagem_claro:
        Caminho da imagem principal da assinatura.

        Essa imagem será incorporada com o CID:
        imagem_claro

    caminho_simbolo_pequeno:
        Caminho do símbolo pequeno da assinatura.

        Essa imagem será incorporada com o CID:
        simbolo_pequeno

    exibir_antes_de_enviar:
        Quando True, abre a mensagem para revisão.
        Quando False, realiza o envio automático.

    Returns
    -------
    str
        "ABERTO_PARA_REVISAO" quando a mensagem for exibida.

        "ENVIADO" quando a mensagem for enviada
        automaticamente.
    """

    destinatarios = str(
        destinatarios or ""
    ).strip()

    assunto = str(
        assunto or ""
    ).strip()

    corpo = str(
        corpo or ""
    )

    if not destinatarios:
        raise ValueError(
            "Nenhum destinatário foi informado."
        )

    if not assunto:
        raise ValueError(
            "O assunto do e-mail não foi informado."
        )

    if not corpo:
        raise ValueError(
            "O corpo do e-mail não foi informado."
        )

    # Inicializa o COM na thread atual.
    pythoncom.CoInitialize()

    try:
        outlook = win32.Dispatch(
            "Outlook.Application"
        )

        mensagem = outlook.CreateItem(0)

        mensagem.To = destinatarios
        mensagem.Subject = assunto

        # Define o formato como HTML.
        mensagem.BodyFormat = 2

        # --------------------------------------------------
        # ANEXO PRINCIPAL
        # --------------------------------------------------

        if caminho_anexo:

            caminho_anexo_validado = validar_arquivo(
                caminho_arquivo=caminho_anexo,
                descricao="Anexo",
            )

            mensagem.Attachments.Add(
                str(caminho_anexo_validado)
            )

        # --------------------------------------------------
        # IMAGEM PRINCIPAL DA ASSINATURA
        # --------------------------------------------------

        if caminho_imagem_claro:

            incorporar_imagem_cid(
                mensagem=mensagem,
                caminho_imagem=(
                    caminho_imagem_claro
                ),
                content_id="imagem_claro",
                mime_type="image/png",
            )

        # --------------------------------------------------
        # SÍMBOLO PEQUENO DA ASSINATURA
        # --------------------------------------------------

        if caminho_simbolo_pequeno:

            incorporar_imagem_cid(
                mensagem=mensagem,
                caminho_imagem=(
                    caminho_simbolo_pequeno
                ),
                content_id="simbolo_pequeno",
                mime_type="image/png",
            )

        # O corpo HTML deve ser definido depois de adicionar
        # os anexos incorporados.
        mensagem.HTMLBody = corpo

        # Salva a mensagem para que o Outlook consolide os
        # anexos e as propriedades Content-ID.
        mensagem.Save()

        if exibir_antes_de_enviar:

            mensagem.Display()

            return "ABERTO_PARA_REVISAO"

        mensagem.Send()

        return "ENVIADO"

    except Exception as erro:

        raise RuntimeError(
            "Não foi possível criar ou enviar o e-mail "
            f"pelo Outlook: {erro}"
        ) from erro

    finally:

        pythoncom.CoUninitialize()