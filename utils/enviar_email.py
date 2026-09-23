"""
Módulo responsável pelo envio de e-mails pelo Microsoft Outlook Classic.

Funcionalidades:
- cria uma mensagem no Outlook;
- adiciona um ou mais destinatários;
- define assunto e corpo HTML;
- anexa uma planilha;
- incorpora imagens no corpo por meio de Content-ID, CID;
- permite abrir a mensagem para revisão;
- permite realizar o envio automático;
- utiliza a conta padrão configurada no Outlook do usuário.

Requisitos:
- Windows;
- Microsoft Outlook Classic instalado;
- conta configurada e autenticada no Outlook;
- pacote pywin32 instalado.
"""

from pathlib import Path
from typing import Optional, Union

import pythoncom
import win32com.client as win32


# ==========================================================
# CONSTANTES DO OUTLOOK
# ==========================================================

# Tipo de item do Outlook:
# 0 representa olMailItem.
OL_MAIL_ITEM = 0

# Formato do corpo:
# 2 representa olFormatHTML.
OL_FORMAT_HTML = 2

# Tipo de anexo:
# 1 representa olByValue.
OL_BY_VALUE = 1


# ==========================================================
# CONSTANTES MAPI
# ==========================================================

# Content-ID utilizado pela imagem incorporada no HTML.
PR_ATTACH_CONTENT_ID = (
    "http://schemas.microsoft.com/mapi/"
    "proptag/0x3712001F"
)

# Tipo MIME do anexo.
PR_ATTACH_MIME_TAG = (
    "http://schemas.microsoft.com/mapi/"
    "proptag/0x370E001F"
)

# Indica que o anexo é uma imagem incorporada.
PR_ATTACHMENT_HIDDEN = (
    "http://schemas.microsoft.com/mapi/"
    "proptag/0x7FFE000B"
)

# Localização do conteúdo incorporado.
PR_ATTACH_CONTENT_LOCATION = (
    "http://schemas.microsoft.com/mapi/"
    "proptag/0x3713001F"
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
        Caminho absoluto do arquivo validado.
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


def normalizar_destinatarios(
    destinatarios: str,
) -> str:
    """
    Normaliza os destinatários para o formato utilizado
    pelo Outlook.

    Aceita destinatários separados por:
    - ponto e vírgula;
    - vírgula;
    - quebra de linha.

    Parameters
    ----------
    destinatarios:
        Um ou mais endereços de e-mail.

    Returns
    -------
    str
        Endereços separados por ponto e vírgula.
    """

    texto = str(
        destinatarios or ""
    ).strip()

    if not texto:
        raise ValueError(
            "Nenhum destinatário foi informado."
        )

    texto = texto.replace(
        "\r",
        ";",
    )

    texto = texto.replace(
        "\n",
        ";",
    )

    texto = texto.replace(
        ",",
        ";",
    )

    lista_destinatarios = []

    for destinatario in texto.split(";"):
        destinatario = destinatario.strip()

        if not destinatario:
            continue

        if "@" not in destinatario:
            raise ValueError(
                "Endereço de e-mail inválido: "
                f"{destinatario}"
            )

        destinatario_normalizado = (
            destinatario.lower()
        )

        destinatarios_existentes = [
            item.lower()
            for item in lista_destinatarios
        ]

        if (
            destinatario_normalizado
            not in destinatarios_existentes
        ):
            lista_destinatarios.append(
                destinatario
            )

    if not lista_destinatarios:
        raise ValueError(
            "Nenhum destinatário válido foi informado."
        )

    return ";".join(
        lista_destinatarios
    )


def obter_outlook():
    """
    Obtém uma instância do Microsoft Outlook Classic.

    Primeiro tenta reutilizar uma instância aberta.
    Caso não encontre, inicia uma nova instância.

    Returns
    -------
    object
        Instância COM do Outlook.
    """

    try:
        outlook = win32.GetActiveObject(
            "Outlook.Application"
        )

    except Exception:
        outlook = win32.Dispatch(
            "Outlook.Application"
        )

    return outlook


def incorporar_imagem_cid(
    mensagem,
    caminho_imagem: Union[str, Path],
    content_id: str,
    mime_type: str = "image/png",
):
    """
    Incorpora uma imagem no corpo HTML do e-mail.

    O Content-ID informado nesta função deve ser exatamente
    igual ao CID utilizado no HTML.

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

    caminho_imagem_validado = validar_arquivo(
        caminho_arquivo=caminho_imagem,
        descricao=(
            f"Imagem incorporada '{content_id}'"
        ),
    )

    content_id = str(
        content_id or ""
    ).strip()

    if not content_id:
        raise ValueError(
            "O Content-ID da imagem não pode ficar vazio."
        )

    mime_type = str(
        mime_type or "image/png"
    ).strip()

    # O parâmetro OL_BY_VALUE adiciona o arquivo
    # diretamente ao item do Outlook.
    anexo_imagem = mensagem.Attachments.Add(
        str(caminho_imagem_validado),
        OL_BY_VALUE,
        0,
        caminho_imagem_validado.name,
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


def selecionar_conta_outlook(
    outlook,
    endereco_remetente: Optional[str],
):
    """
    Procura uma conta específica entre as contas configuradas
    no Outlook.

    Essa função é utilizada somente quando o parâmetro
    endereco_remetente é informado.

    Parameters
    ----------
    outlook:
        Instância do Outlook.

    endereco_remetente:
        Endereço da conta que deverá realizar o envio.

    Returns
    -------
    object ou None
        Conta localizada no Outlook.

    Raises
    ------
    RuntimeError
        Quando o endereço foi informado, mas nenhuma conta
        correspondente foi encontrada.
    """

    endereco_remetente = str(
        endereco_remetente or ""
    ).strip()

    if not endereco_remetente:
        return None

    endereco_procurado = (
        endereco_remetente.lower()
    )

    contas_encontradas = []

    sessao = outlook.Session

    for indice in range(
        1,
        sessao.Accounts.Count + 1,
    ):
        conta = sessao.Accounts.Item(
            indice
        )

        endereco_conta = str(
            getattr(
                conta,
                "SmtpAddress",
                "",
            )
            or ""
        ).strip()

        if endereco_conta:
            contas_encontradas.append(
                endereco_conta
            )

        if (
            endereco_conta.lower()
            == endereco_procurado
        ):
            return conta

    texto_contas = (
        ", ".join(contas_encontradas)
        if contas_encontradas
        else "nenhuma conta SMTP identificada"
    )

    raise RuntimeError(
        "A conta solicitada para envio não foi encontrada "
        f"no Outlook: {endereco_remetente}. "
        f"Contas encontradas: {texto_contas}."
    )


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
    endereco_remetente: Optional[str] = None,
) -> str:
    """
    Cria e envia um e-mail pelo Microsoft Outlook Classic.

    Parameters
    ----------
    destinatarios:
        Um ou mais destinatários separados por ponto e
        vírgula, vírgula ou quebra de linha.

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
        Quando True, abre a mensagem no Outlook para revisão.

        Quando False, realiza o envio automático.

    endereco_remetente:
        Parâmetro opcional para selecionar uma conta específica
        configurada no Outlook.

        Quando não informado, o Outlook utiliza a conta padrão.

    Returns
    -------
    str
        "ABERTO_PARA_REVISAO" quando a mensagem for exibida.

        "ENVIADO" quando a mensagem for enviada
        automaticamente.
    """

    destinatarios_normalizados = (
        normalizar_destinatarios(
            destinatarios
        )
    )

    assunto = str(
        assunto or ""
    ).strip()

    corpo = str(
        corpo or ""
    ).strip()

    if not assunto:
        raise ValueError(
            "O assunto do e-mail não foi informado."
        )

    if not corpo:
        raise ValueError(
            "O corpo do e-mail não foi informado."
        )

    # Faz as validações antes de abrir o Outlook.
    caminho_anexo_validado = None
    imagem_claro_validada = None
    simbolo_pequeno_validado = None

    if caminho_anexo:
        caminho_anexo_validado = validar_arquivo(
            caminho_arquivo=caminho_anexo,
            descricao="Anexo principal",
        )

    if caminho_imagem_claro:
        imagem_claro_validada = validar_arquivo(
            caminho_arquivo=caminho_imagem_claro,
            descricao="Imagem principal da assinatura",
        )

    if caminho_simbolo_pequeno:
        simbolo_pequeno_validado = validar_arquivo(
            caminho_arquivo=caminho_simbolo_pequeno,
            descricao="Símbolo pequeno da assinatura",
        )

    # Inicializa o COM na thread atual.
    # Isso é necessário quando a função é chamada por uma
    # thread criada pelo Tkinter.
    pythoncom.CoInitialize()

    mensagem = None
    outlook = None

    try:
        outlook = obter_outlook()

        mensagem = outlook.CreateItem(
            OL_MAIL_ITEM
        )

        # --------------------------------------------------
        # CONTA DO REMETENTE
        # --------------------------------------------------

        conta_remetente = selecionar_conta_outlook(
            outlook=outlook,
            endereco_remetente=endereco_remetente,
        )

        if conta_remetente is not None:
            mensagem.SendUsingAccount = (
                conta_remetente
            )

        # --------------------------------------------------
        # DESTINATÁRIOS, ASSUNTO E FORMATO
        # --------------------------------------------------

        mensagem.To = destinatarios_normalizados
        mensagem.Subject = assunto
        mensagem.BodyFormat = OL_FORMAT_HTML

        # --------------------------------------------------
        # ANEXO PRINCIPAL
        # --------------------------------------------------

        if caminho_anexo_validado is not None:
            mensagem.Attachments.Add(
                str(caminho_anexo_validado),
                OL_BY_VALUE,
                1,
                caminho_anexo_validado.name,
            )

        # --------------------------------------------------
        # IMAGEM PRINCIPAL DA ASSINATURA
        # --------------------------------------------------

        if imagem_claro_validada is not None:
            incorporar_imagem_cid(
                mensagem=mensagem,
                caminho_imagem=imagem_claro_validada,
                content_id="imagem_claro",
                mime_type="image/png",
            )

        # --------------------------------------------------
        # SÍMBOLO PEQUENO DA ASSINATURA
        # --------------------------------------------------

        if simbolo_pequeno_validado is not None:
            incorporar_imagem_cid(
                mensagem=mensagem,
                caminho_imagem=simbolo_pequeno_validado,
                content_id="simbolo_pequeno",
                mime_type="image/png",
            )

        # O corpo HTML deve ser definido depois da inclusão
        # das imagens incorporadas.
        mensagem.HTMLBody = corpo

        # Salva o item para que o Outlook consolide os anexos
        # e as propriedades Content-ID.
        mensagem.Save()

        # --------------------------------------------------
        # ABRIR PARA REVISÃO
        # --------------------------------------------------

        if exibir_antes_de_enviar:
            mensagem.Display()

            return "ABERTO_PARA_REVISAO"

        # --------------------------------------------------
        # ENVIO AUTOMÁTICO
        # --------------------------------------------------

        mensagem.Send()

        return "ENVIADO"

    except Exception as erro:
        raise RuntimeError(
            "Não foi possível criar ou enviar o e-mail "
            f"pelo Outlook Classic: {erro}"
        ) from erro

    finally:
        # Evita encerrar o Outlook. Apenas libera as referências
        # COM utilizadas pela função.
        mensagem = None
        outlook = None

        pythoncom.CoUninitialize()