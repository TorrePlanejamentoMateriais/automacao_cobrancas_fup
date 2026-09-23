from datetime import datetime
from pathlib import Path

import html
import random
import string
import traceback

import pandas as pd

from utils.destaca_coluna_excel import destacar_coluna_excel
from utils.enviar_email import enviar_email_com_anexo


# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

COLUNA_PREVISAO_FORNECEDOR = (
    "Previsão Atual(Preencha Aqui a nova data)"
)


# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================

def gerar_id_email(tamanho=8):
    """
    Gera um novo identificador para a tentativa de reenvio.
    """

    caracteres = (
        string.ascii_uppercase
        + string.digits
    )

    return "".join(
        random.choices(
            caracteres,
            k=tamanho,
        )
    )


def calcular_data_retorno(
    data_envio,
    dias_para_retorno=2,
):
    """
    Calcula a data limite de retorno.

    Soma dias corridos e, caso a data final caia no sábado
    ou domingo, transfere o prazo para segunda-feira.
    """

    data_retorno = (
        pd.Timestamp(data_envio).normalize()
        + pd.Timedelta(days=dias_para_retorno)
    )

    if data_retorno.weekday() == 5:
        data_retorno += pd.Timedelta(days=2)

    elif data_retorno.weekday() == 6:
        data_retorno += pd.Timedelta(days=1)

    return data_retorno.to_pydatetime()


def obter_primeiro_valor(
    dataframe,
    coluna,
    valor_padrao="",
):
    """
    Obtém o primeiro valor preenchido de uma coluna.
    """

    if coluna not in dataframe.columns:
        return valor_padrao

    valores = (
        dataframe[coluna]
        .dropna()
        .astype(str)
        .str.strip()
    )

    valores = valores[
        valores.ne("")
        & valores.str.casefold().ne("nan")
        & valores.str.casefold().ne("none")
    ]

    if valores.empty:
        return valor_padrao

    return valores.iloc[0]


def obter_emails_fornecedor(grupo):
    """
    Obtém os e-mails únicos existentes no grupo.

    Aceita e-mails separados por ponto e vírgula e remove
    duplicações sem alterar a ordem original.
    """

    if "EmailFornecedor" not in grupo.columns:
        return ""

    emails_unicos = []
    emails_normalizados = set()

    valores = (
        grupo["EmailFornecedor"]
        .dropna()
        .astype(str)
        .str.strip()
    )

    for valor in valores:

        if (
            not valor
            or valor.casefold() in {"nan", "none", "sem e-mail"}
        ):
            continue

        for email_individual in valor.split(";"):

            email_individual = email_individual.strip()

            if not email_individual:
                continue

            email_normalizado = email_individual.casefold()

            if email_normalizado in emails_normalizados:
                continue

            emails_normalizados.add(
                email_normalizado
            )

            emails_unicos.append(
                email_individual
            )

    return ";".join(
        emails_unicos
    )


def limpar_nome_arquivo(nome):
    """
    Remove caracteres inválidos de nomes de arquivos
    no Windows.
    """

    caracteres_invalidos = '<>:"/\\|?*'

    nome_limpo = "".join(
        "_"
        if caractere in caracteres_invalidos
        else caractere
        for caractere in str(nome)
    )

    nome_limpo = " ".join(
        nome_limpo.split()
    )

    return nome_limpo.strip().rstrip(". ")


# ==========================================================
# CORPO DO E-MAIL
# ==========================================================

def montar_corpo_reenvio(
    fornecedor,
    data_retorno_formatada,
):
    """
    Monta o corpo HTML utilizado no reenvio.
    """

    fornecedor_html = html.escape(
        str(fornecedor)
    )

    return f"""
<html>
<head>
    <meta charset="UTF-8">
</head>

<body style="
    margin: 0;
    padding: 0;
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    color: #222222;
">

    <p>
        <strong>Olá, tudo bem?</strong>
    </p>

    <p>
        Estamos reenviando a carteira do fornecedor
        <strong>{fornecedor_html}</strong> para atualização
        da previsão de entrega.
    </p>

    <p>
        Por gentileza, informar a justificativa dos pedidos
        em atraso e preencher a nova previsão de entrega na
        coluna destacada da planilha anexa.
    </p>

    <p>
        Para os itens com entrega parcial, solicitamos o envio
        do cronograma detalhado, com as quantidades pendentes
        e as respectivas datas previstas de entrega.
    </p>

    <p>
        <strong>
            Pedimos o retorno até {data_retorno_formatada}.
        </strong>
    </p>

    <p>Obrigada!</p>

    <br>

    <table
        role="presentation"
        cellpadding="0"
        cellspacing="0"
        border="0"
        style="
            border-collapse: collapse;
            font-family: Calibri, Arial, sans-serif;
            color: #222222;
        "
    >
        <tr>
            <td
                style="
                    vertical-align: middle;
                    padding-right: 14px;
                "
            >
                cid:imagem_claro
            </td>

            <td
                style="
                    vertical-align: middle;
                    font-size: 10pt;
                    line-height: 1.25;
                "
            >
                <div
                    style="
                        font-size: 11pt;
                        font-weight: bold;
                        white-space: nowrap;
                    "
                >
                    GICELIA SANTOS DE OLIVEIRA

                    cid:simbolo_pequeno
                </div>

                <div
                    style="
                        margin-top: 2px;
                        font-size: 8pt;
                        font-weight: bold;
                    "
                >
                    ÁREA CORPORATIVA
                </div>

                <div>
                    Logística | Planejamento de Materiais
                </div>

                <div>
                    55 11 2612-2542
                </div>

                <div>
                    <a
                        href="mailto:gicelia.oliveira@claro.com.br"
                        style="
                            color: #7A007A;
                            text-decoration: underline;
                        "
                    >
                        gicelia.oliveira@claro.com.br
                    </a>
                </div>

                <div>
                    Claro Brasil
                </div>

                <div>
                    <a
                        href="https://www.claro.com.br"
                        style="
                            color: #7A007A;
                            text-decoration: underline;
                        "
                    >
                        www.claro.com.br
                    </a>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>
"""


# ==========================================================
# LOG DE REENVIO
# ==========================================================

def adicionar_log_reenvio(
    lista_logs,
    fornecedor,
    email_fornecedor,
    email_destino,
    id_email,
    assunto,
    status_envio,
    motivo,
    quantidade_itens,
    caminho_arquivo,
    data_tentativa,
    data_envio,
    data_retorno,
):
    """
    Adiciona o resultado do reenvio à lista de logs.

    A estrutura é compatível com salvar_log_excel()
    existente em utils/logs.py.
    """

    if status_envio == "ENVIADO":
        status_retorno = "AGUARDANDO RETORNO"
    else:
        status_retorno = ""

    lista_logs.append(
        {
            "ID_Email": id_email,
            "Fornecedor": str(fornecedor),
            "EmailFornecedor": email_fornecedor,
            "EmailDestino": email_destino,
            "Assunto": assunto,
            "QuantidadeItens": quantidade_itens,
            "DataTentativa": (
                pd.Timestamp(data_tentativa)
                if data_tentativa is not None
                else pd.NaT
            ),
            "DataEnvio": (
                pd.Timestamp(data_envio)
                if data_envio is not None
                else pd.NaT
            ),
            "DataLimiteRetorno": (
                pd.Timestamp(data_retorno)
                if data_retorno is not None
                else pd.NaT
            ),
            "StatusEnvio": status_envio,
            "StatusRetorno": status_retorno,
            "DataResposta": pd.NaT,
            "ResponsavelAcompanhamento": "",
            "MotivoObservacao": motivo,
            "ArquivoEnviado": (
                str(caminho_arquivo)
                if caminho_arquivo
                else ""
            ),
        }
    )


# ==========================================================
# BASE DE REENVIO
# ==========================================================

def salvar_base_reenvio(
    dataframe,
    arquivo_reenvio,
):
    """
    Salva a base de reenvio preservando todas as linhas
    de detalhe dos pedidos.
    """

    arquivo_reenvio = Path(
        arquivo_reenvio
    )

    arquivo_reenvio.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    dataframe = dataframe.reset_index(
        drop=True
    )

    dataframe.to_excel(
        arquivo_reenvio,
        index=False,
        engine="openpyxl",
    )


# ==========================================================
# REENVIO DE E-MAILS
# ==========================================================

def reenviar_emails_pendentes(
    arquivo_reenvio,
    pasta_fornecedores,
    imagem_claro,
    simbolo_pequeno,
    logs_execucao,
    modo_teste=True,
    email_teste="",
    exibir_antes_de_enviar=False,
    dias_para_retorno=2,
    evento_parar=None,
    callback_log=None,
):
    """
    Processa a fila de e-mails não enviados.

    Funcionamento:
    1. lê a base_reenvio.xlsx;
    2. seleciona registros com STATUS_ENVIO = NÃO ENVIADO;
    3. agrupa as linhas pelo ID_EMAIL;
    4. preserva todos os detalhes dos pedidos;
    5. recria o anexo do fornecedor;
    6. tenta enviar novamente;
    7. remove da fila somente os enviados com sucesso;
    8. mantém as falhas para uma nova tentativa;
    9. adiciona o resultado em logs_execucao;
    10. permite interrupção segura por threading.Event.

    Retorna:
        Dicionário com o resumo da execução.
    """

    # ======================================================
    # LOG LOCAL
    # ======================================================

    def log(mensagem=""):
        """
        Exibe a mensagem no terminal e, quando informado,
        envia também para o callback da interface.
        """

        mensagem = str(mensagem)

        print(
            mensagem,
            flush=True,
        )

        if callback_log is not None:
            callback_log(
                mensagem
            )

    def parada_solicitada():
        """
        Retorna True quando o botão Parar Execução
        tiver solicitado a interrupção da rotina.
        """

        return bool(
            evento_parar is not None
            and evento_parar.is_set()
        )

    # ======================================================
    # CAMINHOS
    # ======================================================

    arquivo_reenvio = Path(
        arquivo_reenvio
    )

    pasta_fornecedores = Path(
        pasta_fornecedores
    )

    imagem_claro = Path(
        imagem_claro
    )

    simbolo_pequeno = Path(
        simbolo_pequeno
    )

    resumo = {
        "processados": 0,
        "enviados": 0,
        "nao_enviados": 0,
        "abertos_revisao": 0,
        "pendencias_restantes": 0,
        "interrompido": False,
    }

    log("")
    log("=" * 80)
    log("PROCESSAMENTO DA FILA DE REENVIO")
    log("=" * 80)

    # ======================================================
    # VALIDAÇÕES
    # ======================================================

    if parada_solicitada():
        resumo["interrompido"] = True

        log(
            "Reenvio cancelado antes do início."
        )

        return resumo

    if not arquivo_reenvio.exists():
        log(
            "A base de reenvio não foi encontrada: "
            f"{arquivo_reenvio}"
        )

        return resumo

    if not imagem_claro.exists():
        raise FileNotFoundError(
            f"Imagem não encontrada: {imagem_claro}"
        )

    if not simbolo_pequeno.exists():
        raise FileNotFoundError(
            f"Imagem não encontrada: {simbolo_pequeno}"
        )

    if modo_teste and not str(email_teste).strip():
        raise ValueError(
            "O modo de teste está ativado, mas o endereço "
            "de teste não foi informado."
        )

    pasta_fornecedores.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ======================================================
    # LEITURA DA BASE
    # ======================================================

    try:
        df_reenvio = pd.read_excel(
            arquivo_reenvio,
            engine="openpyxl",
            dtype={
                "ID_EMAIL": "string",
                "STATUS_ENVIO": "string",
                "EMAIL_DESTINO": "string",
            },
        )

    except PermissionError as erro:
        raise PermissionError(
            "Não foi possível abrir a base de reenvio. "
            "Verifique se o arquivo está aberto no Excel."
        ) from erro

    except Exception as erro:
        raise RuntimeError(
            f"Erro ao ler a base de reenvio: {erro}"
        ) from erro

    if df_reenvio.empty:
        log(
            "A base de reenvio está vazia."
        )

        return resumo

    # ======================================================
    # VERIFICAÇÃO DAS COLUNAS
    # ======================================================

    colunas_obrigatorias = [
        "ID_EMAIL",
        "STATUS_ENVIO",
        "Nº CONTA DO FORNECEDOR",
    ]

    colunas_ausentes = [
        coluna
        for coluna in colunas_obrigatorias
        if coluna not in df_reenvio.columns
    ]

    if colunas_ausentes:
        raise KeyError(
            "Colunas obrigatórias ausentes na base "
            f"de reenvio: {colunas_ausentes}. "
            "Colunas disponíveis: "
            f"{df_reenvio.columns.tolist()}"
        )

    colunas_controle_padrao = {
        "EMAIL_DESTINO": "",
        "DATA_ENVIO": pd.NaT,
        "DATA_TENTATIVA_REENVIO": pd.NaT,
        "MOTIVO_ULTIMA_TENTATIVA": "",
    }

    for coluna, valor_padrao in colunas_controle_padrao.items():

        if coluna not in df_reenvio.columns:
            df_reenvio[coluna] = valor_padrao

    # ======================================================
    # NORMALIZAÇÃO
    # ======================================================

    df_reenvio["ID_EMAIL"] = (
        df_reenvio["ID_EMAIL"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df_reenvio["STATUS_ENVIO"] = (
        df_reenvio["STATUS_ENVIO"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df_reenvio["EMAIL_DESTINO"] = (
        df_reenvio["EMAIL_DESTINO"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    mascara_pendencias = (
        df_reenvio["STATUS_ENVIO"].eq(
            "NÃO ENVIADO"
        )
        & df_reenvio["ID_EMAIL"].ne("")
    )

    df_pendencias = df_reenvio[
        mascara_pendencias
    ].copy()

    if df_pendencias.empty:
        log(
            "Não existem e-mails com status "
            "'NÃO ENVIADO' para reprocessar."
        )

        return resumo

    quantidade_tentativas = (
        df_pendencias["ID_EMAIL"].nunique()
    )

    log(
        "Tentativas pendentes encontradas: "
        f"{quantidade_tentativas}"
    )

    log(
        "Linhas de pedidos na fila: "
        f"{len(df_pendencias)}"
    )

    indices_remover_fila = []

    # ======================================================
    # PROCESSAMENTO POR ID_EMAIL
    # ======================================================

    for id_email_anterior, grupo in df_pendencias.groupby(
        "ID_EMAIL",
        dropna=False,
        sort=False,
    ):
        # --------------------------------------------------
        # VERIFICA PARADA ANTES DO PRÓXIMO ID
        # --------------------------------------------------

        if parada_solicitada():
            resumo["interrompido"] = True

            log("")
            log(
                "Reenvio interrompido por solicitação "
                "do usuário."
            )

            log(
                "Nenhum novo ID será processado."
            )

            break

        grupo = grupo.copy()

        indices_grupo = grupo.index.tolist()

        resumo["processados"] += 1

        log("")
        log("-" * 80)
        log(
            f"ID da tentativa anterior: {id_email_anterior}"
        )

        log(
            f"Quantidade de linhas: {len(grupo)}"
        )

        fornecedor = obter_primeiro_valor(
            dataframe=grupo,
            coluna="Nº CONTA DO FORNECEDOR",
            valor_padrao="FORNECEDOR_SEM_NOME",
        )

        email_fornecedor = obter_emails_fornecedor(
            grupo
        )

        email_destino_anterior = obter_primeiro_valor(
            dataframe=grupo,
            coluna="EMAIL_DESTINO",
            valor_padrao="",
        )

        # --------------------------------------------------
        # DEFINIÇÃO DO DESTINATÁRIO
        # --------------------------------------------------

        if modo_teste:
            email_destino = str(
                email_teste
            ).strip()

        else:
            email_destino = email_fornecedor

            if not email_destino:
                email_destino = (
                    email_destino_anterior
                )

        novo_id_email = gerar_id_email(
            tamanho=8
        )

        data_tentativa = datetime.now()

        data_retorno = calcular_data_retorno(
            data_envio=data_tentativa,
            dias_para_retorno=dias_para_retorno,
        )

        data_retorno_formatada = (
            data_retorno.strftime("%d/%m/%Y")
        )

        assunto = (
            f"[{novo_id_email}] "
            "Reenvio - Atualização de Previsão de Entrega - "
            f"{fornecedor}"
        )

        log(
            f"Fornecedor: {fornecedor}"
        )

        log(
            "E-mails encontrados: "
            f"{email_fornecedor or 'Sem e-mail'}"
        )

        log(
            "E-mail de destino: "
            f"{email_destino or 'Sem e-mail'}"
        )

        log(
            f"Novo ID do envio: {novo_id_email}"
        )

        log(
            "Data limite de retorno: "
            f"{data_retorno_formatada}"
        )

        # ==================================================
        # AUSÊNCIA DE DESTINATÁRIO
        # ==================================================

        if not email_destino:
            status_envio = "NÃO ENVIADO"

            motivo = (
                "Reenvio não realizado porque nenhum "
                "e-mail de destino foi encontrado."
            )

            log(
                motivo
            )

            resumo["nao_enviados"] += 1

            df_reenvio.loc[
                indices_grupo,
                "ID_EMAIL",
            ] = novo_id_email

            df_reenvio.loc[
                indices_grupo,
                "STATUS_ENVIO",
            ] = status_envio

            df_reenvio.loc[
                indices_grupo,
                "DATA_ENVIO",
            ] = pd.NaT

            df_reenvio.loc[
                indices_grupo,
                "EMAIL_DESTINO",
            ] = ""

            df_reenvio.loc[
                indices_grupo,
                "DATA_TENTATIVA_REENVIO",
            ] = data_tentativa

            df_reenvio.loc[
                indices_grupo,
                "MOTIVO_ULTIMA_TENTATIVA",
            ] = motivo

            adicionar_log_reenvio(
                lista_logs=logs_execucao,
                fornecedor=fornecedor,
                email_fornecedor=email_fornecedor,
                email_destino="",
                id_email=novo_id_email,
                assunto=assunto,
                status_envio=status_envio,
                motivo=motivo,
                quantidade_itens=len(grupo),
                caminho_arquivo="",
                data_tentativa=data_tentativa,
                data_envio=None,
                data_retorno=data_retorno,
            )

            continue

 