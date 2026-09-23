from pathlib import Path
import sys
from datetime import datetime, timedelta

import html
import random
import re
import string
import traceback

import pandas as pd

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


from utils.enviar_email import enviar_email_com_anexo
from utils.destaca_coluna_excel import destacar_coluna_excel

from utils.logs import (
    ajustar_planilha_log,
    registrar_log,
    salvar_log_excel,
)


def executar_processamento(evento_parar=None, callback_log=None):
    """Executa o processamento completo, com parada cooperativa entre fornecedores."""

    def parada_solicitada():
        return bool(evento_parar is not None and evento_parar.is_set())

    def informar(mensagem):
        print(mensagem)

    # ==========================================================
    # CONFIGURAÇÕES
    # ==========================================================

    # True:
    # Direciona todos os e-mails para EMAIL_TESTE.
    #
    # False:
    # Utiliza os e-mails encontrados em EmailFornecedor.
    MODO_TESTE = True

    EMAIL_TESTE = "tiago.antunes.terceiros@claro.com.br"

    # True:
    # Abre cada e-mail no Outlook para revisão.
    #
    # False:
    # Envia o e-mail automaticamente.
    #
    # Observação:
    # Quando estiver como True, o Python não consegue confirmar
    # se o usuário clicou manualmente no botão Enviar.
    EXIBIR_ANTES_DE_ENVIAR = False

    # Soma dois dias corridos.
    # Se o resultado cair no sábado ou domingo,
    # transfere o prazo para segunda-feira.
    DIAS_PARA_RETORNO = 2



    # ==========================================================
    # TRATAMENTO DE DATAS
    # ==========================================================


    def calcular_data_retorno(
        data_envio=None,
        dias=2,
    ):
        """
        Calcula a data limite para resposta do fornecedor.

        Soma a quantidade de dias corridos informada.

        Se o resultado cair:
        - no sábado, transfere para segunda-feira;
        - no domingo, transfere para segunda-feira.

        Exemplo para envio na sexta-feira:
        - sexta-feira + 2 dias = domingo;
        - prazo ajustado para segunda-feira.
        """

        if data_envio is None:
            data_envio = datetime.now()

        data_retorno = (
            data_envio.replace(
                hour=0,
                minute=0,
                second=0,
                microsecond=0,
            )
            + timedelta(days=dias)
        )

        # Segunda-feira = 0
        # Sexta-feira = 4
        # Sábado = 5
        # Domingo = 6

        if data_retorno.weekday() == 5:
            data_retorno += timedelta(days=2)

        elif data_retorno.weekday() == 6:
            data_retorno += timedelta(days=1)

        return data_retorno


    # ==========================================================
    # FUNÇÕES AUXILIARES
    # ==========================================================

    def gerar_id_email(tamanho=8):
        """
        Gera um ID aleatório para localizar posteriormente
        o e-mail no Outlook.
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


    def limpar_nome_arquivo(nome):
        """
        Remove caracteres inválidos para nomes de arquivos
        no Windows.
        """

        nome_limpo = re.sub(
            r'[<>:"/\\|?*]',
            "_",
            str(nome),
        )

        nome_limpo = re.sub(
            r"\s+",
            " ",
            nome_limpo,
        ).strip()

        return nome_limpo.rstrip(". ")



    # ==========================================================
    # CAMINHOS
    # ==========================================================

    if getattr(sys, "frozen", False):
        BASE_DIR = Path(sys.executable).resolve().parent
    else:
        BASE_DIR = Path(__file__).resolve().parent

    arquivo_excel = (
        BASE_DIR
        / "data"
        / "cleansing"
        / "pedidos_tratados.xlsx"
    )

    pasta_fornecedores = (
        BASE_DIR
        / "base_fornecedores"
    )

    pasta_assinatura = (
        BASE_DIR
        / "assinatura"
    )

    pasta_log = (
        BASE_DIR
        / "log"
    )

    arquivo_reenvio = (
        pasta_log
        / "base_reenvio.xlsx"
    )

    arquivo_log = (
        pasta_log
        / "acompanhamento_envios_fornecedores.xlsx"
    )

    imagem_claro = (
        pasta_assinatura
        / "imagem_claro.png"
    )

    simbolo_pequeno = (
        pasta_assinatura
        / "simbolo_pequeno.png"
    )

    pasta_fornecedores.mkdir(
        parents=True,
        exist_ok=True,
    )

    pasta_log.mkdir(
        parents=True,
        exist_ok=True,
    )


    # ==========================================================
    # VALIDAÇÕES
    # ==========================================================

    if not arquivo_excel.exists():

        raise FileNotFoundError(
            f"Excel não encontrado: {arquivo_excel}"
        )

    if not imagem_claro.exists():

        raise FileNotFoundError(
            f"Imagem não encontrada: {imagem_claro}"
        )

    if not simbolo_pequeno.exists():

        raise FileNotFoundError(
            f"Imagem não encontrada: {simbolo_pequeno}"
        )

    print(f"Pasta destino: {pasta_fornecedores}")
    print(f"Pasta de logs: {pasta_log}")
    print(f"Imagem principal: {imagem_claro}")
    print(f"Símbolo pequeno: {simbolo_pequeno}")


    # ==========================================================
    # LISTA DE LOGS
    # ==========================================================

    logs_execucao = []

    lista_reenvio = []

    # ==========================================================
    # LEITURA DO EXCEL
    # ==========================================================

    df = pd.read_excel(
        arquivo_excel,
        engine="openpyxl",
    )

    print(
        f"Total de registros carregados: {len(df)}"
    )

    data_tentativa = datetime.today()
    data_tentativa_formt = data_tentativa.strftime("%d/%m/%Y")

    # ==========================================================
    # PROCESSAMENTO DOS FORNECEDORES
    # ==========================================================

    try:

        for fornecedor, grupo in df.groupby(
            "Nº CONTA DO FORNECEDOR",
            dropna=False,
        ):

            if parada_solicitada():
                informar("")
                informar("=" * 80)
                informar("PROCESSAMENTO INTERROMPIDO PELO USUÁRIO")
                informar("=" * 80)
                break

            print("\n" + "=" * 80)
            print(f"Fornecedor: {fornecedor}")
            print(
                f"Quantidade de linhas: {len(grupo)}"
            )

            # --------------------------------------------------
            # IDENTIFICADOR PESQUISÁVEL
            # --------------------------------------------------

            id_email = gerar_id_email(
                tamanho=8
            )

            fornecedor_limpo = limpar_nome_arquivo(
                fornecedor
            )

            if not fornecedor_limpo:

                fornecedor_limpo = (
                    "FORNECEDOR_SEM_NOME"
                )

            # --------------------------------------------------
            # E-MAIL DO FORNECEDOR
            # --------------------------------------------------

            email_fornecedor = ";".join(
                    grupo["EmailFornecedor"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .unique()
                )
        

            print(
                "E-mails válidos encontrados: "
                f"{email_fornecedor or 'Nenhum'}"
            )

            # --------------------------------------------------
            # DATA E PRAZO
            # --------------------------------------------------

            data_processamento = datetime.now()

            data_retorno = calcular_data_retorno(
                data_envio=data_processamento,
                dias=DIAS_PARA_RETORNO,
            )

            data_retorno_formatada = (
                data_retorno.strftime("%d/%m/%Y")
            )

            # --------------------------------------------------
            # DESTINATÁRIO
            # --------------------------------------------------

            if MODO_TESTE:

                email_destino = EMAIL_TESTE

            else:

                email_destino = email_fornecedor

            # --------------------------------------------------
            # ASSUNTO
            # --------------------------------------------------

            assunto = (
                f"[{id_email}] "
                f"Atualização de Previsão de Entrega - "
                f"{fornecedor}"
            )

            # --------------------------------------------------
            # FORNECEDOR SEM E-MAIL
            # --------------------------------------------------

            if (
                not MODO_TESTE
                and not email_fornecedor
            ):
                trocar_corpo_msg = False
                mensagem = (
                    "Fornecedor sem e-mail válido na base."
                )

                print(mensagem)
            
                registrar_log(
                    lista_logs=logs_execucao,
                    fornecedor=fornecedor,
                    email_fornecedor="",
                    email_destino="",
                    id_email=id_email,
                    assunto=assunto,
                    status_envio="NÃO ENVIADO",
                    motivo=mensagem,
                    quantidade_itens=len(grupo),
                    data_tentiva = data_tentativa_formt,
                    data_envio=None,
                    data_retorno=data_retorno,
                )

                continue

            if (
                MODO_TESTE
                and not email_fornecedor
            ):
                trocar_corpo_msg = False
                print(
                    "Aviso: fornecedor sem e-mail válido. "
                    "A mensagem será direcionada apenas "
                    "ao endereço de teste."
                )

            print(f"E-mail destino: {email_destino}")
            print(f"ID do envio: {id_email}")

            print(
                "Data limite de retorno: "
                f"{data_retorno_formatada}"
            )

            # --------------------------------------------------
            # PREPARAÇÃO DO EXCEL
            # --------------------------------------------------

            grupo_envio = grupo.drop(
                columns=[
                    "EmailFornecedor",
                    "Previsão atual 1",
                    "Previsão atual 2"
                ],
                errors="ignore",
            ).copy()


            # --------------------------------------------------
            # GERAÇÃO DO EXCEL
            # --------------------------------------------------

            arquivo_fornecedor = (
                pasta_fornecedores
                / f"Fornecedor_{fornecedor_limpo}.xlsx"
            )

            try:

                grupo_envio.to_excel(
                    arquivo_fornecedor,
                    index=False,
                    engine="openpyxl",
                )

                if not arquivo_fornecedor.exists():

                    raise FileNotFoundError(
                        "O arquivo Excel não foi criado."
                    )
                
                destacar_coluna_excel(
                    caminho_arquivo=arquivo_fornecedor,
                    nome_coluna="Previsão Atual(Preencha Aqui a nova data)"
                )

                print(
                    f"Arquivo salvo: {arquivo_fornecedor}"
                )

            except Exception as erro:

                mensagem = (
                    "Erro na geração da planilha: "
                    f"{erro}"
                )

                print(mensagem)

                registrar_log(
                    lista_logs=logs_execucao,
                    fornecedor=fornecedor,
                    email_fornecedor=email_fornecedor,
                    email_destino=email_destino,
                    id_email=id_email,
                    assunto=assunto,
                    status_envio="NÃO ENVIADO",
                    motivo=mensagem,
                    quantidade_itens=len(grupo),
                    caminho_arquivo=arquivo_excel,
                    data_tentativa = data_tentativa_formt,
                    data_envio=None,
                    data_retorno=data_retorno,
                )

                continue

            # --------------------------------------------------
            # CORPO DO E-MAIL
            # --------------------------------------------------

            fornecedor_html = html.escape(
                str(fornecedor)
            )

            corpo = f"""
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
            Segue a carteira para atualização da previsão e
            validação da última data informada.
        </p>

        <p>
            Estamos acompanhando os itens com maior atraso e o
            desempenho dos fornecedores. Por gentileza, informar
            a justificativa dos pedidos em atraso e atualizar a
            previsão de entrega dos pedidos na planilha anexa ou
            no corpo do e-mail.
        </p>

        <p>
            Para os itens com entrega parcial, solicitamos o envio
            do cronograma detalhado, com as quantidades pendentes
            e respectivas datas previstas de entrega.
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
                    <img src="cid:imagem_claro">
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

                       <img src="cid:simbolo_pequeno">
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

                    <div>55 11 2612-2542</div>

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

                    <div>Claro Brasil</div>

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



            # --------------------------------------------------
            # ENVIO
            # --------------------------------------------------

            try:

                if parada_solicitada():
                    informar("Parada solicitada. O próximo e-mail não será iniciado.")
                    break

                resultado_envio = enviar_email_com_anexo(
                    destinatarios=email_destino,
                    assunto=assunto,
                    corpo=corpo,
                    caminho_anexo=str(
                        arquivo_fornecedor
                    ),
                    caminho_imagem_claro =str(
                        imagem_claro
                    ),
                    caminho_simbolo_pequeno = str(
                        simbolo_pequeno
                    ),
                    exibir_antes_de_enviar=(
                        EXIBIR_ANTES_DE_ENVIAR
                    ),
                )

            

                if resultado_envio == "ENVIADO":

                    status_envio = "ENVIADO"

                    motivo = (
                        "E-mail enviado automaticamente "
                        "com sucesso."
                    )

                    data_envio_log = datetime.now()

                elif (
                    resultado_envio
                    == "ABERTO_PARA_REVISAO"
                ):

                    status_envio = (
                        "ABERTO PARA REVISÃO"
                    )

                    motivo = (
                        "Mensagem aberta no Outlook. "
                        "O Python não consegue confirmar "
                        "se o envio manual foi realizado."
                    )

                    data_envio_log = None

                else:

                    status_envio = "NÃO ENVIADO"
               
                    motivo = (
                        "Resultado não reconhecido pela "
                        "função de envio: "
                        f"{resultado_envio}"
                    )

                    data_envio_log = None

                print(motivo)

           

                registrar_log(
                    lista_logs=logs_execucao,
                    fornecedor=fornecedor,
                    email_fornecedor=email_fornecedor,
                    email_destino=email_destino,
                    id_email=id_email,
                    assunto=assunto,
                    status_envio=status_envio,
                    motivo=motivo,
                    quantidade_itens=len(grupo),
                    caminho_arquivo=arquivo_excel,
                    data_tentativa = data_tentativa_formt,
                    data_envio=data_envio_log,
                    data_retorno=data_retorno,
                )

                # --------------------------------------------------
                # CONTROLE DE REENVIO
                # --------------------------------------------------
                # Somente mensagens que NÃO foram enviadas entram na
                # fila de reenvio.
                #
                # "ABERTO PARA REVISÃO" não entra na fila, pois o
                # usuário pode enviar manualmente pelo Outlook.
                if status_envio == "NÃO ENVIADO":

                    df_reenvio = grupo.copy()

                    df_reenvio["ID_EMAIL"] = id_email
                    df_reenvio["STATUS_ENVIO"] = status_envio
                    df_reenvio["DATA_ENVIO"] = data_envio_log
                    df_reenvio["EMAIL_DESTINO"] = email_destino

                    lista_reenvio.append(df_reenvio)

            except Exception as erro:

                mensagem = (
                    f"Erro no envio do e-mail: {erro}"
                )

                print(mensagem)

                traceback.print_exc()

                registrar_log(
                    lista_logs=logs_execucao,
                    fornecedor=fornecedor,
                    email_fornecedor=email_fornecedor,
                    email_destino=email_destino,
                    id_email=id_email,
                    assunto=assunto,
                    status_envio="NÃO ENVIADO",
                    motivo=mensagem,
                    quantidade_itens=len(grupo),
                    caminho_arquivo=arquivo_excel,
                    data_tentativa = data_tentativa_formt,
                    data_envio=None,
                    data_retorno=data_retorno,
                )

        # ------------------------------------------
        # SALVAR CONTROLE DE REENVIO
        # ------------------------------------------
        #
        # IMPORTANTE:
        # Esta etapa NÃO executa o reenvio.
        # Ela apenas mantém uma fila de registros cujo envio
        # não foi realizado, para que uma função de reenvio
        # possa ser criada posteriormente.
        #
        # A base é tratada como FILA DE PENDÊNCIAS:
        # - "NÃO ENVIADO" entra;
        # - "ENVIADO" não entra;
        # - "ABERTO PARA REVISÃO" não entra;
        # - registros duplicados são removidos por ID_EMAIL.
        if lista_reenvio:

            df_controle_novo = pd.concat(
                lista_reenvio,
                ignore_index=True
            )

            if arquivo_reenvio.exists():

                try:

                    df_controle_antigo = pd.read_excel(
                        arquivo_reenvio,
                        engine="openpyxl"
                    )

                except Exception as erro:

                    print(
                        "Aviso: não foi possível ler a base de reenvio "
                        f"existente: {erro}"
                    )

                    df_controle_antigo = pd.DataFrame()

            else:

                df_controle_antigo = pd.DataFrame()

            # Junta as pendências antigas com as novas.
            df_controle = pd.concat(
                [
                    df_controle_antigo,
                    df_controle_novo
                ],
                ignore_index=True
            )

            # Remove registros que não deveriam permanecer na fila.
            if "STATUS_ENVIO" in df_controle.columns:

                status_controle = (
                    df_controle["STATUS_ENVIO"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df_controle = df_controle[
                    status_controle.eq("NÃO ENVIADO")
                ].copy()

            # O mesmo ID_EMAIL representa uma única tentativa de envio.
            # Portanto, não deve existir mais de um registro desse ID.
            if "ID_EMAIL" in df_controle.columns:

                df_controle = (
                    df_controle
                    .drop_duplicates(
                        subset=["ID_EMAIL"],
                        keep="last"
                    )
                    .reset_index(drop=True)
                )

            # Garante que a pasta de log exista antes da gravação.
            arquivo_reenvio.parent.mkdir(
                parents=True,
                exist_ok=True
            )

            df_controle.to_excel(
                arquivo_reenvio,
                engine="openpyxl",
                index=False
            )

            print(
                f"Controle de reenvio atualizado: "
                f"{arquivo_reenvio}"
            )

            print(
                "Fornecedores/itens pendentes para reenvio: "
                f"{len(df_controle)}"
            )

        else:

            print(
                "Nenhum novo e-mail não enviado para adicionar "
                "à base de reenvio."
            )

            # Se a base já existir, não a apagamos e nem alteramos.
            # Ela continua disponível para a futura função de reenvio.
            if arquivo_reenvio.exists():

                try:

                    df_controle_existente = pd.read_excel(
                        arquivo_reenvio,
                        engine="openpyxl"
                    )

                    print(
                        "Pendências já existentes na base de reenvio: "
                        f"{len(df_controle_existente)}"
                    )

                except Exception as erro:

                    print(
                        "Não foi possível consultar a base de reenvio "
                        f"existente: {erro}"
                    )
        
    finally:

        # Garante que o histórico seja salvo mesmo se ocorrer
        # uma falha inesperada durante o processamento.
        salvar_log_excel(
            lista_logs=logs_execucao,
            arquivo_log=arquivo_log,
        )


    print("\nProcesso finalizado.")


if __name__ == "__main__":
    executar_processamento()
