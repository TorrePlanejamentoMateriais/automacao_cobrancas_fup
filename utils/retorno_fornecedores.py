from datetime import datetime
from pathlib import Path
import sys
import time
import traceback

import pandas as pd
import pythoncom
import pywintypes
import win32com.client

from utils.logs import fechar_status_retorno


EXTENSOES_PERMITIDAS = {".xlsx", ".xlsm", ".xls"}
CLASSE_EMAIL_OUTLOOK = 43
PASTA_CAIXA_ENTRADA = 6
PREFIXOS_RESPOSTA = ("RE:", "RES:", "ENC:", "FW:", "FWD:")
ERRO_RPC_INDISPONIVEL = -2147023174
MAXIMO_TENTATIVAS_RPC = 2
INTERVALO_RECONEXAO_SEGUNDOS = 2
DIAS_MAXIMOS_PESQUISA = 120
PROPRIEDADE_ASSUNTO_DASL = "http://schemas.microsoft.com/mapi/proptag/0x0037001F"


def limpar_nome_anexo(nome_arquivo):
    """Remove caracteres invalidos de nomes de arquivos no Windows."""
    caracteres_invalidos = '<>:"/\\|?*'
    nome_limpo = "".join(
        "_" if caractere in caracteres_invalidos else caractere
        for caractere in str(nome_arquivo)
    )
    nome_limpo = " ".join(nome_limpo.split())
    return nome_limpo.strip().rstrip(". ")


def criar_caminho_unico(caminho):
    """Evita substituir um arquivo que ja existe."""
    caminho = Path(caminho)
    if not caminho.exists():
        return caminho

    contador = 1
    while True:
        novo_caminho = caminho.parent / (
            f"{caminho.stem}_{contador}{caminho.suffix}"
        )
        if not novo_caminho.exists():
            return novo_caminho
        contador += 1


def normalizar_data_comparacao(data):
    """Converte uma data para datetime sem timezone."""
    if data is None:
        return None

    try:
        if pd.isna(data):
            return None
    except (TypeError, ValueError):
        pass

    try:
        data_convertida = pd.Timestamp(data).to_pydatetime()
    except Exception:
        return None

    if getattr(data_convertida, "tzinfo", None) is not None:
        data_convertida = data_convertida.replace(tzinfo=None)

    return data_convertida


def obter_email_remetente(mensagem):
    """Obtem o endereco SMTP do remetente."""
    try:
        tipo_remetente = str(
            getattr(mensagem, "SenderEmailType", "") or ""
        ).strip().upper()

        if tipo_remetente == "EX":
            remetente_exchange = getattr(mensagem, "Sender", None)
            if remetente_exchange is not None:
                usuario_exchange = remetente_exchange.GetExchangeUser()
                if usuario_exchange is not None:
                    email_smtp = str(
                        usuario_exchange.PrimarySmtpAddress or ""
                    ).strip()
                    if email_smtp:
                        return email_smtp

        return str(
            getattr(mensagem, "SenderEmailAddress", "") or ""
        ).strip()

    except Exception:
        try:
            return str(
                getattr(mensagem, "SenderEmailAddress", "") or ""
            ).strip()
        except Exception:
            return ""


def assunto_eh_resposta(assunto):
    """Verifica se o assunto indica resposta ou encaminhamento."""
    assunto_normalizado = str(assunto).strip().upper()
    return assunto_normalizado.startswith(PREFIXOS_RESPOSTA)


def obter_codigo_com(erro):
    """Tenta obter o codigo HRESULT de uma excecao COM."""
    if not isinstance(erro, pywintypes.com_error):
        return None

    try:
        return int(erro.args[0])
    except (IndexError, TypeError, ValueError):
        return None


def erro_rpc_indisponivel(erro):
    """Informa se a excecao representa perda da conexao RPC com o Outlook."""
    return obter_codigo_com(erro) == ERRO_RPC_INDISPONIVEL


def liberar_objetos_outlook(*objetos):
    """Remove referencias locais para objetos COM do Outlook."""
    for objeto in objetos:
        try:
            del objeto
        except Exception:
            pass


def conectar_outlook():
    """Cria uma nova conexao COM com o Outlook classico."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    caixa_entrada = namespace.GetDefaultFolder(PASTA_CAIXA_ENTRADA)
    mensagens = caixa_entrada.Items

    try:
        mensagens.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    return outlook, namespace, caixa_entrada, mensagens


def escapar_texto_dasl(texto):
    """Escapa apostrofos para uso seguro em filtros DASL."""
    return str(texto).replace("'", "''")


def restringir_mensagens_por_data(mensagens, data_minima):
    """Reduz a colecao aos e-mails recebidos a partir da data informada."""
    data_minima = normalizar_data_comparacao(data_minima)
    if data_minima is None:
        return mensagens

    filtro_data = (
        "[ReceivedTime] >= '"
        + data_minima.strftime("%m/%d/%Y %I:%M %p")
        + "'"
    )
    return mensagens.Restrict(filtro_data)


def buscar_mensagem_por_id(
    mensagens,
    id_email,
    data_envio=None,
    evento_parar=None,
):
    """Busca rapidamente a resposta pelo ID usando filtro no Outlook."""
    id_email = str(id_email).strip().upper()
    if not id_email:
        return None

    if evento_parar is not None and evento_parar.is_set():
        return None

    marcador_id = f"[{id_email}]"
    marcador_dasl = escapar_texto_dasl(marcador_id)
    data_envio_comparacao = normalizar_data_comparacao(data_envio)

    # O Outlook executa o filtro internamente. Isso evita percorrer toda a
    # caixa de entrada para cada ID pesquisado.
    filtro_assunto = (
        '@SQL="' + PROPRIEDADE_ASSUNTO_DASL + '" '
        "ci_phrasematch '" + marcador_dasl + "'"
    )

    try:
        mensagens_filtradas = mensagens.Restrict(filtro_assunto)
    except pywintypes.com_error as erro:
        if erro_rpc_indisponivel(erro):
            raise
        # Fallback para compatibilidade com ambientes que rejeitem DASL.
        mensagens_filtradas = mensagens

    try:
        mensagens_filtradas.Sort("[ReceivedTime]", True)
    except Exception as erro:
        if erro_rpc_indisponivel(erro):
            raise

    quantidade = mensagens_filtradas.Count

    for indice in range(1, quantidade + 1):
        if evento_parar is not None and evento_parar.is_set():
            return None

        try:
            mensagem = mensagens_filtradas.Item(indice)

            if getattr(mensagem, "Class", None) != CLASSE_EMAIL_OUTLOOK:
                continue

            assunto = str(
                getattr(mensagem, "Subject", "") or ""
            ).strip()

            if marcador_id not in assunto.upper():
                continue

            if not assunto_eh_resposta(assunto):
                continue

            data_recebimento = getattr(mensagem, "ReceivedTime", None)
            data_recebimento_comparacao = normalizar_data_comparacao(
                data_recebimento
            )

            if (
                data_envio_comparacao is not None
                and data_recebimento_comparacao is not None
                and data_recebimento_comparacao <= data_envio_comparacao
            ):
                continue

            return mensagem

        except pywintypes.com_error as erro:
            if erro_rpc_indisponivel(erro):
                raise
            continue
        except Exception:
            continue

    return None


def salvar_anexos_excel(
    mensagem,
    pasta_destino,
    evento_parar=None,
):
    """
    Salva os anexos Excel encontrados.

    Caso o mesmo arquivo ja exista, reutiliza o caminho existente para
    evitar copias duplicadas.
    """
    pasta_destino = Path(pasta_destino)
    pasta_destino.mkdir(parents=True, exist_ok=True)
    arquivos_salvos = []
    anexos = mensagem.Attachments

    for numero_anexo in range(1, anexos.Count + 1):
        if evento_parar is not None and evento_parar.is_set():
            break

        anexo = anexos.Item(numero_anexo)
        nome_anexo = limpar_nome_anexo(anexo.FileName)

        if not nome_anexo:
            continue

        extensao = Path(nome_anexo).suffix.lower()
        if extensao not in EXTENSOES_PERMITIDAS:
            continue

        caminho_saida = pasta_destino / nome_anexo

        if caminho_saida.exists():
            arquivos_salvos.append(str(caminho_saida))
            continue

        anexo.SaveAsFile(str(caminho_saida))
        arquivos_salvos.append(str(caminho_saida))

    return arquivos_salvos


def processar_retornos_fornecedores(
    arquivo_log,
    pasta_respostas,
    callback_log=None,
    evento_parar=None,
):
    """
    Le a aba Aguardando_Retorno, procura respostas no Outlook, baixa anexos
    Excel e fecha o status no Historico.
    """

    def log(mensagem=""):
        mensagem = str(mensagem)
        saida_terminal = getattr(sys, "__stdout__", None)

        if saida_terminal is not None:
            try:
                saida_terminal.write(mensagem + "\n")
                saida_terminal.flush()
            except Exception:
                pass

        if callback_log is not None:
            try:
                callback_log(mensagem)
            except Exception:
                pass

    def parada_solicitada():
        return bool(
            evento_parar is not None and evento_parar.is_set()
        )

    arquivo_log = Path(arquivo_log)
    pasta_respostas = Path(pasta_respostas)

    resumo = {
        "verificados": 0,
        "respondidos_com_anexo": 0,
        "respondidos_sem_anexo": 0,
        "nao_localizados": 0,
        "erros": 0,
        "interrompido": False,
    }

    if parada_solicitada():
        resumo["interrompido"] = True
        log("Consulta de respostas cancelada antes do inicio.")
        return resumo

    if not arquivo_log.exists():
        raise FileNotFoundError(
            f"Arquivo de acompanhamento nao encontrado: {arquivo_log}"
        )

    pasta_respostas.mkdir(parents=True, exist_ok=True)

    try:
        df_aguardando = pd.read_excel(
            arquivo_log,
            sheet_name="Aguardando_Retorno",
            engine="openpyxl",
        )
    except PermissionError as erro:
        raise PermissionError(
            "Nao foi possivel abrir o acompanhamento. "
            "Feche o arquivo no Excel e tente novamente."
        ) from erro
    except ValueError as erro:
        raise ValueError(
            "A aba 'Aguardando_Retorno' nao foi encontrada."
        ) from erro
    except Exception as erro:
        raise RuntimeError(
            f"Nao foi possivel ler o acompanhamento: {erro}"
        ) from erro

    if df_aguardando.empty:
        log("Nao existem fornecedores aguardando retorno.")
        return resumo

    if "ID_Email" not in df_aguardando.columns:
        raise KeyError(
            "A coluna 'ID_Email' nao foi encontrada na aba "
            "Aguardando_Retorno."
        )

    df_aguardando["ID_Email"] = (
        df_aguardando["ID_Email"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    df_aguardando = df_aguardando[
        df_aguardando["ID_Email"].ne("")
    ].copy()
    df_aguardando = (
        df_aguardando
        .drop_duplicates(subset=["ID_Email"], keep="last")
        .reset_index(drop=True)
    )

    if df_aguardando.empty:
        log("Nao existem IDs validos aguardando retorno.")
        return resumo

    log("")
    log("=" * 80)
    log("CONSULTA DE RETORNOS NO OUTLOOK")
    log("=" * 80)
    log(f"IDs aguardando retorno: {len(df_aguardando)}")

    com_inicializado = False
    outlook = None
    namespace = None
    caixa_entrada = None
    mensagens = None

    try:
        # Esta funcao roda em uma thread do Tkinter. Cada thread que usa COM
        # precisa inicializar o proprio apartment COM.
        pythoncom.CoInitialize()
        com_inicializado = True

        outlook, namespace, caixa_entrada, mensagens = conectar_outlook()

        datas_envio_validas = pd.to_datetime(
            df_aguardando.get("DataEnvio"),
            errors="coerce",
            dayfirst=True,
        ) if "DataEnvio" in df_aguardando.columns else pd.Series(dtype="datetime64[ns]")

        if not datas_envio_validas.dropna().empty:
            data_minima_busca = datas_envio_validas.dropna().min().to_pydatetime()
        else:
            data_minima_busca = datetime.now() - pd.Timedelta(days=DIAS_MAXIMOS_PESQUISA)

        mensagens = restringir_mensagens_por_data(
            mensagens,
            data_minima_busca,
        )

        log(
            "Periodo da pesquisa: mensagens recebidas desde "
            f"{data_minima_busca.strftime('%d/%m/%Y')}."
        )

        log(
            "Pasta consultada no Outlook: "
            f"{caixa_entrada.FolderPath}"
        )

        for _, registro in df_aguardando.iterrows():
            if parada_solicitada():
                resumo["interrompido"] = True
                log("")
                log("Consulta interrompida por solicitacao do usuario.")
                break

            id_email = str(registro["ID_Email"]).strip().upper()
            fornecedor = str(
                registro.get("Fornecedor", "FORNECEDOR_SEM_NOME")
                or "FORNECEDOR_SEM_NOME"
            ).strip()

            if not fornecedor or fornecedor.casefold() in {"nan", "none"}:
                fornecedor = "FORNECEDOR_SEM_NOME"

            data_envio = registro.get("DataEnvio")
            resumo["verificados"] += 1

            log("")
            log("-" * 80)
            log(f"ID pesquisado: {id_email}")
            log(f"Fornecedor: {fornecedor}")

            try:
                mensagem = None

                for tentativa in range(1, MAXIMO_TENTATIVAS_RPC + 1):
                    if parada_solicitada():
                        resumo["interrompido"] = True
                        break

                    try:
                        mensagem = buscar_mensagem_por_id(
                            mensagens=mensagens,
                            id_email=id_email,
                            data_envio=data_envio,
                            evento_parar=evento_parar,
                        )
                        break

                    except pywintypes.com_error as erro:
                        if not erro_rpc_indisponivel(erro):
                            raise

                        if tentativa >= MAXIMO_TENTATIVAS_RPC:
                            raise

                        log(
                            "A conexao RPC com o Outlook foi perdida. "
                            f"Reconectando, tentativa {tentativa + 1} de "
                            f"{MAXIMO_TENTATIVAS_RPC}..."
                        )

                        liberar_objetos_outlook(
                            mensagens,
                            caixa_entrada,
                            namespace,
                            outlook,
                        )
                        mensagens = None
                        caixa_entrada = None
                        namespace = None
                        outlook = None

                        time.sleep(INTERVALO_RECONEXAO_SEGUNDOS)

                        outlook, namespace, caixa_entrada, mensagens = (
                            conectar_outlook()
                        )
                        mensagens = restringir_mensagens_por_data(
                            mensagens,
                            data_minima_busca,
                        )

                        log("Conexao com o Outlook restabelecida.")

                if parada_solicitada():
                    resumo["interrompido"] = True
                    log(f"Busca interrompida durante o ID {id_email}.")
                    break

                if mensagem is None:
                    resumo["nao_localizados"] += 1
                    log("Resultado: resposta nao localizada.")
                    continue

                assunto_resposta = str(
                    getattr(mensagem, "Subject", "") or ""
                ).strip()
                email_resposta = obter_email_remetente(mensagem)
                data_resposta = getattr(
                    mensagem,
                    "ReceivedTime",
                    datetime.now(),
                )

                fornecedor_limpo = limpar_nome_anexo(fornecedor)
                if not fornecedor_limpo:
                    fornecedor_limpo = "FORNECEDOR_SEM_NOME"

                pasta_id = pasta_respostas / fornecedor_limpo / id_email

                log("Resposta localizada no Outlook.")
                log(
                    "Remetente: "
                    f"{email_resposta or 'Nao identificado'}"
                )
                log(f"Assunto: {assunto_resposta}")

                arquivos_salvos = salvar_anexos_excel(
                    mensagem=mensagem,
                    pasta_destino=pasta_id,
                    evento_parar=evento_parar,
                )

                if parada_solicitada():
                    resumo["interrompido"] = True
                    log(
                        "Interrupcao solicitada durante o download "
                        f"dos anexos do ID {id_email}."
                    )
                    if arquivos_salvos:
                        for caminho in arquivos_salvos:
                            log(f"Anexo ja salvo: {caminho}")
                    log("O status continuara como AGUARDANDO RETORNO.")
                    break

                if arquivos_salvos:
                    arquivo_resposta = ";".join(arquivos_salvos)
                    atualizado = fechar_status_retorno(
                        arquivo_log=arquivo_log,
                        id_email=id_email,
                        status_retorno="RESPONDIDO",
                        data_resposta=data_resposta,
                        email_resposta=email_resposta,
                        assunto_resposta=assunto_resposta,
                        arquivo_resposta=arquivo_resposta,
                        motivo_resposta=(
                            "Resposta localizada no Outlook e anexo "
                            "Excel baixado."
                        ),
                    )

                    if atualizado:
                        resumo["respondidos_com_anexo"] += 1
                        log("Resultado: respondido com anexo.")
                        for caminho in arquivos_salvos:
                            log(f"Anexo salvo: {caminho}")
                    else:
                        resumo["erros"] += 1
                        log(
                            "O anexo foi salvo, mas o ID nao foi "
                            "localizado no Historico."
                        )
                else:
                    atualizado = fechar_status_retorno(
                        arquivo_log=arquivo_log,
                        id_email=id_email,
                        status_retorno="RESPONDIDO SEM ANEXO",
                        data_resposta=data_resposta,
                        email_resposta=email_resposta,
                        assunto_resposta=assunto_resposta,
                        arquivo_resposta="",
                        motivo_resposta=(
                            "Resposta localizada no Outlook, mas "
                            "nenhum anexo Excel foi encontrado."
                        ),
                    )

                    if atualizado:
                        resumo["respondidos_sem_anexo"] += 1
                        log(
                            "Resultado: resposta localizada, mas sem "
                            "anexo Excel."
                        )
                    else:
                        resumo["erros"] += 1
                        log(
                            "A resposta foi localizada, mas o ID nao "
                            "foi encontrado no Historico."
                        )

            except PermissionError:
                raise
            except pywintypes.com_error as erro:
                resumo["erros"] += 1
                if erro_rpc_indisponivel(erro):
                    log(
                        f"Erro ao processar o ID {id_email}: a conexao "
                        "RPC com o Outlook continuou indisponivel apos "
                        "a tentativa de reconexao."
                    )
                else:
                    log(f"Erro COM ao processar o ID {id_email}: {erro}")
                log(traceback.format_exc())
            except Exception as erro:
                resumo["erros"] += 1
                log(f"Erro ao processar o ID {id_email}: {erro}")
                log(traceback.format_exc())

    except PermissionError:
        raise
    except Exception as erro:
        raise RuntimeError(
            f"Erro durante a consulta no Outlook: {erro}"
        ) from erro
    finally:
        liberar_objetos_outlook(
            mensagens,
            caixa_entrada,
            namespace,
            outlook,
        )
        mensagens = None
        caixa_entrada = None
        namespace = None
        outlook = None

        if com_inicializado:
            pythoncom.CoUninitialize()

    log("")
    log("=" * 80)
    log("RESUMO DA CONSULTA DE RETORNOS")
    log("=" * 80)
    log(f"IDs verificados: {resumo['verificados']}")
    log(
        "Respondidos com anexo: "
        f"{resumo['respondidos_com_anexo']}"
    )
    log(
        "Respondidos sem anexo: "
        f"{resumo['respondidos_sem_anexo']}"
    )
    log(
        "Respostas nao localizadas: "
        f"{resumo['nao_localizados']}"
    )
    log(f"Erros: {resumo['erros']}")
    log(
        "Execucao interrompida: "
        f"{'SIM' if resumo['interrompido'] else 'NAO'}"
    )
    log("=" * 80)

    return resumo
