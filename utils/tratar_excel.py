from pathlib import Path
import re
import shutil
import sys
import traceback

import pandas as pd

from utils.obter_base_pedidos import obter_base_pedidos


def obter_diretorio_base():
    """Retorna a pasta do EXE ou a raiz do projeto em desenvolvimento."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent.parent


def extrair_emails(texto):
    """Extrai, normaliza e remove e-mails duplicados mantendo a ordem."""
    if not isinstance(texto, str):
        return []

    emails = re.findall(
        r"[\w.+'-]+@[\w.-]+\.\w+",
        texto.lower(),
    )

    return list(dict.fromkeys(emails))


def tratar_excel(
    callback_log=None,
    evento_parar=None,
):
    """
    Localiza a base mais recente, copia para data/raw, trata as abas APP e
    FORNECEDOR e salva data/cleansing/pedidos_tratados.xlsx.

    Retorna um dicionário compatível com o app.py.
    """

    resultado = {
        "sucesso": False,
        "interrompido": False,
        "arquivo_origem": None,
        "arquivo_raw": None,
        "arquivo_destino": None,            
        "registros": 0,
    }
    
    def log(mensagem=""):
        mensagem = str(mensagem)
        print(mensagem, flush=True)

        if callback_log is not None:
            callback_log(mensagem)

    def parada_solicitada():
        return bool(
            evento_parar is not None
            and evento_parar.is_set()
        )

    def interromper_se_solicitado(etapa):
        if parada_solicitada():
            resultado["interrompido"] = True
            log(f"Tratamento interrompido antes da etapa: {etapa}.")
            return True

        return False

    try:
        base_dir = obter_diretorio_base()

        pasta_raw = base_dir / "data" / "raw"
        pasta_cleansing = base_dir / "data" / "cleansing"
        arquivo_cleansing = pasta_cleansing / "pedidos_tratados.xlsx"

        pasta_raw.mkdir(parents=True, exist_ok=True)
        pasta_cleansing.mkdir(parents=True, exist_ok=True)

        resultado["arquivo_destino"] = str(arquivo_cleansing)

        log("Localizando a base de pedidos mais recente...")

        if interromper_se_solicitado("localização da base"):
            return resultado

        arquivo_xlsx = obter_base_pedidos()

        if not arquivo_xlsx:
            raise FileNotFoundError(
                "A função obter_base_pedidos não retornou um arquivo."
            )

        arquivo_xlsx = Path(arquivo_xlsx)

        if not arquivo_xlsx.exists():
            raise FileNotFoundError(
                f"Base de pedidos não encontrada: {arquivo_xlsx}"
            )

        if arquivo_xlsx.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(
                "A base deve estar no formato .xlsx ou .xlsm. "
                f"Arquivo recebido: {arquivo_xlsx.name}"
            )

        resultado["arquivo_origem"] = str(arquivo_xlsx)
        log(f"Arquivo encontrado: {arquivo_xlsx}")

        if interromper_se_solicitado("cópia para RAW"):
            return resultado

        arquivo_raw = pasta_raw / arquivo_xlsx.name

        if arquivo_xlsx.resolve() != arquivo_raw.resolve():
            shutil.copy2(arquivo_xlsx, arquivo_raw)

        resultado["arquivo_raw"] = str(arquivo_raw)
        log(f"Arquivo copiado para RAW: {arquivo_raw}")

        if interromper_se_solicitado("leitura da aba APP"):
            return resultado

        log("Lendo a aba APP...")
        df = pd.read_excel(
            arquivo_raw,
            sheet_name="APP",
            skiprows=1,
            engine="openpyxl",
        )

        df.columns = [str(coluna).strip() for coluna in df.columns]
        log(f"Total de registros carregados: {len(df)}")

        colunas = [
            "PO+ITEM",
            "PEDIDO",
            "ITEM",
            "UNID/NEG",
            "DATA DO PEDIDO",
            "CÓDIGO",
            "DESCRIÇÃO",
            "N° CONTA",
            "Nº CONTA DO FORNECEDOR",
            "FAT",
            "QTD PEDIDO",
            "QTD PENDENTE",
            "STATUS",
            "DATA (SLA)",
            "DAT.NECESSIDADE",
            "COMENTÁRIOS",
            "PROJETO",
            "PEP- RESUMO",
            "Previsão atual 1",
            "Previsão atual 2",
        ]

        colunas_faltantes = [
            coluna
            for coluna in colunas
            if coluna not in df.columns
        ]

        colunas_obrigatorias = [
            "Nº CONTA DO FORNECEDOR",
            "STATUS",
            "Previsão atual 2",
        ]

        obrigatorias_faltantes = [
            coluna
            for coluna in colunas_obrigatorias
            if coluna not in df.columns
        ]

        if obrigatorias_faltantes:
            raise KeyError(
                "Colunas obrigatórias não encontradas na aba APP: "
                + ", ".join(obrigatorias_faltantes)
            )

        if colunas_faltantes:
            log(
                "Aviso: colunas não encontradas na aba APP: "
                + ", ".join(colunas_faltantes)
            )

        colunas_existentes = [
            coluna
            for coluna in colunas
            if coluna in df.columns
        ]

        df_filtrado = df[colunas_existentes].copy()

        df_filtrado[
            "PREVISÃO ATUAL(Preencha Aqui a nova data)"
        ] = pd.NA

        colunas_data = [
            "DATA DO PEDIDO",
            "DATA (SLA)",
            "DAT.NECESSIDADE",
            "Previsão atual 1",
            "Previsão atual 2",
        ]

        for coluna in colunas_data:
            if coluna in df_filtrado.columns:
                df_filtrado[coluna] = pd.to_datetime(
                    df_filtrado[coluna],
                    dayfirst=True,
                    errors="coerce",
                )

        for coluna in df_filtrado.columns:
            if df_filtrado[coluna].dtype == "object":
                df_filtrado[coluna] = (
                    df_filtrado[coluna]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

        if interromper_se_solicitado("leitura da aba FORNECEDOR"):
            return resultado

        log("Lendo a aba FORNECEDOR...")
        df_forn = pd.read_excel(
            arquivo_raw,
            sheet_name="FORNECEDOR",
            skiprows=1,
            engine="openpyxl",
        )

        df_forn.columns = [
            str(coluna).strip()
            for coluna in df_forn.columns
        ]

        df_forn = df_forn.drop(
            columns=["Unnamed: 0"],
            errors="ignore",
        )

        colunas_fornecedor_faltantes = [
            coluna
            for coluna in ["FORNECEDOR", "E-MAIL"]
            if coluna not in df_forn.columns
        ]

        if colunas_fornecedor_faltantes:
            raise KeyError(
                "Colunas obrigatórias não encontradas na aba FORNECEDOR: "
                + ", ".join(colunas_fornecedor_faltantes)
            )

        df_filtrado["Nº CONTA DO FORNECEDOR"] = (
            df_filtrado["Nº CONTA DO FORNECEDOR"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

        df_forn["FORNECEDOR"] = (
            df_forn["FORNECEDOR"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

        df_forn["E-MAIL"] = (
            df_forn["E-MAIL"]
            .apply(extrair_emails)
            .str.join("; ")
        )

        # Consolida fornecedores duplicados sem eliminar pedidos da base APP.
        df_forn = (
            df_forn.groupby("FORNECEDOR", as_index=False)["E-MAIL"]
            .agg(
                lambda valores: "; ".join(
                    dict.fromkeys(
                        email.strip()
                        for valor in valores
                        for email in str(valor).split(";")
                        if email.strip()
                    )
                )
            )
        )

        if interromper_se_solicitado("junção dos e-mails"):
            return resultado

        df_filtrado = pd.merge(
            df_filtrado,
            df_forn[["FORNECEDOR", "E-MAIL"]],
            left_on="Nº CONTA DO FORNECEDOR",
            right_on="FORNECEDOR",
            how="left",
        )

        df_filtrado.rename(
            columns={"E-MAIL": "EmailFornecedor"},
            inplace=True,
        )

        df_filtrado.drop(
            columns=["FORNECEDOR"],
            inplace=True,
            errors="ignore",
        )

        df_filtrado["EmailFornecedor"] = (
            df_filtrado["EmailFornecedor"]
            .fillna("")
            .apply(extrair_emails)
            .str.join("; ")
            .replace("", "Sem E-mail")
        )

        status_permitidos = {
            "no aguardo",
            "no aguardado",
            "parcial",
            "entrega parcial",
        }

        status_normalizado = (
            df_filtrado["STATUS"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
        )

        df_filtrado = df_filtrado[
            status_normalizado.isin(status_permitidos)
        ].copy()

        log(
            "Registros após o filtro de status: "
            f"{len(df_filtrado)}"
        )

        coluna_previsao = "Previsão atual 2"
        data_hoje = pd.Timestamp.today().normalize()

        previsao = pd.to_datetime(
            df_filtrado[coluna_previsao],
            dayfirst=True,
            errors="coerce",
        )

        df_filtrado = df_filtrado[
            previsao.isna()
            | previsao.lt(data_hoje)
        ].copy()

        log(
            "Registros após o filtro de previsão: "
            f"{len(df_filtrado)}"
        )

        if interromper_se_solicitado("formatação e gravação"):
            return resultado

        colunas_data_exportacao = [
            "DATA DO PEDIDO",
            "DATA (SLA)",
            "DAT.NECESSIDADE",
            "Previsão atual 1",
            "Previsão atual 2",
        ]

        for coluna in colunas_data_exportacao:
            if coluna not in df_filtrado.columns:
                continue

            datas = pd.to_datetime(
                df_filtrado[coluna],
                errors="coerce",
                dayfirst=True,
            )

            df_filtrado[coluna] = (
                datas.dt.strftime("%d/%m/%Y")
                .fillna("")
            )

        log("Salvando a base tratada...")
        df_filtrado.to_excel(
            arquivo_cleansing,
            engine="openpyxl",
            index=False,
        )

        if not arquivo_cleansing.exists():
            raise FileNotFoundError(
                "O arquivo pedidos_tratados.xlsx não foi criado."
            )

        resultado["sucesso"] = True
        resultado["registros"] = len(df_filtrado)

        log(f"Base tratada salva em: {arquivo_cleansing}")
        log(f"Total de registros tratados: {len(df_filtrado)}")

        return resultado

    except PermissionError as erro:
        log(
            "ERRO DE PERMISSÃO: feche a planilha no Excel e tente novamente. "
            f"Detalhes: {erro}"
        )
        raise

    except Exception as erro:
        log(f"ERRO AO TRATAR EXCEL: {erro}")
        log(traceback.format_exc())
        raise


if __name__ == "__main__":
    tratar_excel()
