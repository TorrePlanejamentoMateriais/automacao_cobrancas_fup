import contextlib
import io
import threading
import traceback
import sys

from main import executar_processamento


class SaidaParaInterface(io.TextIOBase):
    """Redireciona prints do processamento para o terminal e para o Tkinter."""

    def __init__(self, callback_log=None):
        super().__init__()
        self.callback_log = callback_log
        self._buffer = ""
        self._lock = threading.Lock()

    def writable(self):
        return True

    def write(self, texto):
        if texto is None:
            return 0

        texto = str(texto)

        with self._lock:
            self._buffer += texto

            while "\n" in self._buffer:
                linha, self._buffer = self._buffer.split("\n", 1)
                self._enviar(linha.rstrip("\r"))

        return len(texto)

    def flush(self):
        with self._lock:
            if self._buffer:
                self._enviar(self._buffer.rstrip("\r"))
                self._buffer = ""

    def _enviar(self, mensagem):
        if not mensagem:
            return

        # Em executáveis criados com --windowed, sys.__stdout__
        # normalmente é None porque não existe console aberto.
        saida_terminal = getattr(sys, "__stdout__", None)

        if saida_terminal is not None:
            try:
                saida_terminal.write(mensagem + "\n")
                saida_terminal.flush()
            except Exception:
                pass

        # O callback mantém as mensagens visíveis no log do Tkinter.
        if self.callback_log is not None:
            try:
                self.callback_log(mensagem)
            except Exception:
                pass


def executar_envios(callback_log=None, evento_parar=None):
    """Executa diretamente a funcao do main, sem procurar main.py via subprocess."""

    resultado = {
        "sucesso": False,
        "interrompido": False,
        "codigo_retorno": None,
    }

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

    try:
        if evento_parar is not None and evento_parar.is_set():
            resultado["interrompido"] = True
            log("A execucao foi cancelada antes de iniciar.")
            return resultado

        log("=" * 80)
        log("INICIANDO PROCESSAMENTO")
        log("=" * 80)

        saida = SaidaParaInterface(callback_log=callback_log)

        with contextlib.redirect_stdout(saida), contextlib.redirect_stderr(saida):
            executar_processamento(
                evento_parar=evento_parar,
                callback_log=callback_log,
            )

        saida.flush()

        if evento_parar is not None and evento_parar.is_set():
            resultado["sucesso"] = False
            resultado["interrompido"] = True
            resultado["codigo_retorno"] = 0
            log("")
            log("Processamento interrompido pelo usuário.")
            return resultado

        resultado["sucesso"] = True
        resultado["codigo_retorno"] = 0

        log("")
        log("Processo concluído com sucesso.")
        return resultado

    except Exception as erro:
        resultado["sucesso"] = False
        resultado["codigo_retorno"] = 1

        if evento_parar is not None and evento_parar.is_set():
            resultado["interrompido"] = True

        log(f"ERRO NO PROCESSAMENTO: {erro}")
        log(traceback.format_exc())
        return resultado
