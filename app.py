from pathlib import Path
import os
import sys
import threading
import traceback
import inspect

import pythoncom
import win32com.client as win32

import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from arquivo_executar import executar_envios
from utils.tratar_excel import tratar_excel
from utils.logs import salvar_log_excel
from utils.reenviar_email import reenviar_emails_pendentes
from utils.retorno_fornecedores import (
    processar_retornos_fornecedores,
)


# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

MODO_TESTE_PADRAO = True

# O e-mail de teste não fica fixo no código. Quando o modo de
# teste é ativado, a aplicação identifica automaticamente a
# conta padrão configurada no Outlook Classic.
EMAIL_TESTE_PADRAO = ""

EXIBIR_ANTES_DE_ENVIAR = False

DIAS_PARA_RETORNO = 2


# ==========================================================
# CAMINHOS
# ==========================================================

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

PASTA_CLEANSING = BASE_DIR / "data" / "cleansing"
ARQUIVO_BASE_TRATADA = PASTA_CLEANSING / "pedidos_tratados.xlsx"

PASTA_FORNECEDORES = (
    BASE_DIR
    / "base_fornecedores"
)

PASTA_ASSINATURA = (
    BASE_DIR
    / "assinatura"
)

PASTA_LOG = (
    BASE_DIR
    / "log"
)

PASTA_RESPOSTAS = (
    BASE_DIR
    / "respostas_fornecedores"
)

ARQUIVO_REENVIO = (
    PASTA_LOG
    / "base_reenvio.xlsx"
)

ARQUIVO_LOG = (
    PASTA_LOG
    / "acompanhamento_envios_fornecedores.xlsx"
)

IMAGEM_CLARO = (
    PASTA_ASSINATURA
    / "imagem_claro.png"
)

SIMBOLO_PEQUENO = (
    PASTA_ASSINATURA
    / "simbolo_pequeno.png"
)


# ==========================================================
# CRIAÇÃO DAS PASTAS
# ==========================================================

PASTA_CLEANSING.mkdir(parents=True, exist_ok=True)

PASTA_FORNECEDORES.mkdir(
    parents=True,
    exist_ok=True,
)

PASTA_LOG.mkdir(
    parents=True,
    exist_ok=True,
)

PASTA_RESPOSTAS.mkdir(
    parents=True,
    exist_ok=True,
)


# ==========================================================
# APLICAÇÃO
# ==========================================================

class AppEnvios:
    """
    Interface gráfica para:

    - executar o envio normal;
    - reenviar e-mails pendentes;
    - consultar respostas no Outlook;
    - interromper com segurança a execução atual;
    - abrir arquivos e pastas do processo.
    """

    def __init__(self):
        self.root = tk.Tk()

        self.root.title(
            "Acompanhamento de Pedidos"
        )

        self.root.geometry(
            "1250x720"
        )

        self.root.minsize(
            1050,
            620,
        )

        # Thread responsável pelo processamento atual.
        self.thread_execucao = None

        # Identifica o processo que está sendo executado.
        self.tipo_execucao = None

        # Evento utilizado para solicitar a interrupção
        # segura da execução atual.
        self.evento_parar = threading.Event()

        # Impede atualizações da interface depois do
        # fechamento da janela.
        self.aplicacao_fechando = False

        # Configurações controladas diretamente pela interface.
        self.var_modo_teste = tk.BooleanVar(
            value=MODO_TESTE_PADRAO
        )
        self.var_email_teste = tk.StringVar(
            value=""
        )

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.fechar_aplicacao,
        )

        self.criar_interface()
        self.atualizar_modo_teste(
            exibir_erro=False,
        )

    # ======================================================
    # INTERFACE
    # ======================================================

    def criar_interface(self):
        """
        Cria todos os componentes da interface.
        """

        # --------------------------------------------------
        # CABEÇALHO
        # --------------------------------------------------

        frame_topo = ttk.Frame(
            self.root,
            padding=10,
        )

        frame_topo.pack(
            fill="x",
        )

        ttk.Label(
            frame_topo,
            text="Controle de Envios para Fornecedores",
            font=(
                "Segoe UI",
                16,
                "bold",
            ),
        ).pack(
            side="left",
        )

        self.lbl_status = ttk.Label(
            frame_topo,
            text="PARADO",
            foreground="red",
            font=(
                "Segoe UI",
                10,
                "bold",
            ),
        )

        self.lbl_status.pack(
            side="right",
        )

        # --------------------------------------------------
        # MODO DE TESTE
        # --------------------------------------------------

        frame_modo_teste = ttk.LabelFrame(
            self.root,
            text="Segurança dos envios",
            padding=10,
        )

        frame_modo_teste.pack(
            fill="x",
            padx=10,
            pady=5,
        )

        self.chk_modo_teste = ttk.Checkbutton(
            frame_modo_teste,
            text="Ativar modo de teste",
            variable=self.var_modo_teste,
            command=self.atualizar_modo_teste,
        )

        self.chk_modo_teste.pack(
            side="left",
            padx=(0, 15),
        )

        ttk.Label(
            frame_modo_teste,
            text="E-mail identificado no Outlook:",
        ).pack(
            side="left",
        )

        self.lbl_email_teste = ttk.Label(
            frame_modo_teste,
            textvariable=self.var_email_teste,
            foreground="#0067B8",
            font=(
                "Segoe UI",
                9,
                "bold",
            ),
        )

        self.lbl_email_teste.pack(
            side="left",
            padx=8,
        )

        self.btn_atualizar_email = ttk.Button(
            frame_modo_teste,
            text="Atualizar conta",
            command=self.atualizar_email_outlook,
        )

        self.btn_atualizar_email.pack(
            side="right",
        )

        # --------------------------------------------------
        # BOTÕES PRINCIPAIS
        # --------------------------------------------------

        frame_botoes_principais = ttk.LabelFrame(
            self.root,
            text="Processamentos",
            padding=10,
        )

        frame_botoes_principais.pack(
            fill="x",
            padx=10,
            pady=5,
        )

        self.btn_tratar_excel = ttk.Button(
            frame_botoes_principais,
            text="📊 Tratar Excel",
            command=self.iniciar_tratamento_excel,
        )

        self.btn_tratar_excel.pack(
            side="left",
            padx=5,
        )

        self.btn_iniciar = ttk.Button(
            frame_botoes_principais,
            text="▶ Iniciar Envios",
            command=self.iniciar_envios,
        )

        self.btn_iniciar.pack(
            side="left",
            padx=5,
        )

        self.btn_reenviar = ttk.Button(
            frame_botoes_principais,
            text="↻ Reenviar Pendências",
            command=self.iniciar_reenvio,
        )

        self.btn_reenviar.pack(
            side="left",
            padx=5,
        )

        self.btn_retornos = ttk.Button(
            frame_botoes_principais,
            text="⬇ Baixar Retornos",
            command=self.iniciar_retornos,
        )

        self.btn_retornos.pack(
            side="left",
            padx=5,
        )

        self.btn_parar_execucao = ttk.Button(
            frame_botoes_principais,
            text="■ Parar Execução",
            command=self.parar_execucao,
            state="disabled",
        )

        self.btn_parar_execucao.pack(
            side="left",
            padx=5,
        )

        # --------------------------------------------------
        # BOTÕES DE ARQUIVOS
        # --------------------------------------------------

        frame_botoes_arquivos = ttk.LabelFrame(
            self.root,
            text="Arquivos e ações",
            padding=10,
        )

        frame_botoes_arquivos.pack(
            fill="x",
            padx=10,
            pady=5,
        )

        self.btn_abrir_log = ttk.Button(
            frame_botoes_arquivos,
            text="📄 Abrir Log",
            command=self.abrir_log,
        )

        self.btn_abrir_log.pack(
            side="left",
            padx=5,
        )

        self.btn_abrir_reenvio = ttk.Button(
            frame_botoes_arquivos,
            text="📋 Abrir Base de Reenvio",
            command=self.abrir_base_reenvio,
        )

        self.btn_abrir_reenvio.pack(
            side="left",
            padx=5,
        )

        self.btn_abrir_respostas = ttk.Button(
            frame_botoes_arquivos,
            text="📁 Abrir Pasta de Respostas",
            command=self.abrir_pasta_respostas,
        )

        self.btn_abrir_respostas.pack(
            side="left",
            padx=5,
        )

        self.btn_limpar = ttk.Button(
            frame_botoes_arquivos,
            text="🧹 Limpar Tela",
            command=self.limpar_log,
        )

        self.btn_limpar.pack(
            side="left",
            padx=5,
        )

        self.btn_fechar = ttk.Button(
            frame_botoes_arquivos,
            text="✕ Fechar Aplicação",
            command=self.fechar_aplicacao,
        )

        self.btn_fechar.pack(
            side="left",
            padx=5,
        )

        # --------------------------------------------------
        # BARRA DE PROGRESSO
        # --------------------------------------------------

        self.progress = ttk.Progressbar(
            self.root,
            mode="indeterminate",
        )

        self.progress.pack(
            fill="x",
            padx=10,
            pady=5,
        )

        # --------------------------------------------------
        # ÁREA DE LOG
        # --------------------------------------------------

        self.txt_log = ScrolledText(
            self.root,
            font=(
                "Consolas",
                10,
            ),
            wrap=tk.WORD,
        )

        self.txt_log.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10,
        )

        # --------------------------------------------------
        # MENSAGENS INICIAIS
        # --------------------------------------------------

        self.escrever_log(
            "Aplicação iniciada."
        )

        self.escrever_log(
            "Modo de teste inicial: "
            f"{'ATIVADO' if self.var_modo_teste.get() else 'DESATIVADO'}"
        )

        self.escrever_log(
            f"Arquivo de acompanhamento: {ARQUIVO_LOG}"
        )

        self.escrever_log(
            f"Base de reenvio: {ARQUIVO_REENVIO}"
        )

        self.escrever_log(
            f"Pasta de respostas: {PASTA_RESPOSTAS}"
        )

    # ======================================================
    # CONFIGURAÇÃO DO MODO DE TESTE
    # ======================================================

    def obter_email_outlook(self):
        """
        Obtém o endereço SMTP da conta padrão configurada no
        Outlook Classic.

        Returns
        -------
        str
            Endereço encontrado ou uma string vazia.
        """

        pythoncom.CoInitialize()

        try:
            try:
                outlook = win32.GetActiveObject(
                    "Outlook.Application"
                )

            except Exception:
                outlook = win32.Dispatch(
                    "Outlook.Application"
                )

            sessao = outlook.Session

            # A conta de entrega padrão é a melhor referência
            # para identificar quem está executando o aplicativo.
            try:
                conta_padrao = sessao.Accounts.Item(1)
                email_padrao = str(
                    getattr(
                        conta_padrao,
                        "SmtpAddress",
                        "",
                    )
                    or ""
                ).strip()

                if email_padrao:
                    return email_padrao

            except Exception:
                pass

            # Fallback: procura a primeira conta com endereço SMTP.
            for indice in range(
                1,
                sessao.Accounts.Count + 1,
            ):
                conta = sessao.Accounts.Item(indice)
                email = str(
                    getattr(
                        conta,
                        "SmtpAddress",
                        "",
                    )
                    or ""
                ).strip()

                if email:
                    return email

            return ""

        finally:
            pythoncom.CoUninitialize()

    def atualizar_email_outlook(self):
        """Atualiza manualmente a conta exibida na interface."""

        self.atualizar_modo_teste(
            exibir_erro=True,
        )

    def atualizar_modo_teste(
        self,
        exibir_erro=True,
    ):
        """
        Ativa ou desativa o modo de teste e, quando ativado,
        identifica automaticamente a conta padrão do Outlook.
        """

        if not self.var_modo_teste.get():
            self.var_email_teste.set(
                "Modo de teste desativado"
            )

            if hasattr(self, "btn_atualizar_email"):
                self.btn_atualizar_email.config(
                    state="disabled",
                )

            self.escrever_log(
                "Modo de teste DESATIVADO. Os próximos envios "
                "usarão os e-mails dos fornecedores."
            )

            return True

        if hasattr(self, "btn_atualizar_email"):
            self.btn_atualizar_email.config(
                state="normal",
            )

        try:
            email_identificado = self.obter_email_outlook()

        except Exception as erro:
            email_identificado = ""
            self.escrever_log(
                "Não foi possível identificar a conta do "
                f"Outlook: {erro}"
            )

        if email_identificado:
            self.var_email_teste.set(
                email_identificado
            )

            self.escrever_log(
                "Modo de teste ATIVADO. Destinatário de teste: "
                f"{email_identificado}"
            )

            return True

        self.var_email_teste.set(
            "Conta não identificada"
        )

        if exibir_erro:
            messagebox.showerror(
                "Conta do Outlook não identificada",
                (
                    "Não foi possível identificar automaticamente "
                    "o e-mail da conta configurada no Outlook Classic.\n\n"
                    "Abra o Outlook Classic, confirme que a conta está "
                    "conectada e clique em 'Atualizar conta'."
                ),
            )

        return False

    def obter_configuracao_envio(self):
        """
        Retorna e valida as configurações atuais da interface.
        """

        modo_teste = bool(
            self.var_modo_teste.get()
        )

        email_teste = ""

        if modo_teste:
            email_atual = self.var_email_teste.get().strip()

            if (
                not email_atual
                or "@" not in email_atual
            ):
                if not self.atualizar_modo_teste(
                    exibir_erro=True,
                ):
                    return None

                email_atual = self.var_email_teste.get().strip()

            email_teste = email_atual

        return {
            "modo_teste": modo_teste,
            "email_teste": email_teste,
        }

    # ======================================================
    # LOG DA INTERFACE
    # ======================================================

    def escrever_log(
        self,
        mensagem,
    ):
        """
        Agenda a escrita do log na thread principal
        do Tkinter.

        Essa função pode ser chamada pelas threads de
        envio, reenvio e consulta de retornos.
        """

        if self.aplicacao_fechando:
            return

        mensagem = str(
            mensagem
        )

        try:
            if (
                threading.current_thread()
                is threading.main_thread()
            ):
                self._inserir_log(
                    mensagem
                )

            else:
                self.root.after(
                    0,
                    self._inserir_log,
                    mensagem,
                )

        except tk.TclError:
            pass

    def _inserir_log(
        self,
        mensagem,
    ):
        """
        Insere efetivamente a mensagem na caixa de log.

        Esta função deve ser executada pela thread
        principal do Tkinter.
        """

        if self.aplicacao_fechando:
            return

        try:
            self.txt_log.insert(
                tk.END,
                f"{mensagem}\n",
            )

            self.txt_log.see(
                tk.END
            )

        except tk.TclError:
            pass

    def limpar_log(self):
        """
        Limpa as mensagens exibidas na interface.
        """

        try:
            self.txt_log.delete(
                "1.0",
                tk.END,
            )

            self.escrever_log(
                "Tela de log limpa."
            )

        except tk.TclError:
            pass

    # ======================================================
    # CONTROLE DA EXECUÇÃO
    # ======================================================

    def execucao_em_andamento(self):
        """
        Verifica se existe algum processo em execução.
        """

        return bool(
            self.thread_execucao
            and self.thread_execucao.is_alive()
        )

    def validar_execucao_disponivel(self):
        """
        Impede que dois processamentos sejam executados
        simultaneamente.
        """

        if self.execucao_em_andamento():
            messagebox.showwarning(
                "Execução em andamento",
                (
                    "Já existe um processamento em andamento.\n\n"
                    "Aguarde a conclusão ou clique em "
                    "'Parar Execução'."
                ),
            )

            return False

        return True

    def preparar_interface_execucao(
        self,
        texto_status,
    ):
        """
        Prepara a interface para o início de uma rotina.

        Os botões que podem iniciar outra execução são
        bloqueados e o botão de parada é habilitado.
        """

        # Limpa uma solicitação de parada de execução anterior.
        self.evento_parar.clear()

        self.lbl_status.config(
            text=texto_status,
            foreground="green",
        )

        self.progress.start(
            10
        )

        self.btn_tratar_excel.config(
            state="disabled",
        )

        self.btn_iniciar.config(
            state="disabled",
        )

        self.btn_reenviar.config(
            state="disabled",
        )

        self.btn_retornos.config(
            state="disabled",
        )

        self.btn_abrir_log.config(
            state="disabled",
        )

        self.btn_abrir_reenvio.config(
            state="disabled",
        )

        self.btn_abrir_respostas.config(
            state="disabled",
        )

        self.btn_parar_execucao.config(
            state="normal",
        )

        self.chk_modo_teste.config(
            state="disabled",
        )

        self.btn_atualizar_email.config(
            state="disabled",
        )

    def finalizar_interface_execucao(self):
        """
        Libera os componentes da interface depois da
        conclusão ou interrupção da rotina.
        """

        if self.aplicacao_fechando:
            return

        try:
            execucao_interrompida = (
                self.evento_parar.is_set()
            )

            self.progress.stop()

            if execucao_interrompida:
                self.lbl_status.config(
                    text="INTERROMPIDO",
                    foreground="orange",
                )

            else:
                self.lbl_status.config(
                    text="PARADO",
                    foreground="red",
                )

            self.btn_tratar_excel.config(
                state="normal",
            )

            self.btn_iniciar.config(
                state="normal",
            )

            self.btn_reenviar.config(
                state="normal",
            )

            self.btn_retornos.config(
                state="normal",
            )

            self.btn_abrir_log.config(
                state="normal",
            )

            self.btn_abrir_reenvio.config(
                state="normal",
            )

            self.btn_abrir_respostas.config(
                state="normal",
            )

            self.btn_parar_execucao.config(
                state="disabled",
            )

            self.chk_modo_teste.config(
                state="normal",
            )

            if self.var_modo_teste.get():
                self.btn_atualizar_email.config(
                    state="normal",
                )

            self.tipo_execucao = None
            self.thread_execucao = None

            # Limpa o evento somente depois que a execução
            # realmente terminou.
            self.evento_parar.clear()

        except tk.TclError:
            pass

    # ======================================================
    # PARADA DA EXECUÇÃO
    # ======================================================

    def parar_execucao(self):
        """
        Solicita a interrupção segura da execução atual.

        O aplicativo permanece aberto.

        A rotina em andamento terminará depois de concluir
        a etapa atual ou quando alcançar a próxima
        verificação do evento de parada.
        """

        if not self.execucao_em_andamento():
            messagebox.showinfo(
                "Nenhuma execução",
                "Não existe uma execução em andamento.",
            )

            return

        confirmado = messagebox.askyesno(
            "Parar execução",
            (
                "Deseja solicitar a interrupção da "
                "execução atual?\n\n"
                "O item que já estiver em processamento "
                "poderá ser concluído antes da parada.\n\n"
                "O aplicativo permanecerá aberto."
            ),
        )

        if not confirmado:
            return

        self.evento_parar.set()

        self.lbl_status.config(
            text="PARANDO...",
            foreground="orange",
        )

        self.btn_parar_execucao.config(
            state="disabled",
        )

        self.escrever_log(
            ""
        )

        self.escrever_log(
            "=" * 80
        )

        self.escrever_log(
            "SOLICITAÇÃO DE PARADA RECEBIDA"
        )

        self.escrever_log(
            "=" * 80
        )

        self.escrever_log(
            "Tipo de execução: "
            f"{self.tipo_execucao or 'NÃO IDENTIFICADO'}"
        )

        self.escrever_log(
            "Aguardando a conclusão segura da etapa atual..."
        )

    # ======================================================
    # TRATAMENTO DO EXCEL
    # ======================================================

    def iniciar_tratamento_excel(self):
        """Inicia o tratamento da base pelo utils/tratar_excel.py."""
        if not self.validar_execucao_disponivel():
            return

        confirmado = messagebox.askyesno(
            "Tratar Excel",
            (
                "Deseja localizar a base de pedidos mais recente e "
                "gerar o arquivo pedidos_tratados.xlsx?"
            ),
        )
        if not confirmado:
            return

        self.tipo_execucao = "TRATAMENTO_EXCEL"
        self.preparar_interface_execucao(
            texto_status="TRATANDO EXCEL",
        )
        self.escrever_log("")
        self.escrever_log("=" * 80)
        self.escrever_log("INICIANDO TRATAMENTO DO EXCEL")
        self.escrever_log("=" * 80)

        self.thread_execucao = threading.Thread(
            target=self.executar_tratamento_excel_thread,
            name="ThreadTratamentoExcel",
            daemon=True,
        )
        self.thread_execucao.start()

    def executar_tratamento_excel_thread(self):
        """Executa o tratamento sem bloquear a interface."""
        try:
            resultado = tratar_excel(
                callback_log=self.escrever_log,
                evento_parar=self.evento_parar,
            )

            if resultado is None:
                resultado = {}

            if resultado.get("interrompido", False):
                self.escrever_log("Tratamento interrompido pelo usuário.")
            elif resultado.get("sucesso", False):
                self.escrever_log("Base tratada com sucesso.")
                self.escrever_log(
                    f"Arquivo gerado: {resultado.get('arquivo_destino')}"
                )
                self.escrever_log(
                    f"Registros gerados: {resultado.get('registros', 0)}"
                )
            else:
                self.escrever_log("O tratamento não foi concluído com sucesso.")

        except Exception as erro:
            self.escrever_log(f"ERRO AO TRATAR EXCEL: {erro}")
            self.escrever_log(traceback.format_exc())
            self.mostrar_erro_thread(
                titulo="Erro no tratamento",
                mensagem=str(erro),
            )
        finally:
            self.escrever_log("=" * 80)
            if self.evento_parar.is_set():
                self.escrever_log("Tratamento do Excel interrompido.")
            else:
                self.escrever_log("Tratamento do Excel finalizado.")
            self.escrever_log("=" * 80)
            self.agendar_finalizacao_interface()

    # ======================================================
    # ENVIO NORMAL
    # ======================================================

    def iniciar_envios(self):
        """
        Solicita confirmação e inicia os envios normais.
        """

        if not self.validar_execucao_disponivel():
            return

        if not ARQUIVO_BASE_TRATADA.exists():
            messagebox.showwarning(
                "Base não tratada",
                (
                    "O arquivo pedidos_tratados.xlsx não foi encontrado.\n\n"
                    "Clique em 'Tratar Excel' antes de iniciar os envios."
                ),
            )
            return

        configuracao_envio = self.obter_configuracao_envio()

        if configuracao_envio is None:
            return

        self.configuracao_envio_atual = configuracao_envio

        confirmado = messagebox.askyesno(
            "Iniciar envios",
            (
                "Deseja iniciar o processamento normal "
                "dos fornecedores?\n\n"
                f"Modo de teste: "
                f"{'ATIVADO' if self.var_modo_teste.get() else 'DESATIVADO'}"
            ),
        )

        if not confirmado:
            return

        self.tipo_execucao = (
            "ENVIO_NORMAL"
        )

        self.preparar_interface_execucao(
            texto_status="EXECUTANDO ENVIOS",
        )

        self.escrever_log(
            ""
        )

        self.escrever_log(
            "=" * 80
        )

        self.escrever_log(
            "INICIANDO PROCESSAMENTO NORMAL"
        )

        self.escrever_log(
            "=" * 80
        )

        self.thread_execucao = threading.Thread(
            target=self.executar_envios_thread,
            name="ThreadEnvioNormal",
            daemon=True,
        )

        self.thread_execucao.start()

    def executar_envios_thread(self):
        """
        Executa os envios normais fora da thread principal
        da interface.
        """

        try:
            configuracao = getattr(
                self,
                "configuracao_envio_atual",
                self.obter_configuracao_envio(),
            )

            if configuracao is None:
                raise RuntimeError(
                    "Configuração de envio inválida."
                )

            parametros = {
                "callback_log": self.escrever_log,
                "evento_parar": self.evento_parar,
            }

            assinatura = inspect.signature(
                executar_envios
            )

            if "modo_teste" in assinatura.parameters:
                parametros["modo_teste"] = configuracao["modo_teste"]

            if "email_teste" in assinatura.parameters:
                parametros["email_teste"] = configuracao["email_teste"]

            resultado = executar_envios(
                **parametros
            )

            if resultado is None:
                resultado = {}

            if isinstance(
                resultado,
                dict,
            ):
                if resultado.get(
                    "interrompido",
                    False,
                ):
                    self.escrever_log(
                        "O processamento normal foi "
                        "interrompido pelo usuário."
                    )

                elif resultado.get(
                    "sucesso",
                    False,
                ):
                    self.escrever_log(
                        "O processo principal foi "
                        "concluído com sucesso."
                    )

                else:
                    codigo_retorno = resultado.get(
                        "codigo_retorno"
                    )

                    self.escrever_log(
                        "O processo principal não foi "
                        "concluído com sucesso."
                    )

                    if codigo_retorno is not None:
                        self.escrever_log(
                            "Código de retorno: "
                            f"{codigo_retorno}"
                        )

            elif resultado is True:
                self.escrever_log(
                    "O processo principal foi "
                    "concluído com sucesso."
                )

            elif resultado is False:
                self.escrever_log(
                    "O processo principal foi "
                    "finalizado com erro."
                )

        except Exception as erro:
            self.escrever_log(
                f"ERRO NO PROCESSAMENTO: {erro}"
            )

            self.escrever_log(
                traceback.format_exc()
            )

        finally:
            self.escrever_log(
                "=" * 80
            )

            if self.evento_parar.is_set():
                self.escrever_log(
                    "Processamento normal interrompido."
                )

            else:
                self.escrever_log(
                    "Processamento normal finalizado."
                )

            self.escrever_log(
                "=" * 80
            )

            self.agendar_finalizacao_interface()

    # ======================================================
    # REENVIO DE PENDÊNCIAS
    # ======================================================

    def iniciar_reenvio(self):
        """
        Solicita confirmação e inicia o processamento
        da base de reenvio.
        """

        if not self.validar_execucao_disponivel():
            return

        if not ARQUIVO_REENVIO.exists():
            messagebox.showwarning(
                "Base não encontrada",
                (
                    "A base de reenvio não foi encontrada.\n\n"
                    f"Arquivo esperado:\n{ARQUIVO_REENVIO}"
                ),
            )

            return

        configuracao_envio = self.obter_configuracao_envio()

        if configuracao_envio is None:
            return

        self.configuracao_envio_atual = configuracao_envio

        confirmado = messagebox.askyesno(
            "Reenviar pendências",
            (
                "Deseja processar os e-mails pendentes "
                "da base de reenvio?\n\n"
                f"Modo de teste: "
                f"{'ATIVADO' if self.var_modo_teste.get() else 'DESATIVADO'}"
            ),
        )

        if not confirmado:
            return

        self.tipo_execucao = (
            "REENVIO"
        )

        self.preparar_interface_execucao(
            texto_status="REENVIANDO",
        )

        self.escrever_log(
            ""
        )

        self.escrever_log(
            "=" * 80
        )

        self.escrever_log(
            "INICIANDO REENVIO DAS PENDÊNCIAS"
        )

        self.escrever_log(
            "=" * 80
        )

        self.escrever_log(
            f"Base de reenvio: {ARQUIVO_REENVIO}"
        )

        self.thread_execucao = threading.Thread(
            target=self.executar_reenvio_thread,
            name="ThreadReenvio",
            daemon=True,
        )

        self.thread_execucao.start()

    def executar_reenvio_thread(self):
        """
        Executa o processamento da base de reenvio.
        """

        logs_reenvio = []

        try:
            resumo = reenviar_emails_pendentes(
                arquivo_reenvio=ARQUIVO_REENVIO,
                pasta_fornecedores=PASTA_FORNECEDORES,
                imagem_claro=IMAGEM_CLARO,
                simbolo_pequeno=SIMBOLO_PEQUENO,
                logs_execucao=logs_reenvio,
                modo_teste=(
                    self.configuracao_envio_atual["modo_teste"]
                ),
                email_teste=(
                    self.configuracao_envio_atual["email_teste"]
                ),
                exibir_antes_de_enviar=(
                    EXIBIR_ANTES_DE_ENVIAR
                ),
                dias_para_retorno=(
                    DIAS_PARA_RETORNO
                ),
                evento_parar=self.evento_parar,
                callback_log=self.escrever_log,
            )

            if resumo is None:
                resumo = {}

            self.escrever_log(
                ""
            )

            self.escrever_log(
                "RESUMO DO REENVIO"
            )

            self.escrever_log(
                "-" * 80
            )

            self.escrever_log(
                "Tentativas processadas: "
                f"{resumo.get('processados', 0)}"
            )

            self.escrever_log(
                "E-mails reenviados: "
                f"{resumo.get('enviados', 0)}"
            )

            self.escrever_log(
                "E-mails não enviados: "
                f"{resumo.get('nao_enviados', 0)}"
            )

            self.escrever_log(
                "Abertos para revisão: "
                f"{resumo.get('abertos_revisao', 0)}"
            )

            self.escrever_log(
                "Pendências restantes: "
                f"{resumo.get('pendencias_restantes', 0)}"
            )

            if resumo.get(
                "interrompido",
                False,
            ):
                self.escrever_log(
                    "O reenvio foi interrompido "
                    "pelo usuário."
                )

            # --------------------------------------------------
            # SALVA OS LOGS GERADOS PELO REENVIO
            # --------------------------------------------------

            if logs_reenvio:
                self.escrever_log(
                    "Atualizando o histórico de envios..."
                )

                resultado_log = salvar_log_excel(
                    lista_logs=logs_reenvio,
                    arquivo_log=ARQUIVO_LOG,
                )

                if resultado_log:
                    self.escrever_log(
                        "Histórico atualizado com sucesso."
                    )

                else:
                    self.escrever_log(
                        "O histórico não foi atualizado."
                    )

            else:
                self.escrever_log(
                    "Nenhum novo registro de reenvio "
                    "foi gerado para o histórico."
                )

        except PermissionError as erro:
            self.escrever_log(
                f"ERRO DE PERMISSÃO NO REENVIO: {erro}"
            )

            self.mostrar_erro_thread(
                titulo="Arquivo aberto",
                mensagem=(
                    "Não foi possível atualizar um dos "
                    "arquivos do reenvio.\n\n"
                    "Feche os arquivos no Excel e "
                    "tente novamente."
                ),
            )

        except Exception as erro:
            self.escrever_log(
                f"ERRO NO REENVIO: {erro}"
            )

            self.escrever_log(
                traceback.format_exc()
            )

            # Caso parte dos reenvios tenha sido processada
            # antes do erro, tenta preservar os registros.
            if logs_reenvio:
                try:
                    self.escrever_log(
                        "Tentando salvar os registros "
                        "processados antes do erro..."
                    )

                    salvar_log_excel(
                        lista_logs=logs_reenvio,
                        arquivo_log=ARQUIVO_LOG,
                    )

                    self.escrever_log(
                        "Os registros processados antes "
                        "do erro foram salvos."
                    )

                except Exception as erro_log:
                    self.escrever_log(
                        "Erro ao salvar os registros "
                        f"do reenvio: {erro_log}"
                    )

        finally:
            self.escrever_log(
                "=" * 80
            )

            if self.evento_parar.is_set():
                self.escrever_log(
                    "Processamento de reenvio interrompido."
                )

            else:
                self.escrever_log(
                    "Processamento de reenvio finalizado."
                )

            self.escrever_log(
                "=" * 80
            )

            self.agendar_finalizacao_interface()

    # ======================================================
    # CONSULTA DE RETORNOS
    # ======================================================

    def iniciar_retornos(self):
        """
        Solicita confirmação e inicia a consulta das
        respostas no Outlook.
        """

        if not self.validar_execucao_disponivel():
            return

        if not ARQUIVO_LOG.exists():
            messagebox.showwarning(
                "Arquivo não encontrado",
                (
                    "O arquivo de acompanhamento não foi "
                    "encontrado.\n\n"
                    f"Arquivo esperado:\n{ARQUIVO_LOG}"
                ),
            )

            return

        confirmado = messagebox.askyesno(
            "Baixar retornos",
            (
                "Deseja consultar no Outlook as respostas "
                "dos fornecedores?\n\n"
                "A rotina pesquisará os IDs da aba "
                "'Aguardando_Retorno', baixará os anexos "
                "Excel e atualizará o histórico.\n\n"
                "O arquivo de acompanhamento deve estar "
                "fechado no Excel."
            ),
        )

        if not confirmado:
            return

        self.tipo_execucao = (
            "RETORNOS"
        )

        self.preparar_interface_execucao(
            texto_status="CONSULTANDO RETORNOS",
        )

        self.escrever_log(
            ""
        )

        self.escrever_log(
            "=" * 80
        )

        self.escrever_log(
            "INICIANDO CONSULTA DE RETORNOS"
        )

        self.escrever_log(
            "=" * 80
        )

        self.escrever_log(
            f"Arquivo de acompanhamento: {ARQUIVO_LOG}"
        )

        self.escrever_log(
            f"Pasta de respostas: {PASTA_RESPOSTAS}"
        )

        self.thread_execucao = threading.Thread(
            target=self.executar_retornos_thread,
            name="ThreadRetornos",
            daemon=True,
        )

        self.thread_execucao.start()

    def executar_retornos_thread(self):
        """
        Pesquisa as respostas no Outlook, baixa os anexos
        e atualiza os status no histórico.
        """

        try:
            resumo = processar_retornos_fornecedores(
                arquivo_log=ARQUIVO_LOG,
                pasta_respostas=PASTA_RESPOSTAS,
                callback_log=self.escrever_log,
                evento_parar=self.evento_parar,
            )

            if resumo is None:
                resumo = {}

            self.escrever_log(
                ""
            )

            self.escrever_log(
                "RESUMO DOS RETORNOS"
            )

            self.escrever_log(
                "-" * 80
            )

            self.escrever_log(
                "IDs verificados: "
                f"{resumo.get('verificados', 0)}"
            )

            self.escrever_log(
                "Respondidos com anexo: "
                f"{resumo.get('respondidos_com_anexo', 0)}"
            )

            self.escrever_log(
                "Respondidos sem anexo: "
                f"{resumo.get('respondidos_sem_anexo', 0)}"
            )

            self.escrever_log(
                "Respostas não localizadas: "
                f"{resumo.get('nao_localizados', 0)}"
            )

            self.escrever_log(
                "Erros encontrados: "
                f"{resumo.get('erros', 0)}"
            )

            if resumo.get(
                "interrompido",
                False,
            ):
                self.escrever_log(
                    "A consulta de respostas foi "
                    "interrompida pelo usuário."
                )

        except PermissionError as erro:
            self.escrever_log(
                f"ERRO DE PERMISSÃO: {erro}"
            )

            self.escrever_log(
                "Feche o arquivo de acompanhamento no Excel "
                "e tente novamente."
            )

            self.mostrar_erro_thread(
                titulo="Arquivo aberto",
                mensagem=(
                    "Não foi possível atualizar o arquivo "
                    "de acompanhamento.\n\n"
                    "Feche o arquivo no Excel e "
                    "tente novamente."
                ),
            )

        except Exception as erro:
            self.escrever_log(
                f"ERRO AO PROCESSAR RETORNOS: {erro}"
            )

            self.escrever_log(
                traceback.format_exc()
            )

        finally:
            self.escrever_log(
                "=" * 80
            )

            if self.evento_parar.is_set():
                self.escrever_log(
                    "Consulta de retornos interrompida."
                )

            else:
                self.escrever_log(
                    "Consulta de retornos finalizada."
                )

            self.escrever_log(
                "=" * 80
            )

            self.agendar_finalizacao_interface()

    # ======================================================
    # FUNÇÕES SEGURAS PARA THREAD
    # ======================================================

    def agendar_finalizacao_interface(self):
        """
        Agenda a liberação dos componentes na thread
        principal do Tkinter.
        """

        if self.aplicacao_fechando:
            return

        try:
            self.root.after(
                0,
                self.finalizar_interface_execucao,
            )

        except tk.TclError:
            pass

    def mostrar_erro_thread(
        self,
        titulo,
        mensagem,
    ):
        """
        Exibe uma caixa de erro por meio da thread
        principal do Tkinter.
        """

        if self.aplicacao_fechando:
            return

        try:
            self.root.after(
                0,
                messagebox.showerror,
                titulo,
                mensagem,
            )

        except tk.TclError:
            pass

    # ======================================================
    # ABERTURA DE ARQUIVOS E PASTAS
    # ======================================================

    def abrir_log(self):
        """
        Abre o arquivo de acompanhamento no Excel.
        """

        try:
            if not ARQUIVO_LOG.exists():
                messagebox.showwarning(
                    "Arquivo não encontrado",
                    (
                        "O arquivo de acompanhamento não "
                        "foi encontrado.\n\n"
                        f"Arquivo esperado:\n{ARQUIVO_LOG}"
                    ),
                )

                return

            os.startfile(
                str(ARQUIVO_LOG)
            )

        except Exception as erro:
            messagebox.showerror(
                "Erro ao abrir o log",
                str(erro),
            )

    def abrir_base_reenvio(self):
        """
        Abre a base de reenvio no Excel.
        """

        try:
            if not ARQUIVO_REENVIO.exists():
                messagebox.showwarning(
                    "Arquivo não encontrado",
                    (
                        "A base de reenvio não foi encontrada.\n\n"
                        f"Arquivo esperado:\n{ARQUIVO_REENVIO}"
                    ),
                )

                return

            os.startfile(
                str(ARQUIVO_REENVIO)
            )

        except Exception as erro:
            messagebox.showerror(
                "Erro ao abrir a base de reenvio",
                str(erro),
            )

    def abrir_pasta_respostas(self):
        """
        Abre a pasta onde os anexos recebidos são salvos.
        """

        try:
            PASTA_RESPOSTAS.mkdir(
                parents=True,
                exist_ok=True,
            )

            os.startfile(
                str(PASTA_RESPOSTAS)
            )

        except Exception as erro:
            messagebox.showerror(
                "Erro ao abrir a pasta de respostas",
                str(erro),
            )

    # ======================================================
    # FECHAMENTO DA APLICAÇÃO
    # ======================================================

    def fechar_aplicacao(self):
        """
        Fecha a aplicação.

        Caso exista uma execução ativa, solicita confirmação
        e sinaliza a parada antes de fechar a interface.
        """

        if self.execucao_em_andamento():
            confirmado = messagebox.askyesno(
                "Execução em andamento",
                (
                    "Existe um processamento em andamento.\n\n"
                    "O recomendado é clicar em "
                    "'Parar Execução' e aguardar o término "
                    "seguro da rotina.\n\n"
                    "Deseja fechar a aplicação mesmo assim?"
                ),
            )

            if not confirmado:
                return

            self.evento_parar.set()

        self.aplicacao_fechando = True

        try:
            self.root.destroy()

        except tk.TclError:
            pass

    # ======================================================
    # START
    # ======================================================

    def run(self):
        """
        Inicia o loop principal do Tkinter.
        """

        self.root.mainloop()


if __name__ == "__main__":
    AppEnvios().run()