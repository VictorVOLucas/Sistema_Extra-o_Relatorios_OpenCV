import customtkinter as ctk # Módulo para criar interfaces gráficas
import schedule  # Biblioteca para agendamento de tarefas
import time  # Módulo para manipulação de tempo
import os  # Módulo para interação com o sistema operacional
from threading import Thread  # Classe para execução de tarefas em threads separadas
import json  # Módulo para manipulação de arquivos JSON
from utils import Utils
from RoboV4 import RoboV4
import psutil  # Para encerrar os processos do sistema
import sys # Módulo para interação com o sistema


class Widgets:
# Classe para criar os widgets da aplicação
    def __init__(self, root): # Método construtor
        self.root = root # Define a janela principal
        self.app = root 
        self.root.title("Sistema de Extração de Relatórios")  # Define o título da janela principal
        self.root.geometry("1360x700")  # Define o tamanho da janela
        self.tab_frames = {}  # Dicionário para armazenar as abas
        self.time_entries = []  # Inicializa a lista de campos de entrada de horários
        self.entries = {}  # Inicializa o dicionário entries
        self.jobs = []  # Inicializa a lista de tarefas agendadas

        self.create_tabs()  # Cria as abas

        self.is_running = False  # Flag para indicar se o agendamento está ativo

        self.log_text = ctk.CTkTextbox(self.root, height=600, width=600, border_width=5, state="disable")  # Campo de texto para exibir o log de eventos
        self.log_text.pack(pady=5, padx=5, fill="both", expand=True)   # Empacota o campo de texto

        sys.stdout = TextRedirector(self.log_text) # Redireciona o stdout para o campo de texto
        sys.stderr = TextRedirector(self.log_text) # Redireciona o stderr para o campo de texto

    def create_tabs(self): # Método para criar as abas
        self.tab_buttons_frame = ctk.CTkFrame(self.root) # Cria um frame para os botões das abas
        self.tab_buttons_frame.pack(side="left", fill="y", expand=False, padx=10, pady=10, anchor="nw", ipadx=10, ipady=10) # Empacota o frame

        self.tab_content_frame = ctk.CTkFrame(self.root) 
        self.tab_content_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10, anchor="ne", ipadx=10, ipady=10)

        self.tab_frames = {}

        ctk.CTkLabel(self.tab_buttons_frame, text="MENU", font=("Arial", 20, "bold")).pack(pady=20)  # Título da aplicação
        self.add_tab("Inicio", self.aba_inicio) # Adiciona a aba Início
        self.add_tab("Configurações", self.aba_configuracoes)
        self.add_tab("Login", self.aba_login)
        self.add_tab("Nomes Arquivos", self.aba_nomes_arquivos)
        self.add_tab("Telegram", self.aba_Telegram)

        self.show_tab("Inicio") # Exibe a aba Início por padrão

    def add_tab(self, name, create_func): # Método para adicionar uma aba
        button = ctk.CTkButton(self.tab_buttons_frame, text=name, height=40, border_width=2, command=lambda: self.show_tab(name)) # Cria um botão para a aba
        button.pack(fill="x", pady=10) # Empacota o botão

        frame = ctk.CTkFrame(self.tab_content_frame) # Cria um frame para a aba
        self.tab_frames[name] = frame
        create_func(frame)

    def show_tab(self, name): # Método para exibir a aba selecionada
        for frame in self.tab_frames.values(): # Percorre todos os frames
            frame.pack_forget() # Oculta todas as abas
        self.tab_frames[name].pack(fill="both", expand=True) # Exibe a aba selecionada

    def aba_inicio(self, parent): # Método para criar a aba Início
        # ================= ABA INICIAL ================#
        ctk.CTkLabel(parent, text="Horários para extração de relatórios (HH:MM)", font=("Arial", 15, "bold")).pack(pady=5)  # Rótulo de instrução

        for i in range(4):  # Cria quatro campos de entrada para horários
            entry = ctk.CTkEntry(parent)
            entry.pack(pady=5)
            self.time_entries.append(entry)

        self.load_schedule() # Carrega os horários agendados

        self.save_button = ctk.CTkButton(parent, text="Salvar Agendamentos", command=lambda: self.save_schedule())  # Botão para salvar os horários agendados
        self.save_button.pack(pady=20)

        self.stop_button = ctk.CTkButton(parent, text="Encerrar Sistema", command=self.stop_system)  # Botão para encerrar o sistema
        self.stop_button.pack(pady=10)

        
        # ================= FIM DA ABA INICIAL ================#

    def aba_configuracoes(self, parent): # Método para criar a aba Configurações
        # ================= ABA PARA CONFIGURAÇÕES CAMINHO PARA AS PASTAS ================#
        ctk.CTkLabel(parent, text="CONFIGURAÇÕES DA APLICAÇÃO").grid(row=0, column=0, columnspan=2, pady=5)

        # Dicionário para armazenar os campos de entrada
        self.fields_config = [ 
            {"label": "Caminho do Sistema:", "entry": "software_path_entry", "row": 1},
            {"label": "Argumentos para o executável:", "entry": "arguments_entry", "row": 2},
            {"label": "Nome das Sheets planilha MRP:", "entry": "name_sheets_entry", "row": 3},
            {"label": "Caminho do MRP pelo desktop:", "entry": "caminho_mrp_desktop_entry", "row": 4},
            {"label": "Caminho KORP relatórios:", "entry": "caminho_entry", "row": 5},
            {"label": "Caminho KORP parâmetros:", "entry": "caminho_parametros_entry", "row": 6},
            {"label": "Caminho KORP Relatório Movimento Saída de ODF:", "entry": "caminho_mov_odf_entry", "row": 7},
            {"label": "Caminho KORP Relatório Faturamento Peso:", "entry": "caminho_faturamento_peso_entry", "row": 8},
            {"label": "Caminho KORP Relatório do MRP:", "entry": "caminho_mrp_entry", "row": 9},
        ]

        for field in self.fields_config: # Cria os campos de entrada para as configurações
            ctk.CTkLabel(parent, text=field["label"]).grid(row=field["row"], column=0, sticky='e', pady=5, padx=5) # Rótulo do campo de entrada
            entry = ctk.CTkEntry(parent, width=250) # Cria um campo de entrada
            entry.grid(row=field["row"], column=1, pady=5, padx=5) # Empacota o campo de entrada
            self.entries[field["entry"]] = entry # Adiciona o campo de entrada ao dicionário

        save_button = ctk.CTkButton(parent, text="Salvar Caminho e Argumentos", command=lambda: Utils.salvar_caminho_relatorios(self.entries))  # Botão para salvar os caminhos e argumentos
        save_button.grid(row=10, column=0, columnspan=2, pady=10)

        try:
            with open('configuracoes_software.json', 'r') as f:
                config = json.load(f)

                self.entries["software_path_entry"].insert(0, config.get("software_path", ""))
                self.entries["arguments_entry"].insert(0, config.get("arguments", ""))
                self.entries["name_sheets_entry"].insert(0, config.get("name_sheets", ""))
                self.entries["caminho_mrp_desktop_entry"].insert(0, config.get("caminho_mrp_desktop", ""))
                self.entries["caminho_entry"].insert(0, config.get("caminho", ""))
                self.entries["caminho_parametros_entry"].insert(0, config.get("caminho_parametros", ""))
                self.entries["caminho_mov_odf_entry"].insert(0, config.get("caminho_mov_odf", ""))
                self.entries["caminho_faturamento_peso_entry"].insert(0, config.get("caminho_faturamento_peso", ""))
                self.entries["caminho_mrp_entry"].insert(0, config.get("caminho_mrp", ""))
        except FileNotFoundError:
            Utils.show_message("Aviso", "Arquivo configuracoes_software.json não encontrado.")

    def aba_nomes_arquivos(self, parent):
        # ================= ABA PARA CONFIGURAÇÕES NOME DAS PASTAS ================#

        ctk.CTkLabel(parent, text="DEFINA O NOME DOS ARQUIVOS").grid(row=0, column=0, columnspan=2, pady=5)

        self.fields_arquivos = [
            {"label": "Estoque Supper:", "entry": "estoque_supper_entry", "row": 1},
            {"label": "Faturamento Peso:", "entry": "faturamento_peso_entry", "row": 2},
            {"label": "Pedido Compra:", "entry": "pedido_compra_entry", "row": 3},
            {"label": "Requisição Almoxarifado:", "entry": "req_almoxarifado_entry", "row": 4},
            {"label": "Movimento Saida de ODF:", "entry": "mov_saida_odf_entry", "row": 5},
            {"label": "MRP Filial - 2:", "entry": "mrp_filial_dois_entry", "row": 6},
            {"label": "JMS - 1:", "entry": "mrp_jms_entry", "row": 7},
            {"label": "JM - 3:", "entry": "mrp_jm_entry", "row": 8},
            {"label": "Parâmetros:", "entry": "parametros_entry", "row": 9},
        ]

        for field in self.fields_arquivos:
            ctk.CTkLabel(parent, text=field["label"]).grid(row=field["row"], column=0, sticky='e', pady=5, padx=5)
            entry = ctk.CTkEntry(parent, width=250)
            entry.grid(row=field["row"], column=1, pady=5, padx=5)
            self.entries[field["entry"]] = entry

        save_button = ctk.CTkButton(parent, text="Salvar Nomes", command=lambda: Utils.salvar_nome_arquivos(self.entries))
        save_button.grid(row=10, column=0, columnspan=2, pady=10)

        try:
            with open('configuracoes_nomes.json', 'r') as f: # Carrega os nomes dos arquivos do arquivo JSON
                config = json.load(f)

                self.entries["estoque_supper_entry"].insert(0, config.get("estoque_supper", "")) # Insere o nome do arquivo no campo de entrada
                self.entries["faturamento_peso_entry"].insert(0, config.get("faturamento_peso", ""))
                self.entries["pedido_compra_entry"].insert(0, config.get("pedido_compra", ""))
                self.entries["req_almoxarifado_entry"].insert(0, config.get("req_almoxarifado", ""))
                self.entries["mov_saida_odf_entry"].insert(0, config.get("mov_saida_odf", ""))
                self.entries["mrp_filial_dois_entry"].insert(0, config.get("mrp_filial_dois", ""))
                self.entries["mrp_jms_entry"].insert(0, config.get("mrp_jms", ""))
                self.entries["mrp_jm_entry"].insert(0, config.get("mrp_jm", ""))
                self.entries["parametros_entry"].insert(0, config.get("parametros", ""))
        except FileNotFoundError:
            Utils.show_message("Aviso", "Arquivo configuracoes_nomes.json não encontrado.")

        # ================= FIM DA ABA PARA CONFIGURAÇÕES NOME DAS PASTAS ================#

    def aba_login(self, parent): # Método para criar a aba Login
        # ================= ABA PARA CONFIGURAÇÕES DE LOGIN ================#
        ctk.CTkLabel(parent, text="INFORME O LOGIN DO SISTEMA").grid(row=0, column=0, columnspan=2, pady=5)

        self.fields_login = [
            {"label": "Usuario:", "entry": "usuario_entry", "row": 1}, # Dicionário para armazenar os campos de entrada
            {"label": "Senha:", "entry": "senha_entry", "row": 2},
        ]

        for field in self.fields_login:
            ctk.CTkLabel(parent, text=field["label"]).grid(row=field["row"], column=0, sticky='e', pady=5, padx=5)
            entry = ctk.CTkEntry(parent, width=250)
            entry.grid(row=field["row"], column=1, pady=5, padx=5)
            self.entries[field["entry"]] = entry

        save_button = ctk.CTkButton(parent, text="Salvar Login", command=lambda: Utils.salvar_login(self.entries))
        save_button.grid(row=10, column=0, columnspan=2, pady=10)

        try:
            with open('configuracoes_login.json', 'r') as f: # Carrega as configurações de login do arquivo JSON
                config = json.load(f)

                self.entries["usuario_entry"].insert(0, config.get("Usuario", "")) # Insere o usuário no campo de entrada
                self.entries["senha_entry"].insert(0, config.get("Senha", ""))
        except FileNotFoundError:
            Utils.show_message("Aviso", "Arquivo configuracoes_login.json não encontrado.")

    def aba_Telegram(self, parent):
        # ================= ABA PARA CONFIGURAÇÕES DO TELEGRAM ================#
        ctk.CTkLabel(parent, text="INFORME AS INFORMAÇÕES DE API PARA O TELEGRAM").grid(row=0, column=0, columnspan=2, pady=5)

        self.fields_dados = [
            {"label": "Token:", "entry": "Token_entry", "row": 1}, # Dicionário para armazenar os campos de entrada
            {"label": "ChatID:", "entry": "ChatID_entry", "row": 2},
            {"label": "Diretorio:", "entry": "Diretorio_entry", "row": 3},
        ]

        for field in self.fields_dados:
            ctk.CTkLabel(parent, text=field["label"]).grid(row=field["row"], column=0, sticky='e', pady=5, padx=5)
            entry = ctk.CTkEntry(parent, width=250)
            entry.grid(row=field["row"], column=1, pady=5, padx=5)
            self.entries[field["entry"]] = entry

        save_button = ctk.CTkButton(parent, text="Salvar API", command=lambda: Utils.salvar_telegram(self.entries))
        save_button.grid(row=10, column=0, columnspan=2, pady=10)

        try:
            with open('configuracoes_telegram.json', 'r') as f: # Carrega as configurações de login do arquivo JSON
                config = json.load(f)

                self.entries["Token_entry"].insert(0, config.get("Token", "")) # Insere o usuário no campo de entrada
                self.entries["ChatID_entry"].insert(0, config.get("ChatID", ""))
                self.entries["Diretorio_entry"].insert(0, config.get("Diretorio", ""))
        except FileNotFoundError:
            Utils.show_message("Aviso", "Arquivo configuracoes_telegram.json não encontrado.")

    def load_schedule(self): # Método para carregar os horários agendados
        try:
            with open("agendamentos.json", "r") as file: # Abre o arquivo JSON
                data = json.load(file)  # Lê os horários agendados do arquivo JSON
                times = data.get('times', [])

                # Garante que o número de horários não exceda o número de campos de entrada
                for i, time_entry in enumerate(self.time_entries):
                    if i < len(times):
                        time_entry.delete(0, ctk.END)  # Limpa o campo de entrada
                        time_entry.insert(ctk.END, times[i])  # Insere os horários nos campos de entrada
                    else:
                        time_entry.delete(0, ctk.END)  # Limpa qualquer campo extra caso não haja horários suficientes
        except (FileNotFoundError, json.JSONDecodeError):
            pass  #

    def save_schedule(self): # Método para salvar os horários agendados
        times = [entry.get() for entry in self.time_entries]  # Obtém os horários dos campos de entrada
        if all(self.validate_time_format(t) for t in times):  # Verifica se todos os horários são válidos
            data = {'times': times} # Cria um dicionário com os horários
            with open("agendamentos.json", "w") as file:
                json.dump(data, file, indent=4)  # Salva os horários no arquivo JSON
            self.start_schedule(times)  # Inicia o agendamento
        else:
            Utils.show_message("Erro", "Por favor, insira horários válidos no formato HH:MM.")  # Exibe mensagem de erro

    def validate_time_format(self, time_str): # Método para validar o formato do horário
        try:
            time.strptime(time_str, "%H:%M")  # Verifica se o horário está no formato HH:MM
            return True
        except ValueError:
            return False

    def start_schedule(self, times): # Método para iniciar o agendamento
        if not self.is_running:
            self.is_running = True  # Define o flag como True
            # self.status_label.config(text="Status: Ativo", fg="green")  # Atualiza o status para ativo
            print("Agendamento iniciado.")  # Loga a mensagem de início do agendamento
            Thread(target=self.run_schedule, args=(times,)).start()  # Inicia a execução do agendamento em uma thread separada
        else:
            Utils.show_message("Info", "O agendamento já está em execução.")  # Informa que o agendamento já está em execução

    def run_schedule(self, times): # Método para executar o agendamento
        try:
            for job in self.jobs:
                schedule.cancel_job(job) # Cancela todas as tarefas agendadas
            self.jobs.clear()

            robo = RoboV4()  # Cria uma instância de RoboV4 uma vez

            for time_str in times:
                schedule.every().day.at(time_str).do(lambda: robo.extrair_relatorios())  # Usa a instância para agendar a extração dos relatórios

            while self.is_running: # Loop para manter o agendamento ativo
                schedule.run_pending()  # Executa tarefas agendadas
                time.sleep(1)
        except Exception as e:
            Utils.show_message(f"Ocorreu um erro ao iniciar o agendamento: {e}")  # Loga qualquer erro ocorrido durante a execução
            raise

    def stop_system(self): # Método para encerrar o sistema
        self.is_running = False
        self.terminate_all_processes() # Encerra todos os processos relacionados ao sistema

    def terminate_all_processes(self): 
        """Encerra todos os processos relacionados ao sistema usando psutil."""
        current_process = psutil.Process(os.getpid())  # Obtém o ID do processo atual
        children = current_process.children(recursive=True)  # Obtém todos os processos filhos
        for child in children:
            # print(f"Encerrando processo filho: {child.pid}")
            child.terminate()  # Tenta encerrar o processo filho de forma segura
        psutil.wait_procs(children, timeout=5)  # Aguarda o término dos processos filhos
        current_process.terminate()  # Encerra o processo principal
        # print("Sistema encerrado com sucesso.")

    def load_times_from_json(self, filename): # Método para carregar os horários de um arquivo JSON
        try:
            with open(filename, 'r') as file: # Abre o arquivo JSON
                data = json.load(file)
                return data.get('times', [])
        except (FileNotFoundError, json.JSONDecodeError):
            return []


class TextRedirector: # Classe para redirecionar a saída padrão para um widget
    def __init__(self, widget):
        """Redireciona o stdout e stderr para o widget log_text."""
        self.widget = widget

    def write(self, message):
        # Insere a mensagem no widget
        self.widget.configure(state="normal")  # Ativa a edição
        self.widget.insert(ctk.END, message)  # Insere a mensagem no final do widget
        self.widget.see(ctk.END)  # Rola para a última linha
        self.widget.configure(state="disabled")  # Desativa a edição

    def flush(self):
        pass