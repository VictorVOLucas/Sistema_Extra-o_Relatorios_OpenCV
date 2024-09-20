from widgetsCustom import Widgets  # Importa a classe Widgets do arquivo widgets.py
import customtkinter as ctk  # Módulo para criar interfaces gráficas

if __name__ == "__main__": # Função principal
    app = ctk.CTk() # Cria uma instância da classe CTk
    app.title("Sistema de Extração de Relatórios") # Define o título da janela
    widgets = Widgets(app) # Cria uma instância da classe Widgets
    app.mainloop() # Inicia o loop principal da aplicação
