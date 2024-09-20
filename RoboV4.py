import subprocess # Importa a biblioteca subprocess para abrir o software
import pyautogui # Importa a biblioteca pyautogui para realizar a automação
import time     # Importa a biblioteca time para adicionar pausas
from utils import Utils, configurar_logs # Importa a classe Utils e a função configurar_logs
from datetime import datetime # Importa a classe datetime para obter a data e hora atuais
import json # Importa a biblioteca json para manipular arquivos JSON
import logging # Importa a biblioteca logging para gerar logs
from Enviar_Log import EnviarLogs # Importa a classe EnviarLogs para enviar logs para o Telegram

class RoboV4: # Classe principal do RoboV4
    
    @staticmethod
    def abrir_software(): # Método para abrir o software RoboV4
        # Ler o caminho e os argumentos do arquivo de texto
        try:
            with open('configuracoes_software.json', 'r') as f: # Abre o arquivo de configurações
                config = json.load(f)
                
                software_path = config.get("software_path", "")
                arguments = config.get("arguments", "").split(',')  # Supondo que os argumentos estejam separados por vírgula ou outro separador

                if not software_path or not arguments:              
                    Utils.show_message("Erro", f"Arquivo de configurações incompleto: {str(e)}")
                    return

                # Argumentos para o executável
                full_arguments = [software_path] + arguments

                subprocess.Popen(full_arguments, creationflags=subprocess.CREATE_NO_WINDOW)
                logging.info(f"O software foi aberto com sucesso! Caminho do Software: {software_path} \n Argumentos: {arguments}")
                time.sleep(25)  # Aguardar o software abrir completamente
        except Exception as e:
            logging.error(f"Ocorreu um erro ao abrir o software: {e}")
            raise
        
    @staticmethod
    def realizar_login(): # Método para realizar o login no sistema
        try:

            with open('configuracoes_login.json', 'r') as f: # Abre o arquivo de configurações
                config = json.load(f)
                
                usuario = config.get("Usuario", "")
                senha = config.get("Senha", "")  # Supondo que os argumentos estejam separados por vírgula ou outro separador

                Utils.clicar_elemento('img/Logar_Sistema/Senha.png', "Campo 'Senha'")
                pyautogui.write(senha)
                time.sleep(2)  # Ajuste conforme necessário

                Utils.clicar_elemento('img/Logar_Sistema/BotaoEntrar.png', "Botão 'Entrar'")
                time.sleep(10)  # Aguardar a resposta do login

                if not usuario or not senha:
                    Utils.show_message("Erro", f"Arquivo de configurações login incompleto: {str(e)}")
                    return

        except Exception as e:
            logging.error(f"Ocorreu um erro ao realizar login: {e}")
            raise

    @staticmethod
    def abrir_cubo_de_decisao(): # Método para abrir o Cubo de Decisão
        try:
            Utils.clicar_elemento('img/Menu_Relatorio/Relatorios.png', "Menu 'Relatórios'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Menu_Relatorio/Cubo.png', "Submenu 'Cubo de decisão'")
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            logging.error(f"Ocorreu um erro ao abrir o Cubo de Decisão: {e}")
            raise

    @staticmethod
    def relatorio_faturamento_peso(): # Método para gerar o relatório de Faturamento Peso
        try:
            with open('configuracoes_nomes.json', 'r') as f:
                config = json.load(f)
                
                nome_arquivo = config.get("faturamento_peso", "")

            Utils.clicar_elemento('img/Rel_FaturamentoPeso/FaturamentoPeso.png', "Campo 'Faturamento Peso'")
            time.sleep(10)  # Ajuste conforme necessário

            Utils.aplicar_data_mes_atual('img/Rel_FaturamentoPeso/DataInicial.png', 'img/Rel_FaturamentoPeso/DataFinal.png')
            Utils.apicar_modelo_de_relatorio('img/Rel_FaturamentoPeso/FaturamentoPeso_Lista.png')
            Utils.clicar_elemento('img/Rel_FaturamentoPeso/Expand_Notes/CodProd.png', "Botão 'CodProd'")
            Utils.expand_All_Notes('img/Rel_FaturamentoPeso/Expand_Notes/CodCli.png')

            Utils.salvar_arquivo_faturamento_peso(nome_arquivo)

        except Exception as e:
            logging.error(f"Ocorreu um erro ao gerar o relatório de Faturamento Peso: {e}")
            raise

    @staticmethod
    def relatorio_estoque_supper(): # Método para gerar o relatório de Estoque Supper
        try:
            with open('configuracoes_nomes.json', 'r') as f:
                config = json.load(f)
                
                nome_arquivo = config.get("estoque_supper", "")

            Utils.clicar_elemento('img/Rel_EstoqueSupper/EstoqueSupper.png', "Campo 'Estoque Supper'")
            time.sleep(10)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_EstoqueSupper/EstoqueSupper_Lista.png')
            Utils.salvar_arquivo(nome_arquivo)
            time.sleep(10)  # Ajuste conforme necessário

        except Exception as e:
            logging.error(f"Ocorreu um erro ao gerar o relatório de Estoque Supper: {e}")
            raise

    @staticmethod
    def relatorio_pedido_compra(): # Método para gerar o relatório de Pedido Compra
        try:
            with open('configuracoes_nomes.json', 'r') as f:
                config = json.load(f)
                
                nome_arquivo = config.get("pedido_compra", "")

            Utils.clicar_elemento('img/Rel_PedidoCompra/PedidoCompra.png', "Campo 'Pedido Compra'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_PedidoCompra/PedidoCompra_Lista.png')
            Utils.salvar_arquivo(nome_arquivo)
            # Utils.salvar_arquivo_pedido_compra(RoboV4.nomeArquivo_PedidoCompra)
            time.sleep(10)  # Ajuste conforme necessário

        except Exception as e:
            logging.error(f"Ocorreu um erro ao gerar o relatório de Pedido Compra: {e}")
            raise

    @staticmethod
    def abrir_mrp(): # Método para abrir o MRP
        try:
 
            with open('configuracoes_nomes.json', 'r') as f:
                config = json.load(f)
                
                nome_arquivo_filial2 = config.get("mrp_filial_dois", "")
                nome_arquivo_jms1 = config.get("mrp_jms", "")
                nome_arquivo_jm = config.get("mrp_jm", "")

            Utils.clicar_elemento('img/Menu_Logistica/Logistica.png', "Menu 'Logística'")
            time.sleep(10)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Menu_Logistica/MRP.png', "Submenu 'MRP'")
            time.sleep(10)  # Ajuste conforme necessário

            Utils.salvar_arquivo_mrp(nome_arquivo_filial2)

            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta trocar empresa'")
            time.sleep(5)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/JMS_1.png', "Botão 'Empresa JMS 1'")
            time.sleep(10)  # Ajuste conforme necessário

            Utils.salvar_arquivo_mrp(nome_arquivo_jms1)

            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta trocar empresa'")
            time.sleep(5)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/JM_3.png', "Botão 'Empresa JM 3'")
            time.sleep(10)  # Ajuste conforme necessário

            Utils.salvar_arquivo_mrp(nome_arquivo_jm)

            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta trocar empresa'")
            time.sleep(5)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Supper_Filial_2.png', "Botão 'Empresa Supper Filial 2'")
            time.sleep(5)  # Ajuste conforme necessário

            with open('configuracoes_software.json', 'r') as f:
                config = json.load(f)
                caminho_mrp_desktop = config.get("caminho_mrp_desktop", "")
                nome_sheets = config.get("name_sheets", "")
                if not caminho_mrp_desktop or not nome_sheets:
                    logging.info("O arquivo configuracoes_software.json não contém a chave 'caminho_mrp_desktop'.")
                    return

            Utils.renomear_planilhas_arquivos_xls(caminho_mrp_desktop, nome_sheets)
            Utils.converter_xls_para_xlsx(caminho_mrp_desktop)


        except Exception as e:
            logging.error(f"Ocorreu um erro ao abrir o MRP: {e}")
            raise

    @staticmethod
    def relatorio_Req_Almoxarifado(): # Método para gerar o relatório de Req. Almoxarifado
        try:
             
            with open('configuracoes_nomes.json', 'r') as f:
                config = json.load(f)
                
                nome_arquivo = config.get("req_almoxarifado", "")

            Utils.clicar_elemento('img/Rel_Req_Almoxarifado/Req_Almoxarifado.png', "Campo 'Req. Almoxarifado'")
            time.sleep(10)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_Req_Almoxarifado/Req_Almoxarifado_Lista.png')

            Utils.clicar_elemento('img/Rel_Req_Almoxarifado/Produto.png', "Botão 'Produto'")
            time.sleep(5)
            # Utils.expand_All_Notes('img/Rel_Req_Almoxarifado/Expand_Notes/Ano_Entrega.png', 'img/Rel_Req_Almoxarifado/Expand_Notes/Mes_Entrega.png')

            Utils.salvar_arquivo(nome_arquivo)
            time.sleep(10)  # Ajuste conforme necessário
        except Exception as e:
            logging.error(f"Ocorreu um erro ao gerar o relatório de Estoque Supper: {e}")
            raise

    @staticmethod
    def relatorio_Mov_Saida_ODF(): # Método para gerar o relatório de Movimento Saida de ODF
        try:

            with open('configuracoes_nomes.json', 'r') as f:
                config = json.load(f)
                
                nome_arquivo = config.get("mov_saida_odf", "")

            time.sleep(10)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_Mov_Saida_ODF/Mov_Saida_ODF.png', "Campo 'Mov. Saída ODF'")
            time.sleep(10)  # Ajuste conforme necessário

            Utils.aplicar_data_mes_atual('img/Rel_FaturamentoPeso/DataInicial.png', 'img/Rel_FaturamentoPeso/DataFinal.png')

            Utils.apicar_modelo_de_relatorio('img/Rel_Mov_Saida_ODF/Mov_Saida_ODF_Lista.png')

            Utils.clicar_elemento('img/Rel_Mov_Saida_ODF/Desc_Prod.png', "Botão 'Desc_Prod'")
            time.sleep(5)
            Utils.expand_All_Notes('img/Rel_Mov_Saida_ODF/Expand_Notes/Cod_Prod.png')

            Utils.salvar_arquivo_mov_saida_odf(nome_arquivo)
            time.sleep(10)

        except Exception as e:
            logging.error(f"Ocorreu um erro ao gerar o relatório de Movimento Saida de ODF: {e}")
            raise

    @staticmethod
    def relatorio_Parametros(): # Método para gerar o relatório de Parâmetros
        try:

            with open('configuracoes_nomes.json', 'r') as f:
                config = json.load(f)
                
                nome_arquivo = config.get("parametros", "")

            Utils.clicar_elemento('img/Menu_Configuracoes/Configuracoes.png', "Menu 'Configurações'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Menu_Configuracoes/Parametros.png', "Submenu 'Parâmetros'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Rel_Parametros/Filtro.png', "Icone 'Filtro'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Rel_Parametros/Excel.png', "Icone 'Excel'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Rel_Parametros/BotaoOK.png', "Icone 'Excel'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Rel_Parametros/SetaDupla.png', "Icone 'OK'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.salvar_arquivo_parametros(nome_arquivo, 'img/Rel_Parametros/Excel.png')

        except Exception as e:
            logging.error(f"Ocorreu um erro ao gerar o relatório de Parametros: {e}")
            raise

    @staticmethod
    def fechar_sistema(): # Método para fechar o sistema
        try:
            time.sleep(5)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Fechar_Sistema/X.png', "Botão 'X - Fechar'")
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            logging.error(f"Ocorreu um erro ao gerar o relatório de Parametros: {e}")
            raise

    @staticmethod            
    def extrair_relatorios(): # Método para extrair os relatórios
            
            #configura os logs antes de qualquer operação
            configurar_logs()

            try:
                
                logging.info("Iniciando o sistema...")
                RoboV4.abrir_software()  # Abre o software RoboV4
                logging.info("Realizando Login...")
                RoboV4.realizar_login()  # Realiza o login

                #Sequência de chamadas para extrair diferentes relatórios
                logging.info("Abrindo Cubo de Decisão...")
                RoboV4.abrir_cubo_de_decisao()
                logging.info("Retirando Relatorio Estoque SUPPER...")
                RoboV4.relatorio_estoque_supper()
                logging.info("Retirando Relatorio Faturamento Peso...")
                RoboV4.relatorio_faturamento_peso()
                logging.info("Retirando Relatorio Pedido Compra...")
                RoboV4.relatorio_pedido_compra()
                logging.info("Retirando Relatorio Movimento de Saida de ODF...")
                RoboV4.relatorio_Mov_Saida_ODF()

                if datetime.now().weekday() == 0:  # Se for segunda-feira
                    logging.info("Iniciando extração dos Parametros...")
                    RoboV4.relatorio_Parametros()  # Extrai parâmetros específicos
                else:
                    logging.info("Não é segunda-feira, não serão extraídos os parâmetros...")

                logging.info("Retirando Relatorio MRP...")
                RoboV4.abrir_mrp()
                logging.info("Relatórios extraídos com sucesso.")  # Loga a mensagem de sucesso

                logging.info("Fechando o Sistema...")
                RoboV4.fechar_sistema()  # Certifica-se de que o sistema será fechado
                
                message = "Relatórios extraídos com sucesso!"
                EnviarLogs.enviar_mensagem_de_sucesso(message)

            except Exception as e:
                logging.error(f"Ocorreu um erro no main principal(): {e}", exc_info=True)
                EnviarLogs.enviar_logs()
                raise