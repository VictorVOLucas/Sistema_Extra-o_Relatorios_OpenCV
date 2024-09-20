from telegram import Bot
import asyncio
import os
import json
from utils import Utils

class EnviarLogs:

    def __init__(self):
        try: 
            with open('configuracoes_telegram.json', 'r') as f: # Abre o arquivo de configurações
                    config = json.load(f)
                    
                    Token = config.get("Token", "")
                    ChatID = config.get("ChatID", "") 
                    Diretorio = config.get("Diretorio", "") 
        except Exception as e:
            Utils.show_message(f"Configurações do Telegram não encontrado: {e}")
            raise
        
        # Substitua 'YOUR_TOKEN' pelo token do seu bot do Telegram
        self.TOKEN = Token
        
        # Substitua 'CHAT_ID' pelo ID do chat onde você quer enviar o arquivo
        self.CHAT_ID = ChatID
        
        # Caminho para a pasta que contém os arquivos LOG
        self.FOLDER_PATH = Diretorio
        
        # Cria uma instância do bot
        self.bot = Bot(token=self.TOKEN)

    # PEGAR O ÚLTIMO ARQUIVO LOG GERADO NA PASTA
    def get_latest_file(self):
        log_files = [f for f in os.listdir(self.FOLDER_PATH) if f.endswith('.log')]
        if not log_files:
            return None
        full_paths = [os.path.join(self.FOLDER_PATH, f) for f in log_files]
        latest_file = max(full_paths, key=os.path.getmtime)  # Obtém o arquivo mais recente
        return latest_file

    # ENVIAR O ARQUIVO PARA O TELEGRAM (ASSÍNCRONO)
    async def send_file(self, file_path):
        try:
            with open(file_path, 'rb') as file:
                await self.bot.send_document(chat_id=self.CHAT_ID, document=file)  # Usar await
            print(f"Arquivo {os.path.basename(file_path)} enviado com sucesso!")
        except Exception as e:
            print(f"Erro ao enviar o arquivo: {e}")

    # ENVIAR UMA MENSAGEM PARA O TELEGRAM (ASSÍNCRONO)
    async def send_success_message(self, message):
        try:
            await self.bot.send_message(chat_id=self.CHAT_ID, text=message)  # Usar await
            print(f"Mensagem enviada com sucesso: {message}")
        except Exception as e:
            print(f"Erro ao enviar a mensagem: {e}")

    # FUNÇÃO PRINCIPAL ASSÍNCRONA PARA ENVIAR O ARQUIVO
    async def main_send_file(self):
        latest_file = self.get_latest_file()
        if latest_file:
            await self.send_file(latest_file)  # Usar await
        else:
            print("Nenhum arquivo LOG encontrado na pasta.")

    # Função principal para enviar uma mensagem
    async def main_send_message(self, message):
        await self.send_success_message(message)

    # Função para executar o envio de logs de forma assíncrona
    @staticmethod
    def enviar_logs():
        enviar_logs_instance = EnviarLogs()
        asyncio.run(enviar_logs_instance.main_send_file())  # Executar a função de envio de arquivo de forma assíncrona

    # Função para enviar mensagem de sucesso de forma assíncrona
    @staticmethod
    def enviar_mensagem_de_sucesso(message):
        enviar_logs_instance = EnviarLogs()
        asyncio.run(enviar_logs_instance.main_send_message(message))  # Executar a função de envio de mensagem de forma assíncrona