from dotenv import load_dotenv
import os

load_dotenv()

class Variaveis():
    def __init__(self):
    
        # Pastas Exportação
        self.pasta_analise_de_caixa = os.getenv("PASTA_ANALISE_DE_CAIXA")
        self.pasta_exportacao_extratos = os.getenv("PASTA_EXPORTACAO_EXTRATOS")
        self.pasta_input = os.getenv("PASTA_INPUT")

        self.user_bradesco = os.getenv("USER_BRADESCO")
        self.senha_bradesco = os.getenv("SENHA_BRADESCO")
        
        self.host_postgre = os.getenv("HOST_POSTGRE")
        self.user_postgre = os.getenv("USER_POSTGRE")
        self.senha_postgre = os.getenv("SENHA_POSTGRE")
        self.database_postgre = os.getenv("DATABASE_POSTGRE")