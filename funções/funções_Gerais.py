import requests
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def log_message(log_input, message):
    log_input.text += message + '\n'

class Log:
    def __init__(self):
        self.text = ''

log_input = Log() 

class SolicitarPlanilha:
    def __init__(self, log_input):
        self.log_input = log_input

    def escolher_planilha(self):
        # Iniciar uma janela oculta do tkinter
        root = tk.Tk()
        root.withdraw()  # Ocultar a janela principal

        # Abrir o diálogo de seleção de arquivo e permitir que o usuário selecione uma planilha
        file_path = filedialog.askopenfilename(
            title = "Selecione a planilha",
            filetypes = [("Planilhas Excel", "*.xlsx *.xls")]
        )

        # Ler as folhas do arquivo Excel
        try:
            excel_file = pd.ExcelFile(file_path)
            sheets = excel_file.sheet_names

            # Verificar se a folha 'XML' existe
            if 'Processos devolução (programa)' in sheets:
                planilha_df = pd.read_excel(file_path, sheet_name = 'Processos devolução (programa)')

            else:
                planilha_df = pd.read_excel(file_path, sheet_name = sheets[0], header = None)

            log_message(self.log_input, f'Planilha selecionada: {file_path}\n')

            return planilha_df

        except Exception as e:
            log_message(self.log_input, f'Erro ao ler o arquivo: {str(e)}\n')
            return None

solicitar = SolicitarPlanilha(log_input)

planilha_df = solicitar.escolher_planilha()

@staticmethod
def obter_configs():
    """Carrega as configurações do arquivo e retorna como dicionário"""
    try:
        with open('email.txt', 'r') as f:
            dados = f.read()
            dados = [item.strip() for item in dados.split(',')]
            return {
                "email": dados[0],
                "usuario_Elogistic": dados[1],
                "senha_Elogistic": dados[2],
                'usuario_IOB': dados[3],
                'senha_IOB': dados[4],
                'aliquota_interna': dados[5],
                'nome_credenciada': dados[6],
                'cnpj_credenciada': dados[7],
                'caixa_emails': dados[8],
            }
    except FileNotFoundError:
        return {}

if obter_configs:
    conta_email = obter_configs().get('email')
    caixa_emails = obter_configs().get('caixa_emails')
    usuario_elogistica = obter_configs().get('usuario_Elogistic')
    senha_elogistica = obter_configs().get('senha_Elogistic')
    usuario_IOB = obter_configs().get('usuario_IOB')
    senha_IOB = obter_configs().get('senha_IOB')
    aliquota_interna = obter_configs().get('aliquota_interna')
    nome_credenciada = obter_configs().get('nome_credenciada')
    cnpj_credenciada = obter_configs().get('cnpj_credenciada')
 
def obter_usuarios_validos():
    """Obtém a lista de usuários válidos do servidor."""
    try:
        # Fazendo a requisição para obter a lista de usuários válidos
        response = requests.get("https://servidor-para-emme2.onrender.com/usuarios")
        
        # Verificando se a resposta foi bem-sucedida
        if response.status_code == 200:
            data = response.json()  # Converte a resposta para JSON

            if isinstance(data, dict) and "usuarios" in data:
                usuarios = data["usuarios"]
                return usuarios  # Retorna a lista de usuários
            
            else:
                return []  # Se o formato for inesperado, retorna lista vazia
        else:
            return []  # Se houver erro de status, retorna lista vazia
        
    except Exception as e:
        return []  # Se ocorrer erro na requisição, retorna lista vazia

lista_autenticar = obter_usuarios_validos()

def validar_usuario(self, instance):
    """Valida o e-mail do usuário."""    
    email = self.email_autenticar.text.strip().lower()  # Remove espaços e transforma para minúsculas
    if not email:
        self.label_email_autenticar.text = "Por favor, insira um e-mail."
        return

    # Verifique se a lista de usuários não é None ou vazia
    if not lista_autenticar:
        self.label_email_autenticar.text = "Erro ao obter a lista de usuários válidos."
        return

    # Certifique-se de que `usuarios_validos` é uma lista e faça a verificação
    if isinstance(lista_autenticar, list) and email in [u.lower() for u in lista_autenticar]:
        self.label_email_autenticar.text = f" Acesso permitido! Bem-vindo(a)!"
        self.botao_entrar.disabled = False  # Desabilita o botão após sucesso na validação
        self.botao_entrar.color = 0, 1 ,0 ,1
    else:
        self.label_email_autenticar.text = "Acesso negado! Usuário não autorizado."
        self.botao_entrar.disabled = True  # Certifica-se de que o botão está habilitado se necessário
        self.botao_entrar.color = 1, 0 ,0 ,1