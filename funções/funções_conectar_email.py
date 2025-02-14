from pathlib import Path
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from kivy.uix.button import Button
import shutil
import imaplib
import pickle
import email
import os
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from funções.funções_Gerais import log_message, caixa_emails, conta_email, planilha_df

if os.name == 'nt':  # Para Windows
    download_dir = Path(os.getenv('USERPROFILE')) / 'Downloads'
else:  # Para macOS e Linux
    download_dir = Path.home() / 'Downloads'

current_dir = Path(__file__).parent

aux_path_json = current_dir / 'credentials.json'

aux_path_XML_destino = download_dir / 'XML_DELL'

aux_path_XML_destino.mkdir(parents = True, exist_ok = True)

planilha_df.ffill(inplace = True)

# Baixar arquivos HP e valores devolução HP
def conectar_email_e_baixar_arquivos_HP(self, *args):
    substrings_hp = []
    for index, row in planilha_df.iterrows():
        cell_hp = row[1]

        if isinstance(cell_hp, str) and len(cell_hp) >= 34:
            # Extrai a substring e adiciona à lista
            substring_hp = cell_hp[28:34]
            substrings_hp.append(substring_hp)
        elif isinstance(cell_hp, float):
            cell_int = int(cell_hp)
            substrings_hp.append(cell_int)
        else:
            substrings_hp.append(cell_hp)

    SCOPES = ['https://mail.google.com/']

    def get_gmail_service():
        """Realiza a autenticação OAuth2 e retorna o serviço do Gmail."""
        
        creds = None
        # Carrega as credenciais do arquivo 'token.pickle' se ele existir
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        # Verifica se as credenciais são válidas, e se não, tenta atualizar ou obter novas
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())  # Tenta atualizar o token
                except Exception as e:
                    log_message(self.log_input, f'Erro ao atualizar as credenciais: {e}')
                    # Se falhar, precisa obter novas credenciais
                    flow = InstalledAppFlow.from_client_secrets_file(str(aux_path_json), SCOPES)
                    creds = flow.run_local_server(port=0)  # Inicia o fluxo de autenticação
            else:
                # Se não houver credenciais, inicia o fluxo de autenticação
                flow = InstalledAppFlow.from_client_secrets_file(str(aux_path_json), SCOPES)
                creds = flow.run_local_server(port = 0)

            # Salva as credenciais para futuras execuções
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        return creds

    def generate_oauth2_string(username, access_token):
        """Gera a string de autenticação OAuth2 no formato necessário para IMAP."""
        auth_string = f'user={username}\1auth=Bearer {access_token}\1\1'
        return auth_string.encode('ascii')

    def connect_to_gmail_imap(creds, conta_email):
        """Conecta ao Gmail via IMAP usando OAuth2."""
        imap_host = 'imap.gmail.com'
        mail = imaplib.IMAP4_SSL(imap_host)
        
        # Gera a string de autenticação OAuth2
        auth_string = generate_oauth2_string(conta_email, creds.token)
        
        # Autentica no servidor IMAP usando OAuth2
        try:
            mail.authenticate('XOAUTH2', lambda x: auth_string)
            log_message(self.log_input, '\nAutenticação bem sucedida, seguindo para baixar emails\n')
        except Exception as e:
            log_message(self.log_input, f'Ocorreu um erro durante a autenticação com servidor: {e}')
            return None
        
        return mail

    # Inicie a autenticação
    creds = get_gmail_service()

    # Conecte ao Gmail via IMAP usando as credenciais OAuth2
    mail = connect_to_gmail_imap(creds, conta_email)

    if mail:
        # Selecionar a caixa de entrada ou outro rótulo
        mail.select(f'"{caixa_emails}"')

        try:
            processed_files = []

            idx = 1

            for xml in substrings_hp:
                status, email_ids = mail.search(None, f'(SUBJECT "{xml}")')
                log_message(self.log_input, f'{idx} - Buscando emails com assunto: "{xml}"')

                for email_id in email_ids[0].split():
                    status, email_data = mail.fetch(email_id, '(RFC822)')
                    raw_email = email_data[0][1]
                    message = email.message_from_bytes(raw_email)

                    # Processar o e-mail e anexos
                    for part in message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        file_name = part.get_filename()

                        # Verificar se o anexo já foi processado
                        if file_name and file_name.endswith('-nfe.xml') and file_name not in processed_files:
                            log_message(self.log_input, f'{idx} - Anexo baixado: {file_name}\n')

                            try:
                                save_path = os.path.join(aux_path_XML_destino, file_name)
                                os.makedirs(aux_path_XML_destino, exist_ok = True)

                                # Salvar o arquivo
                                with open(save_path, 'wb') as f:
                                    f.write(part.get_payload(decode = True))

                                if os.path.exists(save_path):
                                    destino_path = os.path.join(aux_path_XML_destino, file_name)
                                    shutil.move(save_path, destino_path)

                                    # Adicionar o arquivo processado na lista
                                    processed_files.append(file_name)
                                else:
                                    log_message(self.log_input, f'{idx} - O arquivo de origem não foi encontrado: {save_path}')

                                    idx += 1

                            except Exception as e:
                                log_message(self.log_input, f'{idx} - Ocorreu um erro ao processar o anexo "{file_name}": {e}')

                                idx += 1
                idx += 1

        except Exception as e:
            log_message(self.log_input, f'Não foi possível prosseguir com a busca no email: {e}')
    else:
        log_message(self.log_input, f'A conexão IMAP falhou.')

    substrings_hp.clear()

# Baixar arquivos Dell e Valores devolução Dell
def conectar_email_e_baixar_arquivos_Dell(self, *args):
    substrings_Dell = []

    for index, row in planilha_df.iterrows():

        cell_Dell = row[1]

        if isinstance(cell_Dell, str) and len(cell_Dell) >= 34:
            # Extrai a substring e adiciona à lista
            substring_Dell = cell_Dell[27:34]
            substrings_Dell.append(substring_Dell)
        elif isinstance(cell_Dell, float):
            cell_int = int(cell_Dell)
            substrings_Dell.append(cell_int)
        else:
            substrings_Dell.append(cell_Dell)

    SCOPES_devolução = ['https://mail.google.com/']

    def get_gmail_service():
        """Realiza a autenticação OAuth2 e retorna o serviço do Gmail."""
        
        creds = None
        # Carrega as credenciais do arquivo 'token.pickle' se ele existir
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        # Verifica se as credenciais são válidas, e se não, tenta atualizar ou obter novas
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())  # Tenta atualizar o token
                except Exception as e:
                    log_message(self.log_input, f'Erro ao atualizar as credenciais: {e}')
                    # Se falhar, precisa obter novas credenciais
                    flow = InstalledAppFlow.from_client_secrets_file(str(aux_path_json), SCOPES_devolução)
                    creds = flow.run_local_server(port=0)  # Inicia o fluxo de autenticação
            else:
                # Se não houver credenciais, inicia o fluxo de autenticação
                flow = InstalledAppFlow.from_client_secrets_file(str(aux_path_json), SCOPES_devolução)
                creds = flow.run_local_server(port=0)

            # Salva as credenciais para futuras execuções
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        return creds

    def generate_oauth2_string(username, access_token):
        """Gera a string de autenticação OAuth2 no formato necessário para IMAP."""
        auth_string = f'user={username}\1auth=Bearer {access_token}\1\1'
        return auth_string.encode('ascii')

    def connect_to_gmail_imap(creds, conta_email):
        """Conecta ao Gmail via IMAP usando OAuth2."""
        imap_host = 'imap.gmail.com'
        mail = imaplib.IMAP4_SSL(imap_host)
        
        # Gera a string de autenticação OAuth2
        auth_string = generate_oauth2_string(conta_email, creds.token)
        
        # Autentica no servidor IMAP usando OAuth2
        try:
            mail.authenticate('XOAUTH2', lambda x: auth_string)
            log_message(self.log_input, '\nAutenticação bem sucedida, seguindo para baixar emails\n')
        except Exception as e:
            log_message(self.log_input, f'\nOcorreu um erro durante a autenticação com servidor: {e}')
            return None
        
        return mail

    # Inicie a autenticação
    creds_devolução = get_gmail_service()

    # Conecte ao Gmail via IMAP usando as credenciais OAuth2
    mail = connect_to_gmail_imap(creds_devolução, conta_email)

    if mail:
        # Selecionar a caixa de entrada ou outro rótulo
        mail.select(f'"{caixa_emails}"')

        try:
            processed_files_devolução = []

            idx = 1

            for xml in substrings_Dell:
                status, email_ids = mail.search(None, f'(SUBJECT "{xml}")')
                log_message(self.log_input, f'{idx} - Buscando emails com assunto: "{xml}"\n')

                for email_id in email_ids[0].split():
                    status, email_data = mail.fetch(email_id, '(RFC822)')
                    raw_email = email_data[0][1]
                    message = email.message_from_bytes(raw_email)

                    # Processar o e-mail e anexos
                    for part in message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        file_name = part.get_filename()

                        # Verificar se o anexo já foi processado
                        if file_name and file_name.endswith('-procNFe.xml') and file_name not in processed_files_devolução:
                            log_message(self.log_input, f'{idx} - Anexo baixado: {file_name}\n')

                            try:
                                save_path = os.path.join(aux_path_XML_destino, file_name)
                                os.makedirs(aux_path_XML_destino, exist_ok = True)

                                # Salvar o arquivo
                                with open(save_path, 'wb') as f:
                                    f.write(part.get_payload(decode = True))

                                if os.path.exists(save_path):
                                    destino_path = os.path.join(aux_path_XML_destino, file_name)
                                    shutil.move(save_path, destino_path)

                                    # Adicionar o arquivo processado na lista
                                    processed_files_devolução.append(file_name)
                                else:
                                    log_message(self.log_input, f'{idx} - O arquivo de origem não foi encontrado: {save_path}')

                                    idx += 1

                            except Exception as e:
                                log_message(self.log_input, f'{idx} - Ocorreu um erro ao processar o anexo "{file_name}": {e}')

                                idx += 1
                idx += 1

        except Exception as e:
            log_message(self.log_input, f'Não foi possível prosseguir com a busca no email: {e}')
    else:
        log_message(self.log_input, f'A conexão IMAP falhou.')

    substrings_Dell.clear()

# Botoes_entrada_Dell
def baixar_arquivosXML_DELL(self, *args):
    
    def mostrar_confirmacao(self):
        box = BoxLayout(orientation = 'vertical', padding = 10, spacing = 10)
        
        # Texto do aviso
        label = Label(text = f'Caso existam arquivos na pasta:\n\n {aux_path_XML_destino}, serão excluídos.\n\nDeseja continuar?', halign = 'center', valign = 'center', text_size = (400, None))
        
        # Botões "Sim" e "Não"
        botoes = BoxLayout(size_hint_y = 0.3, spacing = 10)
        botao_sim = Button(text = "Sim")
        botao_nao = Button(text = "Não")
        
        botoes.add_widget(botao_sim)
        botoes.add_widget(botao_nao)
        
        # Adiciona o texto e os botões ao layout principal do popup
        box.add_widget(label)
        box.add_widget(botoes)
        
        # Configura o popup
        popup = Popup(title = 'Aviso', content = box, size_hint = (0.5, 0.4))

        def continuar(*args):

            # Lista todos os arquivos na pasta
            arquivos_excluir = os.listdir(aux_path_XML_destino)

            for arquivo_excluir in arquivos_excluir:
                caminho_arquivo_excluir = os.path.join(aux_path_XML_destino, arquivo_excluir)
                if os.path.isfile(caminho_arquivo_excluir):
                    os.remove(caminho_arquivo_excluir)

            conectar_email_e_baixar_arquivos_Dell(self)

            popup.dismiss()

        def parar(instance):
            popup.dismiss()
            return

        botao_sim.bind(on_release = continuar)
        botao_nao.bind(on_release = parar)
    
        # Exibe o popup
        popup.open()

    mostrar_confirmacao(self)

# Botoes_entrada_HP
def baixar_arquivosXML_HP(self, *args):
        
    def mostrar_confirmacao(self):
        box = BoxLayout(orientation = 'vertical', padding = 10, spacing = 10)
        
        # Texto do aviso
        label = Label(text = f'Caso existam arquivos na pasta:\n\n {aux_path_XML_destino}, serão excluídos.\n\nDeseja continuar?', halign = 'center', valign = 'center', text_size = (400, None))
        
        # Botões "Sim" e "Não"
        botoes = BoxLayout(size_hint_y = 0.3, spacing = 10)
        botao_sim = Button(text = "Sim")
        botao_nao = Button(text = "Não")
        
        botoes.add_widget(botao_sim)
        botoes.add_widget(botao_nao)
        
        # Adiciona o texto e os botões ao layout principal do popup
        box.add_widget(label)
        box.add_widget(botoes)
        
        # Configura o popup
        popup = Popup(title = 'Aviso', content = box, size_hint = (0.5, 0.4))

        def continuar(*args):

            # Lista todos os arquivos na pasta
            arquivos_excluir = os.listdir(aux_path_XML_destino)

            for arquivo_excluir in arquivos_excluir:
                caminho_arquivo_excluir = os.path.join(aux_path_XML_destino, arquivo_excluir)
                if os.path.isfile(caminho_arquivo_excluir):
                    os.remove(caminho_arquivo_excluir)

            conectar_email_e_baixar_arquivos_HP(self)

            popup.dismiss()

        def parar(instance):
            popup.dismiss()
            return

        botao_sim.bind(on_release = continuar)
        botao_nao.bind(on_release = parar)
    
        # Exibe o popup
        popup.open()

    mostrar_confirmacao(self)