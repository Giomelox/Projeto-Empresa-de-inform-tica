import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from kivy.uix.button import Button
from xml.dom import minidom
import shutil
import os
import pandas as pd
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from funções.funções_Gerais import log_message
from funções.funções_conectar_email import aux_path_XML_destino, download_dir, conectar_email_e_baixar_arquivos_HP, conectar_email_e_baixar_arquivos_Dell
from funções.funções_Gerais import usuario_elogistica, senha_elogistica, planilha_df

planilha_df.ffill(inplace = True)

def comparar_cprod(xml_file, cprod_excel, nNF_excel):
    try:
        tree = minidom.parse(xml_file)
        nf_list = tree.getElementsByTagName('nNF')  # Lista com os números de nNF
        det_list = tree.getElementsByTagName('det')  # Lista com os números de det

        # Verifica se há nNF correspondente
        for nf in nf_list:
            if nf.firstChild and nf.firstChild.data.strip() == nNF_excel.strip():  # Compara nNF
                # Para o nNF correspondente, procura os elementos <det>
                for det in det_list:
                    prod_element = det.getElementsByTagName('prod')[0] if det.getElementsByTagName('prod') else None  # Verifica se <prod> existe
                    if prod_element:
                        cProd = prod_element.getElementsByTagName('cProd')[0].firstChild.data.strip() if prod_element.getElementsByTagName('cProd') else None
                        if cProd and cProd == cprod_excel.strip():  # Compara cProd do XML com o valor do Excel
                            return det  # Retorna o <det> correspondente
        return None  # Retorna None se não houver correspondência
    except Exception as e:
        print(f"Erro ao processar o arquivo XML {xml_file}: {e}")
        return None

# Botoes_entrada_Dell e HP
def biparxml(self, *args):

    driver = webdriver.Chrome()
    
    substrings_bipar = []
    
    for index, row in planilha_df.iterrows():

        cell_bipar = row[1]

        if isinstance(cell_bipar, str) and len(cell_bipar) >= 34:
            # Extrai a substring e adiciona à lista
            substring_bipar = cell_bipar
            substrings_bipar.append(substring_bipar)
        else:
            substrings_bipar.append(cell_bipar)

    idx = 1

    # Acessa a página desejada
    driver.get('https://innovation.uninet.com.br/Ulog/Main.asp?All=reset')

    time.sleep(5)

    try:
        element_usuario = WebDriverWait(driver, 500).until(
            EC.presence_of_element_located((By.NAME, 'login'))
        )
        element_usuario.send_keys(usuario_elogistica)

        element_senha = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.NAME, 'password'))
        )
        element_senha.send_keys(senha_elogistica)

        element_submit = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.NAME, 'submit'))
        )
        element_submit.click()

        time.sleep(5)

        element_botao_notificar_recebimento = WebDriverWait(driver, 500).until(
            EC.presence_of_element_located((By.XPATH, "(//input[@id='submit'])[2]"))
        )
        element_botao_notificar_recebimento.click()
    except:
        log_message(self.log_input, 'Elemento não apareceu dentro do tempo esperado.')

    time.sleep(10)

    try:
        for x in substrings_bipar:

            element_xml = WebDriverWait(driver, 500).until(
                EC.presence_of_element_located((By.NAME, 'keynfe'))
            )
            element_xml.send_keys(x)

            time.sleep(2)

            element_submit_xml = driver.find_element(By.NAME, 'submit')
            element_submit_xml.click()
            
            log_message(self.log_input, f'{idx} - Arquivo {x} processado')

            idx += 1

            time.sleep(3)
            
    except Exception as e:
        log_message(self.log_input, f'Não foi possível adicionar o arquivo {x}: {e}\n')

    time.sleep(5)

    substrings_bipar.clear()

    # Fechar o navegador
    driver.quit()

# Botoes_devolução_Dell
def valores_devolução_DELL(self, *args):

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
        
        # Função para o botão "Sim"
        def continuar(instance):

            driver = webdriver.Chrome()

            # Lista todos os arquivos na pasta
            arquivos_excluir = os.listdir(aux_path_XML_destino)

            for arquivo_excluir in arquivos_excluir:
                caminho_arquivo_excluir = os.path.join(aux_path_XML_destino, arquivo_excluir)
                if os.path.isfile(caminho_arquivo_excluir):
                    os.remove(caminho_arquivo_excluir)

            conectar_email_e_baixar_arquivos_Dell(self)
            
            popup.dismiss()

            idx_planilha = 1

            # Acessa a página desejada
            driver.get('https://innovation.uninet.com.br/Ulog/Main.asp?All=reset')

            try:
                element_usuario = WebDriverWait(driver, 200).until(
                    EC.presence_of_element_located((By.NAME, 'login'))
                )
                element_usuario.send_keys(usuario_elogistica)

                element_senha = driver.find_element(By.NAME, 'password')
                element_senha.send_keys(senha_elogistica)

                element_submit = driver.find_element(By.NAME, 'submit')
                element_submit.click()

                time.sleep(5)

                element_peças_pendentes_devolução = WebDriverWait(driver, 200).until(
                    EC.presence_of_element_located((By.XPATH, '(//input[@id="submit"])[4]'))
                )
                element_peças_pendentes_devolução.click()
            except:
                log_message(self.log_input, 'O site demorou mais que o esperado para carregar.')
                
            resultados = {}
            
            for index, row in planilha_df.iloc[:].iterrows():
                CHAMADO = str(row[0]).replace('.0', '').strip()
                NF = str(row[1]).replace('.0', '').strip()
                PN = str(row[2]).strip()
                PPID = str(row[4]).strip()
                STATUS_planilha = str(row[9]).strip()

                if STATUS_planilha.upper() == 'DEFECTIVE' or STATUS_planilha.upper() == 'DOA':

                    if PPID.upper() == 'X':
                        log_message(self.log_input, f'{idx_planilha}: Peça sem PPID na planilha.\n')

                        try:
                                    # Encontra o valor do produto
                                    element_Vprod = elements.find_element(By.XPATH, './/td[9]')
                                    element_Vprod_value = element_Vprod.text

                                    # Adiciona o valor ao dicionário
                                    resultados[f"{PN}_{STATUS_planilha}"] = {'VALOR': element_Vprod_value, 'CST': None, 'NCM': None, 'QTD.': None}

                                    if not os.path.exists(aux_path_XML_destino):
                                        continue
                                    else:  
                                        # Lista todos os arquivos na pasta
                                        arquivos = os.listdir(aux_path_XML_destino)

                                        for arquivo in arquivos:
                                            caminho_arquivo = os.path.join(aux_path_XML_destino, arquivo)

                                            if os.path.isfile(caminho_arquivo):
                                                try:
                                                    # Abrir o arquivo XML no modo leitura
                                                    with open(caminho_arquivo, 'r'):

                                                        # Comparar cProd do arquivo XML com o valor PN do Excel
                                                        det_element = comparar_cprod(caminho_arquivo, PN, NF)

                                                        if det_element:
                                                            # Buscar o valor de <orig> e <NCM> dentro do mesmo <det>
                                                            CST = det_element.getElementsByTagName('orig')
                                                            NCM_prod = det_element.getElementsByTagName('NCM')
                                                            Quantidade_prod = det_element.getElementsByTagName('qCom')
                                                            
                                                            if CST and NCM_prod:
                                                                resultados[f'{PN}_{STATUS_planilha}']['CST'] = str(CST[0].firstChild.data)
                                                                resultados[f'{PN}_{STATUS_planilha}']['NCM'] = str(NCM_prod[0].firstChild.data)
                                                                resultados[f'{PN}_{STATUS_planilha}']['QTD.'] = float(Quantidade_prod[0].firstChild.data)

                                                            else:
                                                                log_message(self.log_input, f'Elementos CST, NCM e/ou Quantidade não encontrados no arquivo {arquivo}.')
                                                except Exception as e:
                                                    log_message(self.log_input, f'Erro ao processar o arquivo {arquivo}: {e}\n')
                                    idx_planilha += 1

                        except Exception as e:
                            log_message(self.log_input, f'{idx_planilha}: Chamado {CHAMADO}/{PN} não consta no site.\n')
                            idx_planilha += 1
                            continue

                    else:
                        try:
                            
                            # Tenta encontrar o elemento correspondente
                            elements = WebDriverWait(driver, 20).until(
                                EC.presence_of_element_located((By.XPATH, f'//tr[contains(td[4], "{CHAMADO}") and contains(td[6], "{PN}")]'))
                            )

                            time.sleep(1)

                            if elements:
                                try:
                                    # Encontra o valor do produto
                                    element_Vprod = elements.find_element(By.XPATH, './/td[9]')
                                    element_Vprod_value = element_Vprod.text

                                    # Adiciona o valor ao dicionário
                                    resultados[f"{PN}_{STATUS_planilha}"] = {'VALOR': element_Vprod_value, 'CST': None, 'NCM': None, 'QTD.': None}

                                    time.sleep(1)

                                    # Encontra o link <a> na mesma linha
                                    element_link = elements.find_element(By.XPATH, './/td[5]')
                                    element_link.click()

                                    time.sleep(1)
                                    
                                    # Insere o PPID da planilha no site
                                    inserir_ppid = WebDriverWait(driver, 60).until(
                                    EC.presence_of_element_located((By.NAME, 'PPID'))
                                    )
                                    inserir_ppid.send_keys(PPID)

                                    botao_enviar_ppid = WebDriverWait(driver, 20).until(
                                    EC.presence_of_element_located((By.NAME, 'submit'))
                                    )
                                    botao_enviar_ppid.click()

                                    time.sleep(1)

                                    alerta = driver.switch_to.alert
                                    alerta.accept()

                                    if not os.path.exists(aux_path_XML_destino):
                                        continue
                                    else:  
                                        # Lista todos os arquivos na pasta
                                        arquivos = os.listdir(aux_path_XML_destino)

                                        for arquivo in arquivos:
                                            caminho_arquivo = os.path.join(aux_path_XML_destino, arquivo)

                                            if os.path.isfile(caminho_arquivo):
                                                try:
                                                    # Abrir o arquivo XML no modo leitura
                                                    with open(caminho_arquivo, 'r'):

                                                        # Comparar cProd do arquivo XML com o valor PN do Excel
                                                        det_element = comparar_cprod(caminho_arquivo, PN, NF)

                                                        if det_element:
                                                            # Buscar o valor de <orig> e <NCM> dentro do mesmo <det>
                                                            CST = det_element.getElementsByTagName('orig')
                                                            NCM_prod = det_element.getElementsByTagName('NCM')
                                                            Quantidade_prod = det_element.getElementsByTagName('qCom')
                                                            
                                                            if CST and NCM_prod:
                                                                resultados[f'{PN}_{STATUS_planilha}']['CST'] = str(CST[0].firstChild.data)
                                                                resultados[f'{PN}_{STATUS_planilha}']['NCM'] = str(NCM_prod[0].firstChild.data)
                                                                resultados[f'{PN}_{STATUS_planilha}']['QTD.'] = float(Quantidade_prod[0].firstChild.data)

                                                            else:
                                                                log_message(self.log_input, f'Elementos CST ou NCM não encontrados no arquivo {arquivo}.')
                                                except Exception as e:
                                                    log_message(self.log_input, f'Erro ao processar o arquivo {arquivo}: {e}\n')

                                            element_peças_pendentes_devolução1 = WebDriverWait(driver, 200).until(
                                            EC.presence_of_element_located((By.XPATH, '(//input[@id="submit"])[4]'))
                                            )

                                    element_peças_pendentes_devolução1.click()
            
                                except Exception as e:
                                    print()

                                log_message(self.log_input, f'{idx_planilha}: Chamado {CHAMADO}/{PN} inserido com sucesso.\n')
                                idx_planilha += 1

                        except Exception as e:

                            log_message(self.log_input, f'{idx_planilha}: Chamado {CHAMADO}/{PN} não consta no site.\n')
                            idx_planilha += 1


                elif STATUS_planilha.upper() == 'GOOD' or STATUS_planilha.lower() == 'nan':
                    try:
                        if os.path.exists(aux_path_XML_destino):
                            arquivos = os.listdir(aux_path_XML_destino)

                            for arquivo in arquivos:
                                caminho_arquivo = os.path.join(aux_path_XML_destino, arquivo)

                                if os.path.isfile(caminho_arquivo):
                                    det_element = comparar_cprod(caminho_arquivo, PN, NF)

                                    if det_element:
                                        vProd = det_element.getElementsByTagName('vProd')
                                        CST = det_element.getElementsByTagName('orig')
                                        NCM_prod = det_element.getElementsByTagName('NCM')
                                        Quantidade_prod = det_element.getElementsByTagName('qCom')
                                        
                                        if CST and NCM_prod and vProd:
                                            # Adiciona as informações ao dicionário
                                            resultados[f"{PN}_{STATUS_planilha}"] = {
                                            'VALOR': str(vProd[0].firstChild.data) if vProd else None,
                                            'CST': str(CST[0].firstChild.data) if CST else None,
                                            'NCM': str(NCM_prod[0].firstChild.data) if NCM_prod else None,
                                            'QTD.': float(Quantidade_prod[0].firstChild.data) if Quantidade_prod else None
                                        }
                                        else:
                                            log_message(self.log_input, f'Elementos CST ou NCM não encontrados no arquivo {arquivo}.')
                                            
                        log_message(self.log_input, f'{idx_planilha}: Chamado {CHAMADO}/{PN} inserido com sucesso.\n')
                        idx_planilha += 1
                        
                    except Exception as e:
                        log_message(self.log_input, f'Não foi possível abrir o arquivo {idx_planilha}: {e}')

                        idx_planilha += 1

            # Criando uma cópia das colunas selecionadas do DataFrame original
            planilha_copia = planilha_df.iloc[:, [0, 1, 2]].copy()

            # Renomeando as colunas da cópia
            planilha_copia.columns = ['CHAMADO', 'REF NF', 'PART NUMBER']

            # Adicionando as colunas ao DataFrame a partir do dicionário
            planilha_copia['VALOR'] = [
                resultados.get(f"{row[2]}_{row[9]}", {}).get('VALOR', None)
                for _, row in planilha_df.iterrows()
            ]

            planilha_copia['CST'] = [
                resultados.get(f"{row[2]}_{row[9]}", {}).get('CST', None)
                for _, row in planilha_df.iterrows()
            ]

            planilha_copia['NCM'] = [
                resultados.get(f"{row[2]}_{row[9]}", {}).get('NCM', None)
                for _, row in planilha_df.iterrows()
            ]

            planilha_copia['QTD.'] = [
                resultados.get(f'{row[2]}_{row[9]}', {}).get('QTD.', None)
                for _, row in planilha_df.iterrows()
            ]

            # Substituindo pontos por vírgulas na coluna VALOR
            planilha_copia['VALOR'] = planilha_copia['VALOR'].astype(str).str.replace('.', ',', regex = False)

            # Exportando o DataFrame para Excel
            planilha_copia.to_excel('Planilha Dell valores_devolução formatada.xlsx', index = False)

            if os.path.exists(f'{download_dir}\\Planilha Dell valores_devolução formatada.xlsx'):
                os.remove(f'{download_dir}\\Planilha Dell valores_devolução formatada.xlsx')

            shutil.move('Planilha Dell valores_devolução formatada.xlsx', download_dir)

            log_message(self.log_input, f'Planilha criada com sucesso: Planilha Dell valores_devolução formatada.xlsx')
            
            time.sleep(5)
            driver.quit()

        def parar(instance):
            popup.dismiss()
            return

        botao_sim.bind(on_release = continuar)
        botao_nao.bind(on_release = parar)
    
        # Exibe o popup
        popup.open()

    mostrar_confirmacao(self)

# Botoes_devolução_HP
def valores_devolução_HP(self, *args):

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
        
        # Função para o botão "Sim"
        def continuar(instance):

            # Lista todos os arquivos na pasta
            arquivos_excluir = os.listdir(aux_path_XML_destino)

            for arquivo_excluir in arquivos_excluir:
                caminho_arquivo_excluir = os.path.join(aux_path_XML_destino, arquivo_excluir)
                if os.path.isfile(caminho_arquivo_excluir):
                    os.remove(caminho_arquivo_excluir)

            conectar_email_e_baixar_arquivos_HP(self)
            
            popup.dismiss()

            idx_planilha = 1

            resultados_hp = {}

            for index, row in planilha_df.iloc[:].iterrows():
                CHAMADO = str(row[0]).strip()
                NF = str(row[1]).strip()
                PN = str(row[2]).strip()
                STATUS_planilha = str(row[9]).strip()

                try:
                    if os.path.exists(aux_path_XML_destino):

                        arquivos = os.listdir(aux_path_XML_destino)

                        for arquivo in arquivos:
                            caminho_arquivo = os.path.join(aux_path_XML_destino, arquivo)

                            if os.path.isfile(caminho_arquivo):
                                det_element = comparar_cprod(caminho_arquivo, PN, NF)

                                if det_element:
                                        
                                    vProd = det_element.getElementsByTagName('vProd')

                                    CST = det_element.getElementsByTagName('orig')

                                    NCM_prod = det_element.getElementsByTagName('NCM')

                                    Quantidade_prod = det_element.getElementsByTagName('qCom')
                                    
                                    if CST and NCM_prod and vProd:

                                        if STATUS_planilha.upper() == 'GOOD':

                                            resultados_hp[f"{PN}_{STATUS_planilha}"] = {
                                            'VALOR': float(f'{float(vProd[0].firstChild.data):.2f}') if vProd else None,
                                            'CST': str(CST[0].firstChild.data) if CST else None,
                                            'NCM': str(NCM_prod[0].firstChild.data) if NCM_prod else None,
                                            'QTD.': float(Quantidade_prod[0].firstChild.data) if Quantidade_prod else None
                                        }
                                            
                                        elif STATUS_planilha in ['DEFECTIVE', 'DOA']:
                                    
                                            resultados_hp[f"{PN}_{STATUS_planilha}"] = {
                                            'VALOR': float(f'{float(vProd[0].firstChild.data) / 10:.2f}') if vProd else None,
                                            'CST': str(CST[0].firstChild.data) if CST else None,
                                            'NCM': str(NCM_prod[0].firstChild.data) if NCM_prod else None,
                                            'QTD.': float(Quantidade_prod[0].firstChild.data) if Quantidade_prod else None
                                        }
                                    else:
                                        log_message(self.log_input, f'Elementos CST ou NCM não encontrados no arquivo {arquivo}.')
                                        
                    log_message(self.log_input, f'{idx_planilha}: Chamado {CHAMADO}/{PN} inserido com sucesso.\n')
                    idx_planilha += 1
                    
                except Exception as e:
                    log_message(self.log_input, f'Não foi possível abrir o arquivo {idx_planilha}: {e}')

                    idx_planilha += 1
            
            # Criando uma cópia das colunas selecionadas do DataFrame original
            planilha_copia = planilha_df.iloc[:, [0, 1, 2]].copy()

            # Renomeando as colunas da cópia
            planilha_copia.columns = ['CHAMADO', 'REF NF', 'PART NUMBER']

            # Adicionando as colunas ao DataFrame a partir do dicionário
            planilha_copia['VALOR'] = [
                resultados_hp.get(f"{row[2]}_{row[9]}", {}).get('VALOR', None)
                for _, row in planilha_df.iterrows()
            ]

            planilha_copia['CST'] = [
                resultados_hp.get(f"{row[2]}_{row[9]}", {}).get('CST', None)
                for _, row in planilha_df.iterrows()
            ]

            planilha_copia['NCM'] = [
                resultados_hp.get(f"{row[2]}_{row[9]}", {}).get('NCM', None)
                for _, row in planilha_df.iterrows()
            ]

            planilha_copia['QTD.'] = [
                resultados_hp.get(f'{row[2]}_{row[9]}', {}).get('QTD.', None)
                for _, row in planilha_df.iterrows()
            ]

            # Substituindo pontos por vírgulas na coluna VALOR
            planilha_copia['VALOR'] = planilha_copia['VALOR'].astype(str).str.replace('.', ',', regex = False)


            # Exportando o DataFrame para Excel
            planilha_copia.to_excel('Planilha HP valores_devolução formatada.xlsx', index = False)

            if os.path.exists(f'{download_dir}\\Planilha HP valores_devolução formatada.xlsx'):
                os.remove(f'{download_dir}\\Planilha HP valores_devolução formatada.xlsx')

            shutil.move('Planilha HP valores_devolução formatada.xlsx', download_dir)

            log_message(self.log_input, f'Planilha criada com sucesso: Planilha HP valores_devolução formatada.xlsx')
            
        def parar(instance):
            popup.dismiss()
            return

        botao_sim.bind(on_release = continuar)
        botao_nao.bind(on_release = parar)
    
        # Exibe o popup
        popup.open()

    mostrar_confirmacao(self)


