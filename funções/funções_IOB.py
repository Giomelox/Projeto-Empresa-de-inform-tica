import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from funções.funções_Gerais import log_message, usuario_IOB, senha_IOB
from funções.funções_conectar_email import aux_path_XML_destino, planilha_df

# Botoes_entrada_Dell
def importar_produtos(self, *args):

    driver = webdriver.Chrome()
    
    idx = 1

    # Acessa a página desejada
    driver.get('https://sso.iob.com.br/signin/?response_type=code&scope=&client_id=c17d4225-9d57-401b-b4fd-32503121f55b&redirect_uri=https://emissor.iob.com.br')

    try:

        element_usuario_emitir = WebDriverWait(driver, 200).until( # Busca pela página o campo de usuário
            EC.presence_of_element_located((By.ID, 'username'))
        )
        element_usuario_emitir.send_keys(usuario_IOB) # Insere o nome de usuário

        time.sleep(2.5)

        element_senha_emitir = WebDriverWait(driver, 50).until( # Busca pela página o campo de senha
            EC.presence_of_element_located((By.ID, 'password'))
        )
        element_senha_emitir.send_keys(senha_IOB) # Insere a senha

        time.sleep(2.5) # tempo para carregar o reCAPTCHA

        element_reCAPTCHA_emitir = WebDriverWait(driver, 50).until( # Buscar o elemento reCAPTCHA na página
            EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[2]/div/div[2]/main/div[2]/div/form/div[4]/div/div/div/iframe'))
        )
        driver.switch_to.frame(element_reCAPTCHA_emitir) # Muda para o contexto do iframe onde o captcha está inserido
        captcha_checkbox = driver.find_element(By.XPATH, '//div[@class="recaptcha-checkbox-border"]') # Busca o elemento clicável do captcha
        captcha_checkbox.click() # Clica no elemento captcha

        driver.switch_to.default_content() # Muda para o contexto padrão da página

        time.sleep(10)

        element_submit_emitir = WebDriverWait(driver, 50).until( # Busca pelo elemento submit na página
            EC.presence_of_element_located((By.XPATH, '//*[@id="formButton"]'))
        )
        element_submit_emitir.click()  # Clica no elemento submit

        element_emissão_notas_emitir = WebDriverWait(driver, 200).until( # Busca pelo botao de emissão de notas
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[title="Emissão de Notas"]'))
        )

        element_form = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'btn-2'))
        )
        try:
            element_form.click()
        except:
            pass
                
        if element_emissão_notas_emitir:
            driver.get('https://emissor2.iob.com.br/notafiscal/incoming_nfes') # Caso encontre o botão de notas, entra na página para emitir

        try:
            element_form.click()
        except:
            pass

        try:
            for x in os.listdir(aux_path_XML_destino):
                # Remover o overlay antes de qualquer clique
                try:
                    driver.execute_script("""
                        let overlay = document.getElementById('UIDialogLayer');
                        if (overlay) {
                            overlay.remove();  // Remove completamente o overlay da página
                        }
                    """)
                    time.sleep(0.5)  # Pequena pausa para garantir que o overlay seja removido
                except Exception as e:
                    print("Não foi possível remover o overlay:", e)

                # Iniciar o processo de importação
                element_importar = WebDriverWait(driver, 200).until(
                    EC.presence_of_element_located((By.XPATH, '//*[contains(@class, "UILink button pull-right")]'))
                )
                time.sleep(0.5)
                element_importar.click()

                # Aguardar e clicar no input de arquivo
                element_abrir_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="import_nfe_nfe_xml"]'))
                )
                caminho = os.path.join(aux_path_XML_destino, x)
                element_abrir_input.send_keys(caminho)

                # Aguardar o botão de "Confirmar" (ou similar) aparecer após o upload
                element_confirmar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[contains(@class, "save button primary")]'))
                )
                element_confirmar.click()

                # Navegar de volta para a página inicial (após enviar um arquivo)
                driver.get('https://emissor2.iob.com.br/notafiscal/incoming_nfes')

                # Aguardar a página carregar e clicar novamente no botão "Importar"
                element_importar1 = WebDriverWait(driver, 200).until(
                    EC.presence_of_element_located((By.XPATH, '//*[contains(@class, "UILink button pull-right")]'))
                )
                time.sleep(0.5)
                element_importar1.click()

                log_message(self.log_input, f'{idx}: {x} importado com sucesso.')

                idx += 1

                # Aguardar um breve momento para garantir que a próxima iteração não aconteça muito rápido
                time.sleep(2)

        except Exception as e:
            log_message(self.log_input, f'Erro encontrado: {e}')

    except Exception as e:
        log_message(self.log_input, f'Ocorreu um erro: {e}')

    finally:
        driver.close()

def emitir_nf_circulação(self, *args):

    driver = webdriver.Chrome()

    actions = ActionChains(driver)

    # Acessa a página desejada
    driver.get('https://sso.iob.com.br/signin/?response_type=code&scope=&client_id=c17d4225-9d57-401b-b4fd-32503121f55b&redirect_uri=https://emissor.iob.com.br')
    
    try:
        element_usuario_emitir = WebDriverWait(driver, 200).until( # Busca pela página o campo de usuário
            EC.presence_of_element_located((By.ID, 'username'))
        )
        element_usuario_emitir.send_keys(usuario_IOB) # Insere o nome de usuário

        time.sleep(2.5)

        element_senha_emitir = WebDriverWait(driver, 50).until( # Busca pela página o campo de senha
            EC.presence_of_element_located((By.ID, 'password'))
        )
        element_senha_emitir.send_keys(senha_IOB) # Insere a senha

        time.sleep(2.5) # tempo para carregar o reCAPTCHA

        element_reCAPTCHA_emitir = WebDriverWait(driver, 50).until( # Buscar o elemento reCAPTCHA na página
            EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[2]/div/div[2]/main/div[2]/div/form/div[4]/div/div/div/iframe'))
        )
        driver.switch_to.frame(element_reCAPTCHA_emitir) # Muda para o contexto do iframe onde o captcha está inserido
        captcha_checkbox = driver.find_element(By.XPATH, '//div[@class="recaptcha-checkbox-border"]') # Busca o elemento clicável do captcha
        captcha_checkbox.click() # Clica no elemento captcha

        driver.switch_to.default_content() # Muda para o contexto padrão da página

        time.sleep(20)

        element_submit_emitir = WebDriverWait(driver, 50).until( # Busca pelo elemento submit na página
            EC.presence_of_element_located((By.XPATH, '//*[@id="formButton"]'))
        )
        element_submit_emitir.click()  # Clica no elemento submit

        element_emissão_notas_emitir = WebDriverWait(driver, 200).until( # Busca pelo botao de emissão de notas
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[title="Emissão de Notas"]'))
        )

        element_form = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'btn-2'))
        )
        try:
            element_form.click()
        except:
            pass
                
        if element_emissão_notas_emitir:
            driver.get('https://emissor2.iob.com.br/notafiscal/nfes') # Caso encontre o botão de notas, entra na página para emitir

        try:
            element_form.click()
        except:
            pass

        element_nova_nota_emitir = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
        )
        element_nova_nota_emitir.click()

    except Exception as e:
        print(f'não foi possível entrar no site: {e}')
        driver.quit()

    idx = 0

    planilha_df.ffill(inplace = True)

    ler_planilha_mesclada = planilha_df.groupby('CHAMADO', as_index = False, sort = False).agg({
        'NF OU XML PARA BAIXAR (PROGRAMA)': lambda x: ', '.join(map(str, filter(pd.notna, x))),
        'PN CAIXA.': lambda x: ', '.join(x),
        'VALOR': lambda x: ', '.join(map(str, x)),
        'CST': lambda x: ', '.join(map(str, x)),
        'NCM': lambda x: ', '.join(map(str, x)),
        'QTD.': lambda x: ', '.join(map(str, x))
    })

    for index, row in ler_planilha_mesclada.iloc[:].iterrows(): # Lê todas as linhas da planilha com as colunas selecionadas
        CHAMADO = str(row['CHAMADO']).replace('.0', '').strip()
        nf_lista = str(row['NF OU XML PARA BAIXAR (PROGRAMA)']).replace('.0', '').strip().split(', ')
        pn_lista = str(row['PN CAIXA.']).strip().split(', ')
        valor_lista = str(row['VALOR']).replace('.', ',').strip().split(', ')
        cst_lista = str(row['CST']).split(', ')
        ncm_lista = str(row['NCM']).strip().split(', ')
        qtd_lista = str(row['QTD.']).strip().split(', ')

        try:

            try:
                element_contingencia = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//a[contains(text(), "Aguardar")]'))
                )
                element_contingencia.click()
            except:
                pass

            # Encontra o campo de cliente
            element_cliente_emitir = WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_client_id'))
            )
            element_cliente_emitir.send_keys('CLEDSON BATISTA NASCIMENTO BARROS') # Insere no campo cliente

            element_results_client_emitir = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results')) # Lê as opções dadas pelo site com a lista de clientes
            )

            element_first_li_cliente_emitir = element_results_client_emitir.find_element(By.CSS_SELECTOR, 'ul li:first-child') # Seleciona o primeiro cliente que aparece
            element_first_li_cliente_emitir.click()

            # Encontra o campo de natureza
            element_natureza_emitir = WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_nature_id'))
            )
            element_natureza_emitir.clear() # Limpa o que estiver escrito no campo natureza

            element_natureza_emitir.send_keys('Mercadoria de Terceiro aplicada em Garantia à sua ordem') # Insere no campo natureza

            element_results_natureza_emitir = WebDriverWait(driver, 50).until( # Lê as opções dadas pelo site com a lista de natureza
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results'))
            )

            element_first_li_natureza_emitir = element_results_natureza_emitir.find_element(By.CSS_SELECTOR, 'ul li:first-child') # Seleciona a primeira natureza que aparece
            element_first_li_natureza_emitir.click()

            idx_element = 1

            idx_qtd = 0
            
            for pn, valor, cst, ncm, qtd in zip(pn_lista, valor_lista, cst_lista, ncm_lista, qtd_lista):
                
                # Busca pelo campo para inserir o produto
                element_produto = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'nfe_product_id'))
                )

                driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_produto)

                # Insere o produto que está na planilha
                element_produto.send_keys(pn) 

                time.sleep(1)

                # Encontra a lista de opções do produto
                element_produto_results = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results')) 
                )

                # Seleciona o primeiro produto que aparece                                          
                element_first_li_produto_emitir = element_produto_results.find_element(By.CSS_SELECTOR, f'ul li:first-child') 
                element_first_li_produto_emitir.click() 

                time.sleep(1.5)
                   
                try:

                    element_produto_selecionar = WebDriverWait(driver, 50).until( 
                        EC.element_to_be_clickable((By.XPATH, f'/html/body/div[4]/section/div/div[3]/form/div[1]/div[9]/div/div[7]/div/div[1]/div/div/table/tbody/tr[{idx_element}]/td[4]/div/span/input')) 
                    )
                    driver.execute_script("""
                        arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
                        window.scrollBy(0, 230);  // Ajusta o scroll um pouco para baixo (230px)
                    """, element_produto_selecionar)

                    element_produto_selecionar.click()

                    # Encontra o elemento NCM do produto e envia os dados da planilha
                    element_NCM = WebDriverWait(driver, 50).until(
                        EC.presence_of_element_located((By.ID, 'item_product_ncm_id'))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_NCM)
                    element_NCM.clear()
                    time.sleep(1.5)

                    element_NCM.send_keys(ncm)

                    time.sleep(3)

                    # Lê as opções dadas pelo site com a lista de NCM
                    
                    element_NCM = WebDriverWait(driver, 50).until( 
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results'))
                    )
                    time.sleep(2)

                    # Seleciona o primeiro NCM que aparece
                    element_NCM_prod = element_NCM.find_element(By.CSS_SELECTOR, 'ul li:first-child') 
                    element_NCM_prod.click()
                    actions.send_keys(Keys.TAB).perform()

                    element_CFOP = Select(driver.find_element(By.ID, 'item_cfop'))
            
                    element_CFOP.select_by_value('5949')

                    # Altera o display da informação ''quantidade'' do produto para ser possível acessar pelo selenium
                    element_quantidade_comercial = WebDriverWait(driver, 50).until(
                        EC.presence_of_element_located((By.ID, 'item_quantity'))
                    )
                    driver.execute_script("""
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input'));
                        arguments[0].dispatchEvent(new Event('change'));
                        """, element_quantidade_comercial, qtd)

                    # Altera o display da informação ''valor unitário'' do produto para ser possível acessar pelo selenium
                    driver.execute_script("document.getElementById('item_unit_price').style.display = 'block';") 
                    element_valor_produto = WebDriverWait(driver, 50).until( 
                        EC.presence_of_element_located((By.ID, 'item_unit_price'))
                    )

                    # Força a inserção do valor no site
                    driver.execute_script("""
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input'));
                        arguments[0].dispatchEvent(new Event('change'));
                        """, element_valor_produto, valor)
                    actions.send_keys(Keys.TAB).perform()

                    element_valor_taxa = WebDriverWait(driver, 50).until( 
                        EC.presence_of_element_located((By.ID, 'item_taxable_price'))
                    )
                    element_valor_taxa.clear()

                    time.sleep(0.5) # Tempo para salvar o valor do produto

                    # Encontra o campo de ICMS e clica nele
                    element_campo_ICMS = WebDriverWait(driver, 50).until( 
                        EC.presence_of_element_located((By.XPATH, '//*[@id="proxy-reserve"]/div[9]/div/form/div[3]/ul/li[2]/a'))
                    )
                    element_campo_ICMS.click()

                    # Encontra o campo ORIG e altera para o orig do produto
                    element_orig = Select(driver.find_element(By.XPATH, f'//*[@id="item_icms_origin"]'))

                    element_orig.select_by_value(cst)

                    # Encontra o campo para alterar o CST do produto para 400
                    element_cst = Select(driver.find_element(By.ID, 'item_icms_cst'))

                    element_cst.select_by_value('400')

                    element_submit_produto = driver.find_element(By.XPATH, '//*[@id="proxy-reserve"]/div[9]/div/form/div[4]/div/span[1]/button') # Encontra o botão de salvar as informações do produto e clica
                    driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_submit_produto)
                    driver.execute_script("arguments[0].click();", element_submit_produto)

                    idx_qtd += 1

                    idx_element += 1
            
                except Exception as e:
                    print(f'{idx}: Erro ao tentar inserir um dado do produto {CHAMADO}/{pn}')

                    driver.get('https://emissor2.iob.com.br/notafiscal/nfes')

                    driver.refresh()

                    element_nova_nota_emitir1 = WebDriverWait(driver, 200).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
                    )
                    element_nova_nota_emitir1.click()   

                    continue

            # Encontra e clica na modalidade de frete
            element_frete =  WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_carriage_modality'))
            )
            time.sleep(0.5)
            driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_frete)
            time.sleep(0.5)
            element_frete.click()

            # Encontra dentro das opções, o frete do processo

            element_frete_terceiros = Select(driver.find_element(By.ID, 'nfe_carriage_modality'))
            
            element_frete_terceiros.select_by_value('9')

            # Encontra o bloco de quantidade de produtos e insere a quantidade 1
            element_quantidade = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'nfe_carriage_quantity')) 
            )
            element_quantidade.clear()
            element_quantidade.send_keys(idx_qtd)
            actions.send_keys(Keys.TAB).perform()

            # Encontra o bloco de espécie de produtos e insere 'volume'
            element_especie = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'nfe_carriage_kind'))
            )
            element_especie.send_keys('CAIXA')
            actions.send_keys(Keys.TAB).perform() 

            time.sleep(0.5) # Tempo para processar o peso bruto

            # Encontra o campo de observação e adiciona as informações de acordo com a planilha
            element_observação = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.ID, 'nfe_notes'))
            )

            element_observação.send_keys(f'@ENV ;@REF {CHAMADO} REF NF {nf_lista} PN {pn_lista}\n I - "DOCUMENTO EMITIDO POR ME OU EPP OPTANTE PELO SIMPLES NACIONAL";II - "NÃO GERA DIREITO A CRÉDITO FISCAL DE ICMS, DE ISS E DE IPI".')

            # Encontra o bloco de pagamento e seleciona o 90
            element_formaPagamento = Select(driver.find_element(By.ID, 'nfe_payment_method'))
            element_formaPagamento.select_by_value('90')

            # Encontra o bloco de valor total e zera
            element_pagamento_total = driver.find_element(By.ID, 'nfe_total_information_attributes_paid_total')
            element_pagamento_total.clear()
            actions.send_keys(Keys.TAB).perform() 

            time.sleep(1.5)

            # Encontra o elemento para emitir a nota e clica
            element_emitir_nota = driver.find_element(By.ID, 'save-and-submit-nfe')
            driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_emitir_nota)
            time.sleep(1)
            element_emitir_nota.click()

            time.sleep(1) # Aguarda até aparecer a contingencia ou confirmar a NF
        
            element_contingencia1 = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.XPATH, '//a[contains(text(), "Aguardar")]'))
            )

            try:
                element_contingencia1.click()
            except:
                pass

            # Confirma a emissão da NF
            element_confirmar_emissão = WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="proxy-reserve"]/div[12]/div[2]/span/button'))
            )
            element_confirmar_emissão.click()

            element_carregar = WebDriverWait(driver, 100).until( # Aguarda o elemento do carregamento da NF desaparecer para continuar
                EC.invisibility_of_element_located((By.XPATH, '//*[@id="proxy-reserve"]/div[5]/div[2]/div/div/a'))
            )

            if element_carregar:
                driver.get('https://emissor2.iob.com.br/notafiscal/nfes')

            driver.refresh()

            element_nova_nota_emitir2 = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
            )
            element_nova_nota_emitir2.click()

            idx += 1
            
            print(f'{idx}: {CHAMADO}/{pn_lista} inseridos com sucesso no site.')

            time.sleep(2.5)

        except Exception as e:
            print(f'Não foi possível inserir os dados da planilha no site:\n{idx}: {CHAMADO}/{pn_lista}')

            driver.get('https://emissor2.iob.com.br/notafiscal/nfes')

            driver.refresh()

            element_nova_nota_emitir3 = WebDriverWait(driver, 200).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
            )
            element_nova_nota_emitir3.click()       

            continue
        
    time.sleep(5)

    driver.close()

def emitir_nf_entrada_tec(self, *args):
    pass

def emitir_NF_dev_dell(self, *args):

    idx = 0

    driver = webdriver.Chrome()

    actions = ActionChains(driver)

    # Acessa a página desejada
    driver.get('https://sso.iob.com.br/signin/?response_type=code&scope=&client_id=c17d4225-9d57-401b-b4fd-32503121f55b&redirect_uri=https://emissor.iob.com.br')

    try:
        element_usuario_emitir = WebDriverWait(driver, 200).until( # Busca pela página o campo de usuário
            EC.presence_of_element_located((By.ID, 'username'))
        )
        element_usuario_emitir.send_keys(usuario_IOB) # Insere o nome de usuário

        time.sleep(2.5)

        element_senha_emitir = WebDriverWait(driver, 50).until( # Busca pela página o campo de senha
            EC.presence_of_element_located((By.ID, 'password'))
        )
        element_senha_emitir.send_keys(senha_IOB) # Insere a senha

        time.sleep(2.5) # tempo para carregar o reCAPTCHA

        element_reCAPTCHA_emitir = WebDriverWait(driver, 50).until( # Buscar o elemento reCAPTCHA na página
            EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[2]/div/div[2]/main/div[2]/div/form/div[4]/div/div/div/iframe'))
        )
        driver.switch_to.frame(element_reCAPTCHA_emitir) # Muda para o contexto do iframe onde o captcha está inserido
        captcha_checkbox = driver.find_element(By.XPATH, '//div[@class="recaptcha-checkbox-border"]') # Busca o elemento clicável do captcha
        captcha_checkbox.click() # Clica no elemento captcha

        driver.switch_to.default_content() # Muda para o contexto padrão da página

        time.sleep(20)

        element_submit_emitir = WebDriverWait(driver, 50).until( # Busca pelo elemento submit na página
            EC.presence_of_element_located((By.XPATH, '//*[@id="formButton"]'))
        )
        element_submit_emitir.click()  # Clica no elemento submit

        element_emissão_notas_emitir = WebDriverWait(driver, 200).until( # Busca pelo botao de emissão de notas
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[title="Emissão de Notas"]'))
        )

        element_form = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'btn-2'))
        )
        try:
            element_form.click()
        except:
            pass
                
        if element_emissão_notas_emitir:
            driver.get('https://emissor2.iob.com.br/notafiscal/nfes') # Caso encontre o botão de notas, entra na página para emitir

        try:
            element_form.click()
        except:
            pass

        element_nova_nota_emitir = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
        )
        element_nova_nota_emitir.click()

    except Exception as e:
        log_message(self.log_input, f'não foi possível entrar no site: {e}')
        driver.quit() 

    for index, row in planilha_df.iloc[:].iterrows(): # Lê todas as linhas da planilha com as colunas selecionadas
        CHAMADO = str(row[0]).strip()
        NF = str(row[1]).strip()
        PN = str(row[2]).strip()
        PN_CAIXA = str(row[3]).strip()
        VALOR = str(row[5]).strip().replace('.', ',')
        CST = row[6]
        NCM = str(row[7]).strip() 
        QTD = str(row[8]).strip()
        STATUS_planilha = str(row[9]).strip()

        try:

            element_contingencia = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//a[contains(text(), "Aguardar")]'))
            )

            try: 
                element_contingencia.click()
            except:
                pass

            # Encontra o campo de cliente
            element_cliente_emitir = WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_client_id'))
            )
            element_cliente_emitir.send_keys('DELL COMPUTADORES DO BRASIL LTDA') # Insere no campo cliente

            element_results_client_emitir = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results')) # Lê as opções dadas pelo site com a lista de clientes
            )

            element_first_li_cliente_emitir = element_results_client_emitir.find_element(By.CSS_SELECTOR, 'ul li:first-child') # Seleciona o primeiro cliente que aparece
            element_first_li_cliente_emitir.click()

            # Encontra o campo de natureza
            element_natureza_emitir = WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_nature_id'))
            )
            element_natureza_emitir.clear() # Limpa o que estiver escrito no campo natureza

            if STATUS_planilha == 'GOOD': # Se o produto for GOOD:
                element_natureza_emitir.send_keys('DEVOLUÇÃO REMESSA DE SUBSTITUIÇÃO EM GARANTIA') # Insere no campo natureza

                element_results_natureza_emitir = WebDriverWait(driver, 50).until( # Lê as opções dadas pelo site com a lista de natureza
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results'))
                )

                element_first_li_natureza_emitir = element_results_natureza_emitir.find_element(By.CSS_SELECTOR, 'ul li:first-child') # Seleciona a primeira natureza que aparece
                element_first_li_natureza_emitir.click()

            else:
                element_natureza_emitir.send_keys('RETORNO APÓS SUBSTITUIÇÃO EM GARANTIA') # Insere no campo natureza

                element_results_natureza_emitir = WebDriverWait(driver, 50).until( # Lê as opções dadas pelo site com a lista de natureza
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results')) 
                )

                element_first_li_natureza_emitir = element_results_natureza_emitir.find_element(By.CSS_SELECTOR, 'ul li:first-child') # Seleciona a primeira natureza que aparece
                element_first_li_natureza_emitir.click()
            
            # Busca pelo campo para inserir o produto
            element_produto = WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_product_id'))
            )

            # Insere o produto que está na planilha
            element_produto.send_keys(f'{PN}') 

            # Encontra a lista de opções do produto
            element_produto_results = WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results')) 
            )

            # Seleciona o primeiro produto que aparece
            element_first_li_produto_emitir = element_produto_results.find_element(By.XPATH, '//*[@id="new_nfe"]/div[1]/div[9]/div/div[7]/div/div[1]/div/div/div[1]/div/div[2]/div/span/div/ul/li[1]') 
            element_first_li_produto_emitir.click()

            time.sleep(1.5)
            
            try:
                # Clica no produto para alterar as informações
                element_produto_selecionar = WebDriverWait(driver, 50).until( 
                    EC.element_to_be_clickable((By.XPATH, "//input[contains(@id, '_product_name')]")) 
                )
                element_produto_selecionar.click()

                # Encontra o elemento NCM do produto e envia os dados da planilha
                element_NCM = WebDriverWait(driver, 50).until(
                    EC.presence_of_element_located((By.ID, 'item_product_ncm_id'))
                )
                element_NCM.clear()
                time.sleep(0.5)

                element_NCM.send_keys(f'{NCM}')

                time.sleep(4)

                # Lê as opções dadas pelo site com a lista de NCM
                
                element_NCM = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results'))
                )
                time.sleep(0.5)

                # Seleciona o primeiro NCM que aparece
                element_NCM_prod = element_NCM.find_element(By.CSS_SELECTOR, 'ul li:first-child') 
                element_NCM_prod.click()
                actions.send_keys(Keys.TAB).perform()

                element_CFOP = Select(driver.find_element(By.ID, 'item_cfop'))
            
                element_CFOP.select_by_value('6949')

                # Altera o display da informação ''quantidade'' do produto para ser possível acessar pelo selenium
                element_quantidade_comercial = WebDriverWait(driver, 50).until(
                    EC.presence_of_element_located((By.ID, 'item_quantity'))
                )
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input'));
                    arguments[0].dispatchEvent(new Event('change'));
                    """, element_quantidade_comercial, QTD)

                # Altera o display da informação ''valor unitário'' do produto para ser possível acessar pelo selenium
                driver.execute_script("document.getElementById('item_unit_price').style.display = 'block';") 
                element_valor_produto = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'item_unit_price'))
                )

                # Força a inserção do valor no site
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input'));
                    arguments[0].dispatchEvent(new Event('change'));
                    """, element_valor_produto, VALOR)
                actions.send_keys(Keys.TAB).perform()

                element_valor_taxa = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'item_taxable_price'))
                )
                element_valor_taxa.clear()

                time.sleep(0.5) # Tempo para salvar o valor do produto

                # Encontra o campo de ICMS e clica nele
                element_campo_ICMS = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.XPATH, '//*[@id="proxy-reserve"]/div[9]/div/form/div[3]/ul/li[2]/a'))
                )
                element_campo_ICMS.click()

                # Encontra o campo ORIG e altera para o orig do produto
                element_orig = Select(driver.find_element(By.XPATH, f'//*[@id="item_icms_origin"]'))

                element_orig.select_by_value(f'{CST}')

                # Encontra o campo para alterar o CST do produto para 400
                element_cst = Select(driver.find_element(By.ID, 'item_icms_cst'))

                element_cst.select_by_value('400')

                element_submit_produto = driver.find_element(By.XPATH, '//*[@id="proxy-reserve"]/div[9]/div/form/div[4]/div/span[1]/button') # Encontra o botão de salvar as informações do produto e clica
                driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_submit_produto)
                driver.execute_script("arguments[0].click();", element_submit_produto)
            
            except Exception as e:
                log_message(self.log_input, f'{idx}: Erro ao tentar inserir um dado do produto {CHAMADO}/{PN}')

                driver.get('https://emissor2.iob.com.br/notafiscal/nfes')

                driver.refresh()

                element_nova_nota_emitir1 = WebDriverWait(driver, 200).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
                )
                element_nova_nota_emitir1.click()   

                continue

            # Encontra e clica na modalidade de frete
            element_frete =  WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_carriage_modality'))
            )
            driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_frete)
            element_frete.click()
            
            element_frete_terceiros = Select(driver.find_element(By.ID, 'nfe_carriage_modality'))
            
            element_frete_terceiros.select_by_value('1')

            # Encontra o bloco de selecionar a transportadora e insere no campo de transportadora
            element_transportadora = WebDriverWait(driver, 50).until( 
                EC.presence_of_element_located((By.ID, 'nfe_carrier_id'))
            )
            element_transportadora.send_keys('TNT MERCURIO CARGAS E ENCOMENDAS EXPRESSAS LTDA.') 

            time.sleep(1)

            # Lê as opções dadas pelo site com a lista de transportadora
            element_transportadora = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.results')) 
            )

            time.sleep(1)

            # Seleciona a primeira transportadora que aparece
            element_transportadora_tnt = element_transportadora.find_element(By.CSS_SELECTOR, 'ul li:first-child') 
            element_transportadora_tnt.click()

            # Encontra o bloco de quantidade de produtos e insere a quantidade 1
            element_quantidade = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'nfe_carriage_quantity')) 
            )
            element_quantidade.clear()
            element_quantidade.send_keys('1')
            actions.send_keys(Keys.TAB).perform()

            # Encontra o bloco de espécie de produtos e insere 'volume'
            element_especie = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'nfe_carriage_kind'))
            )
            element_especie.send_keys('VOLUME')
            actions.send_keys(Keys.TAB).perform() 

            time.sleep(0.5) # Tempo para processar o peso bruto

            # Encontra o bloco de peso bruto de produtos e insere a quantidade 1
            PESO_BRUTO = 1

            element_peso_bruto = WebDriverWait(driver, 50).until( 
                    EC.presence_of_element_located((By.ID, 'nfe_carriage_gross_weight')) 
            )
            driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input'));
                    arguments[0].dispatchEvent(new Event('change'));
                    """, element_peso_bruto, PESO_BRUTO)
            time.sleep(0.5)
            actions.send_keys(Keys.TAB).perform()

            # Encontra o campo de observação e adiciona as informações de acordo com a planilha
            element_observação = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.ID, 'nfe_notes'))
            )

            if PN == PN_CAIXA:
                element_observação.send_keys(f'@DEV ;@REF {CHAMADO} REF NF {NF} PN {PN} {STATUS_planilha}\n I - "DOCUMENTO EMITIDO POR ME OU EPP OPTANTE PELO SIMPLES NACIONAL";II - "NÃO GERA DIREITO A CRÉDITO FISCAL DE ICMS, DE ISS E DE IPI".')
            else:
                element_observação.send_keys(f'@DEV ;@REF {CHAMADO} REF NF {NF} PN {PN} SOFREU ALTERAÇÃO PARA {PN_CAIXA} {STATUS_planilha}\n I - "DOCUMENTO EMITIDO POR ME OU EPP OPTANTE PELO SIMPLES NACIONAL";II - "NÃO GERA DIREITO A CRÉDITO FISCAL DE ICMS, DE ISS E DE IPI".')

            time.sleep(3)

            # Encontra o bloco de pagamento e seleciona o 90
            element_formaPagamento = Select(driver.find_element(By.ID, 'nfe_payment_method'))
            element_formaPagamento.select_by_value('90')

            # Encontra o bloco de valor total e zera
            element_pagamento_total = driver.find_element(By.ID, 'nfe_total_information_attributes_paid_total')
            element_pagamento_total.clear()
            actions.send_keys(Keys.TAB).perform() 

            time.sleep(0.5)

            # Encontra o elemento para emitir a nota e clica
            element_emitir_nota = driver.find_element(By.ID, 'save-and-submit-nfe')
            driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element_emitir_nota)
            time.sleep(1)
            element_emitir_nota.click()

            time.sleep(1) # Aguarda até aparecer a contingencia ou confirmar a NF

            element_contingencia1 = WebDriverWait(driver, 50).until(
                    EC.presence_of_element_located((By.XPATH, '//a[contains(text(), "Aguardar")]'))
                )
            
            try:
                element_contingencia1.click()
            except:
                pass

            # Confirma a emissão da NF
            element_confirmar_emissão = WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="proxy-reserve"]/div[12]/div[2]/span/button'))
            )
            element_confirmar_emissão.click()

            WebDriverWait(driver, 100).until( # Aguarda o elemento do carregamento da NF desaparecer para continuar
                EC.invisibility_of_element_located((By.XPATH, '//*[@id="proxy-reserve"]/div[5]/div[2]/div/div/a'))
            )

            # Encontra o elemento para baixar o XML e clica
            element_baixar_xml = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-main"]/div[3]/div/a[4]'))
            )
            element_baixar_xml.click()

            time.sleep(2.5) # Tempo para baixar o arquivo

            # Encontra o elemento para ir para danfe e clica
            element_ir_para_danfe = WebDriverWait(driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-main"]/div[3]/div/a[3]'))
            )
            element_ir_para_danfe.click()

            # Encontra o elemento para baixar o PDF e clica
            element_baixar_pdf = WebDriverWait(driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-main"]/a[2]'))
            )
            element_baixar_pdf.click()

            time.sleep(2.5) # Tempo para baixar o arquivo

            if element_baixar_pdf:
                driver.get('https://emissor2.iob.com.br/notafiscal/nfes') # Caso baixe o pdf, entra na página para emitir novamente

            element_nova_nota_emitir2 = WebDriverWait(driver, 200).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
            )
            element_nova_nota_emitir2.click()

            idx += 1
            
            log_message(self.log_input, f'{idx}: {CHAMADO}/{PN} inseridos com sucesso no site.')

            time.sleep(2.5)

        except Exception as e:
            log_message(self.log_input, f'Não foi possível inserir os dados da planilha no site:\n{idx}: {CHAMADO}/{PN}')

            driver.get('https://emissor2.iob.com.br/notafiscal/nfes')

            driver.refresh()

            element_nova_nota_emitir3 = WebDriverWait(driver, 200).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="UIContainer button-grid"] a[class="UILink button primary with_margin"]'))
            )
            element_nova_nota_emitir3.click()       

            continue
        
    time.sleep(5)

    driver.quit()