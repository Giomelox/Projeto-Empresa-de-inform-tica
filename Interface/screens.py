from kivy.config import Config

Config.set('graphics', 'width', '1000')
Config.set('graphics', 'height', '800')

from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.uix.scrollview import ScrollView
from kivy.uix.screenmanager import Screen, SlideTransition
from kivy.uix.textinput import TextInput
from kivy.graphics import Color, RoundedRectangle
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.button import Button
from funções.funções_conectar_email import baixar_arquivosXML_DELL, baixar_arquivosXML_HP
from funções.funções_criar_planilhas import criar_planilha_entrada_nf_DELL, criar_planilha_entrada_nf_HP, criar_planilha_difal
from funções.funções_Gerais import log_message, validar_usuario, obter_usuarios_validos, SolicitarPlanilha
from funções.funções_elogistc import valores_devolução_DELL, valores_devolução_HP, biparxml
from funções.funções_IOB import importar_produtos, emitir_NF_dev_dell, emitir_nf_circulação

class BorderedButton(Button):
    def __init__(self, **kwargs):
        super(BorderedButton, self).__init__(**kwargs)
        
        # Remove o fundo padrão do botão
        self.background_normal = ''
        self.background_color = (0, 0, 0, 0)  # Fundo transparente
        
        # Variável para armazenar a cor original
        self.default_color = (1, 1, 1, 1)  # Cor de fundo normal
        self.pressed_color = (0.4, 1, 0.9, 1)  # Cor ao clicar

        # Atualiza o botão quando ele muda de posição ou tamanho
        self.bind(pos = self.update_rect, size = self.update_rect)

        # Muda a cor ao clicar
        self.bind(on_press = self.on_press_button, on_release = self.on_release_button)

    def update_rect(self, *args):
        # Limpa o canvas existente antes de desenhar o fundo
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.default_color)  # Cor de fundo normal do botão
            # Desenha o fundo arredondado
            self.bg_rect = RoundedRectangle(pos = self.pos, size = self.size, radius = [20])

    def on_press_button(self, *args):
        # Muda a cor de fundo ao pressionar o botão
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.pressed_color)  # Cor ao clicar
            RoundedRectangle(pos = self.pos, size = self.size, radius = [20])
    
    def on_release_button(self, *args):
        # Volta para a cor de fundo original ao soltar o botão
        self.update_rect()

class BorderedButton_top(Button):
    def __init__(self, **kwargs):
        super(BorderedButton_top, self).__init__(**kwargs)
        
        # Remove o fundo padrão do botão
        self.background_normal = ''
        self.background_color = (0, 0, 0, 0)  # Fundo transparente
        
        # Variável para armazenar a cor original
        self.default_color = (1, 0.1, 0.1, 1)  # Cor de fundo normal
        self.pressed_color = (1, 0.4, 0.4, 1)  # Cor ao clicar

        # Atualiza o botão quando ele muda de posição ou tamanho
        self.bind(pos = self.update_rect, size = self.update_rect)

        # Muda a cor ao clicar
        self.bind(on_press = self.on_press_button, on_release = self.on_release_button)

    def update_rect(self, *args):
        # Limpa o canvas existente antes de desenhar o fundo
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.default_color)  # Cor de fundo normal do botão
            # Desenha o fundo arredondado
            self.bg_rect = RoundedRectangle(pos = self.pos, size = self.size, radius = [20])

    def on_press_button(self, *args):
        # Muda a cor de fundo ao pressionar o botão
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.pressed_color)  # Cor ao clicar
            RoundedRectangle(pos = self.pos, size = self.size, radius = [20])
    
    def on_release_button(self, *args):
        # Volta para a cor de fundo original ao soltar o botão
        self.update_rect()

# Criação da interface Kivy
class Tela_autenticação(Screen):
    def ir_para_MainMenu(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'MainMenu'
        
    def __init__(self, **kwargs):
        super(Tela_autenticação, self).__init__(**kwargs)

        autenticação_layout = FloatLayout()

        imagemMatec = Image(source = 'imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})
         
        self.label_email_autenticar = Label(text = 'Digite seu email para entrar:', font_size = '20sp', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.5, 'y': 0.6}, color = (1,1,1,1))
        self.email_autenticar = TextInput(font_size = '20sp', size_hint = (0.4,0.07), pos_hint = {'center_x': 0.5, 'y': 0.55}, multiline = False, write_tab = False, halign = 'center', padding_y = (15, 15))

        self.botao_validar = BorderedButton(text = 'Validar', font_size = '20sp', size_hint = (0.2, 0.1), pos_hint = {'center_x': 0.42, 'y': 0.2}, color = (0,0,0,1))
        self.botao_validar.bind(on_release = self.validar_usuario)
        self.botao_validar.bind(on_release = self.salvar_configs_entrar)

        self.botao_entrar = BorderedButton(text = 'Entrar', font_size = '20sp', size_hint = (0.1, 0.1), pos_hint = {'center_x': 0.6, 'y': 0.2}, disabled = True, color = (1,1,1,1))
        self.botao_entrar.bind(on_release = self.ir_para_MainMenu)

        autenticação_layout.add_widget(imagemMatec)
        autenticação_layout.add_widget(self.label_email_autenticar)
        autenticação_layout.add_widget(self.email_autenticar)
        autenticação_layout.add_widget(self.botao_validar)
        autenticação_layout.add_widget(self.botao_entrar)

        self.add_widget(autenticação_layout)

        self.carregar_configs_entrar()

    def validar_usuario(self, *args):
        validar_usuario(self, *args)

    def obter_usuarios_validos(self, *args):
        obter_usuarios_validos(self, *args)

    def salvar_configs_entrar(self, instance):
        email_autenticar = self.email_autenticar.text

        with open('usuario_autenticado.txt', 'w') as f:
            # Escreve os dados no arquivo, separados por vírgulas
            f.write(f'{email_autenticar}')

    def carregar_configs_entrar(self):
        try:
            with open('usuario_autenticado.txt', 'r') as f:
                dados_entrar = f.read()

                # Divide a string lida nos campos usando a vírgula como separador
                dados_entrar = [item.strip() for item in dados_entrar.split(',')]

                # Atribui os valores lidos para os campos de texto
                self.email_autenticar.text = dados_entrar[0]

        except FileNotFoundError:
            pass  # Caso o arquivo não exista, não faz nada
        except ValueError:
            pass  # Caso os dados estejam no formato incorreto

class MainMenu(Screen):
    def ir_para_escolher_entrada(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'Escolher_processo_entrada'
    
    def ir_para_escolher_saída(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'Escolher_processo_saída'

    def ir_para_BotoesDifal(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'BotoesDifal'

    def ir_para_config(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'BotoesConfig'
        
    def __init__(self, **kwargs):
        super(MainMenu, self).__init__(**kwargs)

        main_layout = FloatLayout()

        imagemMatec = Image(source='imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})
         
        botao_recebimento_peças = BorderedButton(text = 'Processo de Recebimento', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.2, 'top': 0.9}, font_size = '20sp', color = (0, 0, 0, 1))
        botao_recebimento_peças.bind(on_release = self.ir_para_escolher_entrada)

        botao_devolução_peças = BorderedButton(text = 'Processo de devolução', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.2, 'top': 0.75}, font_size = '20sp', color = (0, 0, 0, 1))
        botao_devolução_peças.bind(on_release = self.ir_para_escolher_saída)

        botao_difal = BorderedButton(text = 'Processo Difal', size_hint=(0.3, 0.1), pos_hint = {'center_x': 0.2, 'top': 0.60}, font_size='20sp', background_down='', color = (0, 0, 0, 1))
        botao_difal.bind(on_release = self.ir_para_BotoesDifal)

        botao_ir_config = BorderedButton(text = 'Configurações de contas', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.8, 'top': 0.9}, font_size = '20sp', color = (0, 0, 0, 1))
        botao_ir_config.bind(on_release = self.ir_para_config)

        main_layout.add_widget(imagemMatec)
        main_layout.add_widget(botao_recebimento_peças)
        main_layout.add_widget(botao_devolução_peças)
        main_layout.add_widget(botao_difal)
        main_layout.add_widget(botao_ir_config)

        self.add_widget(main_layout)

class Escolher_processo_entrada(Screen):
    def voltar_MainMenu(self, instance):
        self.manager.transition = SlideTransition(direction = 'right')
        self.manager.current = 'MainMenu'
    
    def ir_entrada_dell(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'Botoes_entrada_Dell'

    def ir_entrada_hp(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'Botoes_entrada_HP'

    def __init__(self, **kwargs):
        super(Escolher_processo_entrada, self).__init__(**kwargs)

        layout_escolher_processos = FloatLayout()

        imagemMatec = Image(source='imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        botao_processo_dell = BorderedButton(text = 'Recebimento Dell', font_size = '25sp', size_hint = (0.3, 0.08), pos_hint = {'center_x': 0.5, 'y': 0.70}, color = (0, 0, 0, 1))
        botao_processo_dell.bind(on_release = self.ir_entrada_dell)

        botao_processo_hp = BorderedButton(text = 'Recebimento HP', font_size = '25sp', size_hint = (0.3, 0.08), pos_hint = {'center_x': 0.5, 'y': 0.60}, color = (0, 0, 0, 1))
        botao_processo_hp.bind(on_release = self.ir_entrada_hp)

        botao_MainMenu = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.17, 'y': 0.93})
        botao_MainMenu.bind(on_release = self.voltar_MainMenu)

        layout_escolher_processos.add_widget(imagemMatec)
        layout_escolher_processos.add_widget(botao_processo_dell)
        layout_escolher_processos.add_widget(botao_processo_hp)
        layout_escolher_processos.add_widget(botao_MainMenu)

        self.add_widget(layout_escolher_processos)

class Botoes_entrada_Dell(Screen):

    def voltar_escolher_processo(self, instance):
            self.manager.transition = SlideTransition(direction = 'right')
            self.manager.current = 'Escolher_processo_entrada'

    def __init__(self, **kwargs):
        super(Botoes_entrada_Dell, self).__init__(**kwargs)

        layout_entrada_dell = FloatLayout()

        imagemMatec = Image(source = 'imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        botao_baixar_xml_dell = BorderedButton(text = 'Baixar XMLs Dell', font_size = '25sp', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.17, 'y': 0.78}, color = (0, 0, 0, 1))
        botao_baixar_xml_dell.bind(on_release = self.baixar_arquivosXML_DELL)

        botao_biparxml_dell = BorderedButton(text = 'Bipar NFs Dell', font_size = '25sp', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.50, 'y': 0.78}, color = (0, 0, 0, 1))
        botao_biparxml_dell.bind(on_release = self.biparxml)

        botao_entradaNF_dell = BorderedButton(text = 'Criar planilha de entrada\nDell', font_size = '25sp', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.83, 'y': 0.78}, color = (0, 0, 0, 1))
        botao_entradaNF_dell.bind(on_release = self.criar_planilha_entrada_nf_DELL)

        Importar_NF = BorderedButton(text = 'Importar NF', font_size = '25sp', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.17, 'y': 0.66}, color = (0, 0, 0, 1))
        Importar_NF.bind(on_release = self.importar_produtos)

        Emitir_nf_circulação = BorderedButton(text = 'Emitir NF circulação', font_size = '25sp', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.50, 'y': 0.66}, color = (0, 0, 0, 1))
        Emitir_nf_circulação.bind(on_release = self.emitir_nf_circulação)

        Emitir_nf_entrada_tecnico = BorderedButton(text = 'Emitir NF recebimento\nTec.', font_size = '25sp', size_hint = (0.3, 0.1), pos_hint = {'center_x': 0.83, 'y': 0.66}, color = (0, 0, 0, 1))
        Emitir_nf_entrada_tecnico.bind(on_release = self.emitir_nf_entrada_tec)

        self.log_input = TextInput(text = '**Logs e informações**\n\n', size_hint = (0.96, 0.6), pos_hint = {'center_x': 0.50, 'y': 0.03}, multiline = True, readonly = True, font_size = '25sp')
        
        botao_escolher_processo = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.17, 'y': 0.93})
        botao_escolher_processo.bind(on_release = self.voltar_escolher_processo)

        botao_escolher_planilha = BorderedButton_top(text = 'Escolher outra planilha', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.83, 'y': 0.93})
        botao_escolher_planilha.bind(on_release = self.escolher_planilha)
        
        layout_entrada_dell.add_widget(imagemMatec)
        layout_entrada_dell.add_widget(botao_escolher_processo)
        layout_entrada_dell.add_widget(botao_escolher_planilha)
        layout_entrada_dell.add_widget(botao_baixar_xml_dell)
        layout_entrada_dell.add_widget(botao_biparxml_dell)
        layout_entrada_dell.add_widget(botao_entradaNF_dell)
        layout_entrada_dell.add_widget(Importar_NF)
        layout_entrada_dell.add_widget(Emitir_nf_circulação)
        layout_entrada_dell.add_widget(Emitir_nf_entrada_tecnico)
        layout_entrada_dell.add_widget(self.log_input)

        self.add_widget(layout_entrada_dell)

    def baixar_arquivosXML_DELL(self, *args):
        baixar_arquivosXML_DELL(self, *args) 

    def criar_planilha_entrada_nf_DELL(self, *args):
        criar_planilha_entrada_nf_DELL(self, *args) 
    
    def importar_produtos(self, *args):
        importar_produtos(self, *args)

    def biparxml(self, *args):
        biparxml(self, *args)
    
    def emitir_nf_circulação(self, *args):
        emitir_nf_circulação(self, *args)

    def emitir_nf_entrada_tec(self, *args):
        emitir_nf_entrada_tec(self, *args) # type: ignore

    def escolher_planilha(self, *args):
        """Escolher uma nova planilha e atualizar planilha_df."""
        global planilha_df  # Usar a variável global para permitir a atualização
        nova_planilha_df = SolicitarPlanilha.escolher_planilha(self) 
        if nova_planilha_df is not None:
            planilha_df = nova_planilha_df  # Atualiza a variável global
            log_message(self.log_input, 'Planilha substituída com sucesso.\n')
        else:
            log_message(self.log_input, 'Não foi possível substituir a planilha.\n')

class Botoes_entrada_HP(Screen):
    def voltar_escolher_processo(self, instance):
        self.manager.transition = SlideTransition(direction = 'right')
        self.manager.current = 'Escolher_processo_entrada'

    def __init__(self, **kwargs):
        super(Botoes_entrada_HP, self).__init__(**kwargs)

        layout_entrada_HP = FloatLayout()

        imagemMatec = Image(source='imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        botao_baixar_xml_hp = BorderedButton(text = 'Baixar arquivos XMLs HP', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.17, 'y': 0.65}, color = (0, 0, 0, 1))
        botao_baixar_xml_hp.bind(on_release = self.baixar_arquivosXML_HP)

        botao_biparxml_hp = BorderedButton(text = 'Bipar notas fiscais HP', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.50, 'y': 0.65}, color = (0, 0, 0, 1))
        botao_biparxml_hp.bind(on_release = self.biparxml)

        botao_entradaNF_hp = BorderedButton(text = 'Criar planilha NF entrada\nHP', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.83, 'y': 0.65}, color = (0, 0, 0, 1))
        botao_entradaNF_hp.bind(on_release = self.criar_planilha_entrada_nf_HP)

        self.log_input = TextInput(text = '**Logs e informações**\n\n', size_hint = (0.96, 0.6), pos_hint = {'center_x': 0.50, 'y': 0.03}, multiline = True, readonly = True, font_size = '25sp')
        
        botao_escolher_processo = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.17, 'y': 0.93})
        botao_escolher_processo.bind(on_release = self.voltar_escolher_processo)

        botao_escolher_planilha = BorderedButton_top(text = 'Escolher outra planilha', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.83, 'y': 0.93})
        botao_escolher_planilha.bind(on_release = self.escolher_planilha)
        
        layout_entrada_HP.add_widget(imagemMatec)
        layout_entrada_HP.add_widget(botao_escolher_processo)
        layout_entrada_HP.add_widget(botao_escolher_planilha)
        layout_entrada_HP.add_widget(botao_baixar_xml_hp)
        layout_entrada_HP.add_widget(botao_biparxml_hp)
        layout_entrada_HP.add_widget(botao_entradaNF_hp)
        layout_entrada_HP.add_widget(self.log_input)

        self.add_widget(layout_entrada_HP)

    def baixar_arquivosXML_HP(self, *args):
        baixar_arquivosXML_HP(self, *args)

    def criar_planilha_entrada_nf_HP(self, *args):
        criar_planilha_entrada_nf_HP(self, *args)

    def biparxml(self, *args):
        biparxml(self, *args)

    def escolher_planilha(self, *args):
        """Escolher uma nova planilha e atualizar planilha_df."""
        global planilha_df  # Usar a variável global para permitir a atualização
        nova_planilha_df = SolicitarPlanilha.escolher_planilha(self) 
        if nova_planilha_df is not None:
            planilha_df = nova_planilha_df  # Atualiza a variável global
            log_message(self.log_input, 'Planilha substituída com sucesso.\n')
        else:
            log_message(self.log_input, 'Não foi possível substituir a planilha.\n')

class escolher_processo_saída(Screen):
    def voltar_MainMenu(self, instance):
        self.manager.transition = SlideTransition(direction = 'right')
        self.manager.current = 'MainMenu'
    
    def ir_saída_dell(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'Botoes_saída_Dell'

    def ir_saída_hp(self, instance):
        self.manager.transition = SlideTransition(direction = 'left')
        self.manager.current = 'Botoes_saída_HP'

    def __init__(self, **kwargs):
        super(escolher_processo_saída, self).__init__(**kwargs)

        layout_escolher_processos = FloatLayout()

        imagemMatec = Image(source='imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        botao_processo_dell = BorderedButton(text = 'Devolução Dell', font_size = '25sp', size_hint = (0.3, 0.08), pos_hint = {'center_x': 0.5, 'y': 0.70}, color = (0, 0, 0, 1))
        botao_processo_dell.bind(on_release = self.ir_saída_dell)

        botao_processo_hp = BorderedButton(text = 'Devolução HP', font_size = '25sp', size_hint = (0.3, 0.08), pos_hint = {'center_x': 0.5, 'y': 0.60}, color = (0, 0, 0, 1))
        botao_processo_hp.bind(on_release = self.ir_saída_hp)

        botao_MainMenu = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.17, 'y': 0.93})
        botao_MainMenu.bind(on_release = self.voltar_MainMenu)

        layout_escolher_processos.add_widget(imagemMatec)
        layout_escolher_processos.add_widget(botao_processo_dell)
        layout_escolher_processos.add_widget(botao_processo_hp)
        layout_escolher_processos.add_widget(botao_MainMenu)

        self.add_widget(layout_escolher_processos)

class Botoes_devolução_Dell(Screen):
    def voltar_escolher_processo(self, instance):
            self.manager.transition = SlideTransition(direction = 'right')
            self.manager.current = 'Escolher_processo_saída'

    def __init__(self, **kwargs):
        super(Botoes_devolução_Dell, self).__init__(**kwargs)

        imagemMatec = Image(source = 'imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        layout_devolução_Dell = FloatLayout()

        botao_valores_devolução_Dell = BorderedButton(text = 'Valores devolução Dell', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.17, 'y': 0.65}, color = (0, 0, 0, 1)) 
        botao_valores_devolução_Dell.bind(on_release = self.valores_devolução_DELL)

        botao_emitir_nf_Dell = BorderedButton(text = 'Emitir notas fiscais Dell', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.5, 'y': 0.65}, color = (0, 0, 0, 1))
        botao_emitir_nf_Dell.bind(on_release = self.emitir_NF_dev_dell)

        self.log_input = TextInput(text = '**Logs e informações**\n\n', size_hint = (0.96, 0.6), pos_hint = {'center_x': 0.50, 'y': 0.03}, multiline = True, readonly = True, font_size = '25sp')

        botao_voltar_escolher_processo_dev = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.17, 'y': 0.93})
        botao_voltar_escolher_processo_dev.bind(on_release = self.voltar_escolher_processo)

        botao_escolher_planilha_Dell = BorderedButton_top(text = 'Escolher outra planilha', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.83, 'y': 0.93})
        botao_escolher_planilha_Dell.bind(on_release = self.escolher_planilha)

        layout_devolução_Dell.add_widget(imagemMatec)
        layout_devolução_Dell.add_widget(botao_voltar_escolher_processo_dev)
        layout_devolução_Dell.add_widget(botao_escolher_planilha_Dell)
        layout_devolução_Dell.add_widget(botao_valores_devolução_Dell)
        layout_devolução_Dell.add_widget(botao_emitir_nf_Dell)
        layout_devolução_Dell.add_widget(self.log_input)

        self.add_widget(layout_devolução_Dell)

    def valores_devolução_DELL(self, *args):
        valores_devolução_DELL(self, *args) 
    
    def emitir_NF_dev_dell(self, *args):
        emitir_NF_dev_dell(self, *args)

    def escolher_planilha(self, *args):
        """Escolher uma nova planilha e atualizar planilha_df."""
        global planilha_df  # Usar a variável global para permitir a atualização
        nova_planilha_df = SolicitarPlanilha.escolher_planilha(self) 
        if nova_planilha_df is not None:
            planilha_df = nova_planilha_df  # Atualiza a variável global
            log_message(self.log_input, 'Planilha substituída com sucesso.\n')
        else:
            log_message(self.log_input, 'Não foi possível substituir a planilha.\n')

class Botoes_devolução_HP(Screen):
    def voltar_escolher_processo(self, instance):
        self.manager.transition = SlideTransition(direction = 'right')
        self.manager.current = 'Escolher_processo_saída'

    def __init__(self, **kwargs):
        super(Botoes_devolução_HP, self).__init__(**kwargs)

        imagemMatec = Image(source='imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        layout_devolução_HP = FloatLayout()

        botao_valores_devolução_HP = BorderedButton(text = 'Valores devolução HP', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.17, 'y': 0.65}, color = (0, 0, 0, 1)) 
        botao_valores_devolução_HP.bind(on_release = self.valores_devolução_HP)

        botao_emitir_nf_HP = BorderedButton(text = 'Emitir notas fiscais HP', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.5, 'y': 0.65}, color = (0, 0, 0, 1))
        #botao_emitir_nf_HP.bind(on_relase = self.emitir_nf_HP)

        self.log_input = TextInput(text = '**Logs e informações**\n\n', size_hint = (0.96, 0.6), pos_hint = {'center_x': 0.50, 'y': 0.03}, multiline = True, readonly = True, font_size = '25sp')

        botao_voltar_escolher_processo_dev = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.17, 'y': 0.93})
        botao_voltar_escolher_processo_dev.bind(on_release = self.voltar_escolher_processo)

        botao_escolher_planilha_HP = BorderedButton_top(text = 'Escolher outra planilha', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.83, 'y': 0.93})
        botao_escolher_planilha_HP.bind(on_release = self.escolher_planilha)

        layout_devolução_HP.add_widget(imagemMatec)
        layout_devolução_HP.add_widget(botao_voltar_escolher_processo_dev)
        layout_devolução_HP.add_widget(botao_escolher_planilha_HP)
        layout_devolução_HP.add_widget(botao_valores_devolução_HP)
        layout_devolução_HP.add_widget(botao_emitir_nf_HP)
        layout_devolução_HP.add_widget(self.log_input)

        self.add_widget(layout_devolução_HP)


    def valores_devolução_HP(self, *args):
        valores_devolução_HP(self, *args)
    
    def emitir_nf_HP(self, *args):
        emitir_nf_HP(self, *args) # type: ignore

    def escolher_planilha(self, *args):
        """Escolher uma nova planilha e atualizar planilha_df."""
        global planilha_df  # Usar a variável global para permitir a atualização
        nova_planilha_df = SolicitarPlanilha.escolher_planilha(self) 
        if nova_planilha_df is not None:
            planilha_df = nova_planilha_df  # Atualiza a variável global
            log_message(self.log_input, 'Planilha substituída com sucesso.\n')
        else:
            log_message(self.log_input, 'Não foi possível substituir a planilha.\n')

class Botoes_difal(Screen):
    def ir_para_MainMenu(self, instance):
            self.manager.transition = SlideTransition(direction = 'right')
            self.manager.current = 'MainMenu'

    def __init__(self, **kwargs):
        super(Botoes_difal, self).__init__(**kwargs)

        layout_difal = FloatLayout()

        imagemMatec = Image(source='imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        botao_difal_dell = BorderedButton(text = 'Criar planilha Difal', font_size = '25sp', size_hint = (0.3, 0.2), pos_hint = {'center_x': 0.5, 'y': 0.65}, color = (0, 0, 0, 1))
        botao_difal_dell.bind(on_release = self.criar_planilha_difal)

        self.log_input = TextInput(text = '**Logs e informações**\n\n', size_hint = (0.96, 0.6), pos_hint = {'center_x': 0.50, 'y': 0.03}, multiline = True, readonly = True, font_size = '25sp')

        botao_MainMenu = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.17, 'y': 0.93})
        botao_MainMenu.bind(on_release = self.ir_para_MainMenu)

        botao_escolher_planilha = BorderedButton_top(text = 'Escolher outra planilha', font_size = '20sp', size_hint = (0.3, 0.05), pos_hint = {'center_x': 0.83, 'y': 0.93})
        botao_escolher_planilha.bind(on_release = self.escolher_planilha)


        layout_difal.add_widget(imagemMatec)
        layout_difal.add_widget(botao_MainMenu)
        layout_difal.add_widget(botao_escolher_planilha)
        layout_difal.add_widget(botao_difal_dell)
        layout_difal.add_widget(self.log_input)

        self.add_widget(layout_difal)

    def escolher_planilha(self, *args):
        """Escolher uma nova planilha e atualizar planilha_df."""
        global planilha_df  # Usar a variável global para permitir a atualização
        nova_planilha_df = SolicitarPlanilha.escolher_planilha(self) 
        if nova_planilha_df is not None:
            planilha_df = nova_planilha_df  # Atualiza a variável global
            log_message(self.log_input, 'Planilha substituída com sucesso.\n')
        else:
            log_message(self.log_input, 'Não foi possível substituir a planilha.\n')
    
    def criar_planilha_difal(self, *args):
        criar_planilha_difal(self, *args)

class configurar_contas(Screen):
    def ir_para_MainMenu(self, instance):
        self.manager.transition = SlideTransition(direction = 'right')
        self.manager.current = 'MainMenu'

    def __init__(self, **kwargs):
        super(configurar_contas, self).__init__(**kwargs)
        
        scroll_view = ScrollView(size_hint = (1, 1), do_scroll_x = False, do_scroll_y = True)
        
        # Layout para conter os widgets
        layout_contas_config = FloatLayout(size_hint_y=None)
        layout_contas_config.height = 1200  # Altura total para permitir rolagem

        imagemMatec = Image(source='imagemfundo.png', allow_stretch = True, keep_ratio = False, size_hint = (1, 1), pos_hint = {'x': 0, 'y': 0})

        botao_MainMenu = BorderedButton_top(text = 'Voltar', font_size = '20sp', size_hint = (0.3, 0.03), pos_hint = {'center_x': 0.17, 'y': 0.95})
        botao_MainMenu.bind(on_release = self.ir_para_MainMenu)

        label_email = Label(text = 'Email Matec:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.16, 'y': 0.88})
        self.email_input = TextInput(font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.25, 'y': 0.86}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_usuario_Elogistica = Label(text = 'Usuário E-logistica:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.185, 'y': 0.78})
        self.usuario_Elogistica = TextInput(font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.25, 'y': 0.76}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_senha_Elogistica = Label(text = 'Senha E-logistica:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.18, 'y': 0.68})
        self.senha_Elogistica = TextInput(font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.25, 'y': 0.66}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_usuario_IOB = Label(text = 'Usuário IOB:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.156, 'y': 0.58})
        self.usuario_IOB = TextInput(font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.25, 'y': 0.56}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_senha_IOB = Label(text = 'Senha IOB:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.15, 'y': 0.48})
        self.senha_IOB = TextInput(font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.25, 'y': 0.46}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_aliquota_interna = Label(text = 'Alíquota interna do estado', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.719, 'y': 0.88})
        self.aliquota_interna = TextInput(text = 'Planilha Difal', font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.75, 'y': 0.86}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_nome_credenciada = Label(text = 'Nome credenciada:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.687, 'y': 0.78})
        self.nome_credenciada = TextInput(text = 'Planilha Difal', font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.75, 'y': 0.76}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_cnpj_credenciada = Label(text = 'CNPJ credenciada:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.684, 'y': 0.68})
        self.cnpj_credenciada = TextInput(text = 'Planilha Difal', font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.75, 'y': 0.66}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        label_caixa_emails = Label(text = 'Nome da pasta de arquivos:', color = (1,1,1,1) ,font_size = '20sp', size_hint = (0.3, 0.07), pos_hint = {'center_x': 0.72, 'y': 0.58})
        self.caixa_emails = TextInput(text = 'Email', font_size = '20sp', size_hint = (0.3, 0.04), pos_hint = {'center_x': 0.75, 'y': 0.56}, multiline = False, write_tab = False, halign = 'center', padding_y = (10, 0))

        botao_salvar = BorderedButton_top(text = 'Salvar', font_size = '20sp', size_hint = (0.25, 0.05), pos_hint = {'center_x': 0.5, 'y': 0.35})
        botao_salvar.bind(on_release = self.salvar_configs)

        layout_contas_config.add_widget(imagemMatec)
        layout_contas_config.add_widget(botao_MainMenu)
        layout_contas_config.add_widget(label_email)
        layout_contas_config.add_widget(self.email_input)
        layout_contas_config.add_widget(label_usuario_Elogistica)
        layout_contas_config.add_widget(self.usuario_Elogistica)
        layout_contas_config.add_widget(label_senha_Elogistica)
        layout_contas_config.add_widget(self.senha_Elogistica)
        layout_contas_config.add_widget(label_usuario_IOB)
        layout_contas_config.add_widget(self.usuario_IOB)
        layout_contas_config.add_widget(label_senha_IOB)
        layout_contas_config.add_widget(self.senha_IOB)
        layout_contas_config.add_widget(label_aliquota_interna)
        layout_contas_config.add_widget(self.aliquota_interna)
        layout_contas_config.add_widget(label_nome_credenciada)
        layout_contas_config.add_widget(self.nome_credenciada)
        layout_contas_config.add_widget(label_cnpj_credenciada)
        layout_contas_config.add_widget(self.cnpj_credenciada)
        layout_contas_config.add_widget(label_caixa_emails)
        layout_contas_config.add_widget(self.caixa_emails)

        layout_contas_config.add_widget(botao_salvar)

        scroll_view.add_widget(layout_contas_config)

        self.add_widget(scroll_view)

        self.carregar_configs()

    def salvar_configs(self, instance):
        email = self.email_input.text
        usuario_Elogistica = self.usuario_Elogistica.text
        senha_Elogistica = self.senha_Elogistica.text
        usuario_IOB = self.usuario_IOB.text
        senha_IOB = self.senha_IOB.text
        aliquota_interna = self.aliquota_interna.text
        nome_credenciada = self.nome_credenciada.text
        cnpj_credenciada = self.cnpj_credenciada.text
        caixa_emails = self.caixa_emails.text

        with open('email.txt', 'w') as f:
            # Escreve os dados no arquivo, separados por vírgulas
            f.write(f'{email}, {usuario_Elogistica}, {senha_Elogistica}, {usuario_IOB}, {senha_IOB}, {aliquota_interna}, {nome_credenciada}, {cnpj_credenciada}, {caixa_emails}')

    def carregar_configs(self):
        try:
            with open('email.txt', 'r') as f:
                dados = f.read()

                # Divide a string lida nos campos usando a vírgula como separador
                dados = [item.strip() for item in dados.split(',')]

                # Atribui os valores lidos para os campos de texto
                self.email_input.text = dados[0]
                self.usuario_Elogistica.text = dados[1]
                self.senha_Elogistica.text = dados[2]
                self.usuario_IOB.text = dados[3]
                self.senha_IOB.text = dados[4]
                self.aliquota_interna.text = dados[5]
                self.nome_credenciada.text = dados[6]
                self.cnpj_credenciada.text = dados[7]
                self.caixa_emails.text = dados[8]

        except FileNotFoundError:
            pass  # Caso o arquivo não exista, não faz nada
        except ValueError:
            pass  # Caso os dados estejam no formato incorreto