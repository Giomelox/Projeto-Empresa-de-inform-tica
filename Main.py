from kivy.config import Config

Config.set('graphics', 'width', '1000')
Config.set('graphics', 'height', '800')

from kivy.uix.screenmanager import ScreenManager
from kivy.app import App
from Interface.screens import Tela_autenticação, MainMenu, Escolher_processo_entrada, Botoes_entrada_Dell
from Interface.screens import Botoes_entrada_HP, escolher_processo_saída, Botoes_devolução_Dell, Botoes_devolução_HP
from Interface.screens import Botoes_difal, configurar_contas


class GerenciadorDeTelas(ScreenManager):
    pass

class Matec(App):

    def build(self):

        self.icon = 'logo_matec.ico'

        gerenciador = GerenciadorDeTelas()

        gerenciador.add_widget(Tela_autenticação(name = 'Autenticar'))

        gerenciador.add_widget(MainMenu(name = 'MainMenu'))

        gerenciador.add_widget(Escolher_processo_entrada(name = 'Escolher_processo_entrada'))

        gerenciador.add_widget(Botoes_entrada_Dell(name = 'Botoes_entrada_Dell'))

        gerenciador.add_widget(Botoes_entrada_HP(name = 'Botoes_entrada_HP'))

        gerenciador.add_widget(escolher_processo_saída(name = 'Escolher_processo_saída'))

        gerenciador.add_widget(Botoes_devolução_Dell(name = 'Botoes_saída_Dell'))

        gerenciador.add_widget(Botoes_devolução_HP(name = 'Botoes_saída_HP'))

        gerenciador.add_widget(Botoes_difal(name = 'BotoesDifal'))

        gerenciador.add_widget(configurar_contas(name = 'BotoesConfig'))

        return gerenciador

# Executa o app Kivy
if __name__ == "__main__":
    Matec().run()