# pylint: disable=missing-docstring,empty-docstring
import locale
from string import Template
from pathlib import Path
from datetime import datetime
from pytz import timezone


class EmailBody:
    def __init__(self) -> None:
        self.get_email_path()
        self.prepare_email_content()

    def get_email_path(self):
        self.caminho_msg = Path(__file__).parent / 'igpmensagem_email.html'

    def prepare_email_content(self):
        data = datetime.now(timezone('America/Sao_paulo')).strftime('%d/%m/%Y')
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        compra = {
            'nome': 'Pablo Jr',
            'valor': locale.currency(935.99, grouping=True),
            'data': data,
            'servico': 'Automação',
            'numero': '+55 (31) 99423-4449'
        }

        try:
            with open(self.caminho_msg, 'r', encoding='utf-8') as email:
                txt = email.read()
                template = Template(txt)
                self.text_email = template.substitute(compra)
        except FileNotFoundError:
            print(f"Erro: Arquivo {self.caminho_msg} não encontrado.")
            return
        except Exception as e:
            print(f"Erro ao ler o arquivo de mensagem: {e}")
            return
