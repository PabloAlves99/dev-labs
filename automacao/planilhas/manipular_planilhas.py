import os
import pandas as pd
from faker import Faker
from cpf_generator import CPF


class ManipuladorPlanilhas:
    def __init__(self, nome_arquivo='dados_pessoais.xlsx'):
        self.nome_arquivo = nome_arquivo
        self.df = pd.DataFrame()
        self.carregar_planilha()

    def carregar_planilha(self):
        if os.path.exists(self.nome_arquivo):
            self.df = pd.read_excel(self.nome_arquivo)
        else:
            self.df = pd.DataFrame(columns=[
                'Nome', 'CPF', 'Idade', 'Gênero', 'Endereço',
                'Grau de Escolaridade', 'Profissão', 'Informações Relevantes'
            ])

    def adicionar_registro(self, dados, salvar=True):
        if not self.df.empty:
            df_novo = pd.DataFrame([dados], columns=self.df.columns)
            self.df = pd.concat([self.df, df_novo], ignore_index=True)
        else:
            self.df = pd.DataFrame([dados], columns=[
                'Nome', 'CPF', 'Idade', 'Gênero', 'Endereço',
                'Grau de Escolaridade', 'Profissão', 'Informações Relevantes'
            ])

        if salvar:
            self.salvar_planilha()

    def salvar_planilha(self):
        self.df.to_excel(self.nome_arquivo, index=False)
        print(f"Planilha {self.nome_arquivo} salva com sucesso!")

    def gerador_de_dados(self, quantidade=1):
        for _ in range(quantidade):
            dados = {
                'Nome': Faker().name(),
                'CPF': CPF.generate(),
                'Idade': Faker().random_int(min=18, max=80),
                'Gênero': Faker().random_element(elements=('Masculino', 'Feminino')),
                'Endereço': Faker().address(),
                'Grau de Escolaridade': Faker().random_element(elements=(
                    'Ensino Fundamental', 'Ensino Médio', 'Tecnico',
                    'Ensino Superior', 'Pós-Graduação', 'Mestrado', 'Doutorado'
                )),
                'Profissão': Faker().job(),
                'Informações Relevantes': Faker().text(max_nb_chars=200)
            }
            self.adicionar_registro(dados, salvar=False)
        self.salvar_planilha()


if __name__ == "__main__":
    manipulador = ManipuladorPlanilhas()
    manipulador.gerador_de_dados(50)
