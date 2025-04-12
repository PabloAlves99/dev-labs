from datetime import datetime
from tkinter.filedialog import askdirectory
import os


class OrganizadorDeArquivos:
    def __init__(self):
        self.data_atual = datetime.now().strftime("%d%m%Y")
        self.segundo_atual = datetime.now().strftime("%S")
        self.minuto_atual = datetime.now().strftime("%M")
        self.arquivos_com_erro_permissao = []

    def organizar_arquivos(self):
        """
        Organizador de arquivos por extensão.
        O script irá organizar os arquivos de uma pasta em pastas separadas
        de acordo com a extensão do arquivo.

        Exemplo:
        Se a pasta tiver os arquivos: arquivo1.txt, arquivo2.txt, arquivo3.jpg
        O script irá criar as pastas: TXT e JPG e mover os arquivos para as
        respectivas pastas.

        Uso:
        - Selecione a pasta que deseja organizar.
        - O script irá criar uma pasta chamada "organizadorDeArquivos" dentro
        da pasta selecionada.
        - Dentro da pasta "organizadorDeArquivos" serão criadas pastas de
        acordo com a extensão dos arquivos.
        - Os arquivos serão movidos para as pastas de acordo com a extensão.
        - Caso haja algum arquivo com erro de permissão, o script irá informar
        quais arquivos não puderam ser movidos.
        """
        caminho = askdirectory(title="Selecione uma pasta")
        if not caminho:
            return
        lista_arquivos = os.listdir(caminho)

        caminho_organizador = f"{caminho}/organizadorDeArquivos"
        if not os.path.exists(caminho_organizador):
            os.makedirs(caminho_organizador)

        for arquivo in lista_arquivos:
            _, extensao = os.path.splitext(f"{caminho}/{arquivo}")

            if extensao:
                try:
                    extensao_formatada = extensao.upper().replace(".", "")
                    pasta_extensao = f"{caminho_organizador}/{extensao_formatada}"

                    if not os.path.exists(pasta_extensao):
                        os.makedirs(pasta_extensao)

                    os.rename(f"{caminho}/{arquivo}",
                              f"{pasta_extensao}/{arquivo}")

                except PermissionError:
                    self.arquivos_com_erro_permissao.append(arquivo)

                except FileExistsError:
                    novo_nome = f"{self.minuto_atual}{self.segundo_atual}_{arquivo}"
                    os.rename(f"{caminho}/{arquivo}",
                              f"{pasta_extensao}/{novo_nome}")

        if self.arquivos_com_erro_permissao:
            arquivos = '\n'.join(self.arquivos_com_erro_permissao)
            print(f"Os seguintes arquivos não puderam ser movidos devido a "
                  f" problemas de permissão:\nCaso o arquivo esteja aberto em "
                  f"algum lugar, feche antes de executar o script \n\n"
                  f"{arquivos}.")

        print("Arquivos organizados com sucesso!")


if __name__ == '__main__':
    OrganizadorDeArquivos().organizar_arquivos()
