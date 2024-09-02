import shutil
import os

diretorio = "C:/Users/danie/Documents/teste2"
pastas = os.listdir(diretorio)

# Função para tentar criar as pastas
def criar_pasta(caminho):
    arquivos_caminho = os.listdir(caminho)
    print(arquivos_caminho)
    pastas_extensoes_padrao = {
        ".docx" : "WORD",
        ".doc" : "WORD",
        ".pptx" : "POWER POINT",
        ".ppt" : "POWER POINT",
        ".xlsx" : "EXCEL",
        ".xls" : "EXCEL"
    }

    for extensao in pastas_extensoes_padrao:

        nome_pasta = pastas_extensoes_padrao[extensao]
        for arquivo in arquivos_caminho:
            if arquivo.endswith(extensao):
                try:
                    os.mkdir(f"{caminho}/{nome_pasta}")
                except FileExistsError:
                    for arquivo in arquivos_caminho:
                        print(arquivo)
                        if arquivo.upper() == nome_pasta:
                            os.renames(f"{caminho}/{arquivo}", f"{caminho}/{nome_pasta}")
                            break
                        else:
                            pass
            elif arquivo.upper() == nome_pasta:
                os.renames(f"{caminho}/{arquivo}", f"{caminho}/{nome_pasta}")


for pasta in pastas:
    os.chdir(f"{diretorio}/{pasta}")
    caminho_pasta_aluno_atual = os.getcwd()
    arquivos_caminho_atual = os.listdir(caminho_pasta_aluno_atual)

    # Tentando criar as pastas:
    criar_pasta(caminho_pasta_aluno_atual)

    for arquivo in arquivos_caminho_atual:

        if arquivo.endswith(".docx") or arquivo.endswith(".doc"):
            shutil.move(f"{caminho_pasta_aluno_atual}/{arquivo}", f"{caminho_pasta_aluno_atual}/WORD/{arquivo}")

        elif arquivo.endswith(".pptx") or arquivo.endswith(".ppt"):
            shutil.move(f"{caminho_pasta_aluno_atual}/{arquivo}", f"{caminho_pasta_aluno_atual}/POWER POINT/{arquivo}")

        elif arquivo.endswith(".xlsx") or arquivo.endswith(".xls"):
            shutil.move(f"{caminho_pasta_aluno_atual}/{arquivo}", f"{caminho_pasta_aluno_atual}/EXCEL/{arquivo}")

