import os
from datetime import date


def get_last_file():
    # Obtendo o nome do arquivo mais recente da pasta downloads
    path = r'C:\Users\jorge.nascimento\Downloads'
    lista_arquivos = os.listdir(path)
    lista_datas = []
    for arquivo in lista_arquivos:
        data = os.path.getmtime(f'{path}/{arquivo}')
        lista_datas.append((data,arquivo))

    lista_datas.sort(reverse=True)
    return f'{path}/{lista_datas[0][1]}'

