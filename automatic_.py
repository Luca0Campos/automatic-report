from docx import Document
import pandas as pd
import openpyxl
import numpy
import win32com.client as win32
from docx2pdf import convert
import easygui
import os
import shutil
import pyodbc

try:
    tabela = pd.read_excel("")
except:
    easygui.msgbox("O nome ou caminho da pasta pode estar errado", title="Erro na Tabela")

#variáveis das pastas p/ arquivos
enviar_arquivos = ("")
arquivos_word = ("")
arquivos_pdf = ("")
arquivos_enviados = ("")

# se não existir, criar as respectivas pastas
if (not os.path.exists(arquivos_word)):
    os.mkdir(arquivos_word)
if (not os.path.exists(arquivos_pdf)):
    os.mkdir(arquivos_pdf)
if (not os.path.exists(arquivos_enviados)):
    os.mkdir(arquivos_enviados)

# se o txt estiver na pasta de enviar, não criar um novo
if enviar_arquivos in arquivos_enviados:
    pass
else:
    with open(f"{arquivos_enviados}Arquivos Enviados.txt", "a") as nome_de_envio:

        for linha in tabela.index:
            try:
                documento = Document("")
            except:
                easygui.msgbox("O nome ou caminho da pasta pode estar errado", title="Erro no Doc Word")

            # setando as variáveis e setando a linha na tabela.
            try:
                razao_social = tabela.loc[linha, "RazaoSocial"]
                cpf = str(tabela.loc[linha, "CNPJ"])
                arquivo = (tabela.loc[linha, "Arquivo"])
                nome_email = str(tabela.loc[linha, "Email"])
            except KeyError as erro:
                easygui.msgbox("O nome da coluna pode estar errado", title="Erro na Tabela")

            # variável para saber se o arquivo já foi baixado.
            documento_salvo = f"{arquivos_word}{arquivo}.docx"

            # condicional para adicionar ou não zeros a esquerda.
            if len(cpf) == 13:
                cpf = str("0" + cpf)
            elif len(cpf) == 12:
                cpf = str("00" + cpf)
            elif len(cpf) == 14:
                cpf = str(cpf)
            elif len(cpf) == 11:
                cpf = str(cpf)
            elif len(cpf) < 10:
                print(f'não consta nenhum cpf/cnpj para: {razao_social}')
                pass

            # condicional para ajustar a forma correta
            if len(cpf) == 14:
                cpf = cpf.zfill(14)
                cpf = '{}.{}.{}/{}-{}'.format(cpf[:2], cpf[2:5], cpf[5:8], cpf[8:12], cpf[12:14])
            elif len(cpf) == 12:
                cpf = cpf.zfill(12)
                cpf = '{}.{}.{}/{}-{}'.format(cpf[:2], cpf[2:5], cpf[5:8], cpf[8:12], cpf[12:14])  # formato cnpj
            elif len(cpf) == 13:
                cpf = cpf.zfill(13)
                cpf = '{}.{}.{}/{}-{}'.format(cpf[:2], cpf[2:5], cpf[5:8], cpf[8:12], cpf[12:14])

            # entre "" o nome onde irá realizar a alteração no doc word
            referencias = {
                "RAZÃO SOCIAL ASSOCIADO": razao_social,
                "CNPJ/CPFCLIENTE": cpf,
            }

            for paragrafo in documento.paragraphs:
                for codigo in referencias:
                    paragrafo.text = paragrafo.text.replace(codigo, referencias[codigo])

            # se o word não estiver na pasta, pode criá-lo
            if (not os.path.exists(documento_salvo)):
                documento.save(f"C:/SINCOMERCIO/Arquivos Word/{arquivo}.docx")

                # conversor para pdf
                new_documento = convert(f"C:/SINCOMERCIO/Arquivos Word/{arquivo}.docx", f"C:/SINCOMERCIO/Arquivos PDF/{arquivo}.pdf")
                arquivo_pdf = (f"C:/SINCOMERCIO/Arquivos PDF/{arquivo}.pdf")

                # abre o outlook
                outlook = win32.Dispatch("outlook.application")

                # cria um email
                email = outlook.CreateItem(0)

                # enviar para:

                email.To = nome_email
                email.Subject = "[LGPD] Notificação de Proteção de Dados"
                email.HTMLBody = """
                <p>Prezados,</p>
                <p>Conforme é de amplo conhecimento, desde 18 de setembro de 2020 está em vigor a Lei nº 13.709 de 2018, mais conhecida como Lei Geral de Proteção de Dados Pessoais (LGPD).</br>
                <br>Desta forma, estamos tomando todas as providências necessárias para garantir o seu cumprimento, conforme documento anexo, cuja finalidade é demonstrar os cuidados que estamos tomando, visando sempre a segurança dos nossos clientes.</br>

                <p>Grato,</p>

                <p>Jurídico Carriers</p>
                 """

                # setando o anexo criado anteriormente
                anexo = (f"{arquivo_pdf}")

                # adicionando o anexo ao envio do email
                email.Attachments.Add(anexo)

                if len(nome_email) == 3:
                    print(f" Não consta nenhum email de {razao_social} ")
                    pass
                else:
                    arquivos_pdf_enviados = shutil.copy2(f"{arquivo_pdf}", f"{arquivos_enviados}{arquivo}.pdf")
                    email.Display()
                    print(f"email enviado para {razao_social}")
                    nome_escrito = nome_de_envio.write(f"\nO email foi enviado para: {razao_social}")

                    # integrando o banco de dados
                    dados_conexao = (
                        "Driver={SQL Server};"
                        "Server=DESKTOP-NMVBR54;"
                        "Database=PythonSQL"
                    )

                    conexão = pyodbc.connect(dados_conexao)

                    cursor =  conexão.cursor()
                    # tabela mostrando as informações selecionadas
                    comando = f"""INSERT INTO ArquivosEnviadosSINCOMERCIO(razao_social, cnpj, arquivo_enviado)
                    VALUES
                        ('{razao_social}', '{cpf}', 'Yes')"""

                    cursor.execute(comando)
                    cursor.commit()