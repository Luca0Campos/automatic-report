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
from tkinter import*

nomeArquivo = "/Contratos/Corpo do Email.txt"

def colDados():
    if (not os.path.exists(nomeArquivo)):
        with open(nomeArquivo,"w") as corpo_email_txt:
            corpo_email_txt.write(corpo_email.get("1.0",END))

if (not os.path.exists(nomeArquivo)):
    app =Tk()
    app.title("Corpo do Email")
    app.geometry('850x550')
    app.configure(background="#dde")

    Label(app, text= "Corpo do Email:",background="#dde", foreground="#009",anchor=N).place(x=30,y=15,width=100,height=20)
    corpo_email = Text(app)
    corpo_email.place(x=30,y=40,width=750,height=400)

    Button(app, text="Salvar", command=colDados).place(x=32,y=450, width=100,height=20)

    app.mainloop()



try:
    tabela = pd.read_excel("PARTES.xlsx")
except:
    easygui.msgbox("O nome/caminho da pasta ou extensão do arquivo pode estar errado", title="Erro na Arquivo Excel")

# verificar como a plataforma irá reconhecer a barra
enviar_arquivos = ("/Contratos/Arquivos Enviados/Arquivos Enviados.txt")
arquivos_word = ("/Contratos/Arquivos Word/")
arquivos_pdf = ("/Contratos/Arquivos PDF/")
arquivos_enviados = ("/Contratos/Arquivos Enviados/")

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
        nome_de_envio.write(
            " Arquivo            CNPJ                       Razão Social Cliente                         CNPJ Cliente                 Email                  Razão Social")

        for linha in tabela.index:
            try:
                documento = Document("MINUTA.docx")
            except:
                easygui.msgbox("O nome/caminho da pasta ou extensão do arquivo pode estar errado",
                               title="Erro no Doc Word")

            # setando as variáveis e setando a linha na tabela.
            try:

                razao_social = str(tabela.loc[linha, "RazaoSocial/Nome"])
                cpf = str(tabela.loc[linha, "CNPJ"])
                arquivo = "DPA_" + cpf
                nome_email = str(tabela.loc[linha, "Email"])
                razao_social_cliente = str(tabela["RazaoSocialCliente"].iloc[0])
                cnpj_cliente = str(tabela["CnpjCliente"].iloc[0])
                genero = str(tabela.loc[linha, "Genero"])
                nome_pessoa = str(tabela.loc[linha, "Nome"])

            except KeyError as erro:
                easygui.msgbox("O nome da coluna pode estar errado", title="Erro na Tabela")

            # variável para saber se o arquivo já foi baixado.
            documento_salvo = f"{arquivos_word}{arquivo}.docx"
            arquivo_pdf = (f"/Contratos/Arquivos PDF/{arquivo}.pdf")

            # condicional para adicionar ou não zeros a esquerda.
            if len(cpf) == 13:
                cpf = str("0" + cpf)
            elif len(cpf) == 12:
                cpf = str("00" + cpf)
            elif len(cpf) == 14:
                cpf = str(cpf)
            elif len(cpf) == 11:
                cpf = str(cpf)
            elif len(cpf) == 10:
                cpf = str("0" + cpf)
            elif len(cpf) < 10:
                print(f'Não consta nenhum cpf/cnpj para: {razao_social}')
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
            elif len(cpf) == 11:
                cpf = cpf.zfill(11)
                cpf = '{}.{}.{}-{}'.format(cpf[:3], cpf[3:6], cpf[6:9], cpf[9:])
            elif len(cpf) == 10:
                cpf = cpf.zfill(10)
                cpf = '{}.{}.{}-{}'.format(cpf[:3], cpf[3:6], cpf[6:9], cpf[9:])

            # entre "" o nome onde irá realizar a alteração no doc word
            referencias = {
                "RAZÃO SOCIAL OU NOME": razao_social,
                "CNPJ/CPF DA PARTE": cpf,
                "RAZÃO SOCIAL CLIENTE": razao_social_cliente,
                "CNPJ CLIENTE": cnpj_cliente,
            }

            for paragrafo in documento.paragraphs:
                for codigo in referencias:
                    paragrafo.text = paragrafo.text.replace(codigo, referencias[codigo])

            # se o word não estiver na pasta, pode criá-lo
            if (not os.path.exists(documento_salvo)):
                documento.save(f"/Contratos/Arquivos Word/{arquivo}.docx")

            if (not os.path.exists(arquivo_pdf)):
                # conversor para pdf
                new_documento = convert(f"/Contratos/Arquivos Word/{arquivo}.docx",
                                        f"/Contratos/Arquivos PDF/{arquivo}.pdf")

                # abre o outlook
                outlook = win32.Dispatch("outlook.application")

                # cria um email
                email = outlook.CreateItem(0)

                with open(nomeArquivo, "r") as ler_corpo_email:
                    leitura = str(ler_corpo_email.read())

                email.Body = str(leitura)

                # enviar para:
                email.To = nome_email
                email.Subject = "[LGPD] Notificação de Proteção de Dados"

                # setando o anexo criado anteriormente
                anexo = (f"{arquivo_pdf}")

                # adicionando o anexo ao envio do email
                email.Attachments.Add(anexo)

                if len(nome_email) < 6:
                    print(f" Não consta nenhum email de {razao_social} ")
                    pass
                else:
                    arquivos_pdf_enviados = shutil.copy2(f"{arquivo_pdf}", f"{arquivos_enviados}{arquivo}.pdf")
                    email.Display()
                    print(f"Email enviado para {razao_social}")
                    nome_escrito = nome_de_envio.write(f"\n{arquivo}|{cpf}| {razao_social_cliente} | {cnpj_cliente} | {nome_email} | {razao_social}")

                    #integrando o banco de dados
                    dados_conexao = (
                        "Driver={SQL SERVER};"
                        "Server=DESKTOP-NMVBR54;"
                        "Database=PythonSQL;"
                    )

                    conexão = pyodbc.connect(dados_conexao)

                    cursor = conexão.cursor()
                    #tabela mostrando as informações selecionadas
                    comando = f"""INSERT INTO ArquivosEnviadosSINCOMERCIO(razao_social, cnpj, arquivo_enviado)
                    VALUES
                        ('{razao_social}', '{cpf}', 'Yes')"""

                    cursor.execute(comando)
                    cursor.commit()