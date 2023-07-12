import getpass
import pandas as pd
from datetime import datetime
from mailmerge import MailMerge
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def parse_date(date_str):
    return pd.to_datetime(date_str, format='%d/%m', errors='coerce')


def convert_email(value):
    if isinstance(value, str):
        return value.strip()
    return value


agora = datetime.now()
hora = agora.hour

if hora < 12:
    print("Bom dia\n")
    saudacao = "Bom dia"
elif hora < 18:
    print("Boa tarde\n")
    saudacao = "Boa Tarde"
else:
    print("Boa noite\n")
    saudacao = "Boa noite"

senha = getpass.getpass('Digite sua senha: ')

# Carrega os dados da planilha e converte a coluna "Aniversário" para o tipo datetime
dados = pd.read_excel('C:\\Users\\João Lucas\\Downloads\\Aniversáriantes Julho3.xlsx', parse_dates=['Aniversãrio'],
                      date_parser=parse_date, converters={'Emails': convert_email})

# Filtra os aniversariantes do dia
hoje = datetime.now().strftime('%d/%m')
aniversariantes_do_dia = dados[dados['Aniversãrio'].dt.strftime('%d/%m') == hoje]

if not aniversariantes_do_dia.empty:
    # Carrega o modelo do documento de cartão de aniversário
    template = 'C:\\Users\\João Lucas\\Downloads\\Cartão de Aniversário - modelo.docx'

    # Gera os cartões de aniversário e envia por e-mail
    for index, row in aniversariantes_do_dia.iterrows():
        documentos = MailMerge(template)
        name = row['Nome']
        aniversario = row['Aniversãrio']
        email = row['Emails']
        comissão = row['Descrição']
        telefone = row['Telefones']

        name = str(name)
        nome = name.title()

        documentos.merge(Nome=nome)  # Substitua NOME pelo nome da variável no modelo do documento

        # Salva cada cartão de aniversário como um documento separado
        documento_salvo = f"C:\\Users\\João Lucas\\Documents\\{nome}.docx"
        documentos.write(documento_salvo)

        # Envia o cartão de aniversário por e-mail
        msg = MIMEMultipart()
        msg['Subject'] = 'Cartão de Aniversário'
        msg['From'] = 'jaojao04999@outlook.com'
        if pd.isna(email):
            msg['To'] = 'jaojao04999@outlook.com'
            mensagem = f"{saudacao}!\n\nHoje é aniversário de {nome} da {comissão}. Seu telefone é: {telefone} \n\nSegue em anexo o cartão de aniversário.\n\nAtenciosamente,\nSua Equipe"
        else:
            msg['To'] = 'jaojao04999@outlook.com'
            mensagem = f"{saudacao} {nome}!\n Feliz aniversário!"

        msg.attach(MIMEText(mensagem, 'plain'))

        with open(documento_salvo, 'rb') as file:
            anexo = MIMEApplication(file.read(), _subtype='docx')
            anexo.add_header('Content-Disposition', 'attachment', filename=documento_salvo)
            msg.attach(anexo)

        server = smtplib.SMTP('smtp.outlook.com', 587)
        server.starttls()
        server.login("jaojao04999@outlook.com", senha)  # Insira seu e-mail e senha aqui
        server.send_message(msg)
        server.quit()

    # Feche o documento modelo
    documentos.close()
else:
    print("Não há aniversariantes hoje.")
    # Feche o documento modelo
    documentos.close()
else:
    print("Não há aniversariantes hoje.")
