import getpass
import pandas as pd
from datetime import datetime
from mailmerge import MailMerge
from O365 import Account, MSGraphProtocol
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
    print("Bom dia Tarcizo!\n")
elif hora < 18:
    print("Boa tarde Tarcizo!\n")
else:
    print("Boa noite Tarcizo!\n")

senha = input('Digite a sua senha: ')

# Carrega os dados da planilha e converte a coluna "Aniversário" para o tipo datetime
dados = pd.read_excel('C:\\Users\\João Lucas\\Downloads\\Aniversáriantes Julho3.xlsx', parse_dates=['Aniversãrio'], date_parser=parse_date, converters={'Emails': convert_email})

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

        nome = name

        documentos.merge(Nome=nome)  # Substitua NOME pelo nome da variável no modelo do documento

        # Salva cada cartão de aniversário como um documento separado
        documento_salvo = f"C:\\Users\\João Lucas\\Documents\\{nome}.docx"
        documentos.write(documento_salvo)

        # Envia o cartão de aniversário por e-mail
        msg = MIMEMultipart()
        msg['Subject'] = 'Cartão de Aniversário'
        msg['From'] = 'joao.lima@oab.org.br'
        if pd.isna(email):
            msg['To'] = 'jaojao04999@outlook.com'
            mensagem = f"Boa tarde!\n\nHoje é aniversário de {nome} da {comissão}. Seu telefone é: {telefone} \n\nSegue em anexo o cartão de aniversário.\n\nAtenciosamente,\nSua Equipe"
        else:
            msg['To'] = 'jaojao04999@outlook.com'
            mensagem = f"Boa tarde {nome}!\n Estou aqui por meio desta mensagem em nome de toda a GAC para te desejar um feliz aniversário!"

        msg.attach(MIMEText(mensagem, 'plain'))

        with open(documento_salvo, 'rb') as file:
            anexo = MIMEApplication(file.read(), _subtype='docx')
            anexo.add_header('Content-Disposition', 'attachment', filename=documento_salvo)
            msg.attach(anexo)

        # Autenticação usando Microsoft Authenticator
        credentials = ('joao.lima@oab.org.br', senha)  # Insira seu e-mail e senha aqui

        account = Account(credentials, protocol=MSGraphProtocol())

        # Define os escopos solicitados
        scopes = ['basic', 'message_all']

        # Se não houver token válido, será necessário autenticar novamente
        if not account.is_authenticated:
            if account.authenticate(scopes=scopes):
                account.connection.save_token(token_path='.', token_filename='o365_token.txt')

        # Cria um novo e-mail
        mail = account.new_message()
        mail.subject = "Cartão de Aniversário"
        mail.body = mensagem

        # Anexa o documento ao e-mail
        attachment = f"C:\\Users\\João Lucas\\Documents\\{nome}.docx"
        mail.attachments.add(attachment)

        # Define o destinatário e envia o e-mail
        mail.to.add(email)
        mail.send()

    # Feche o documento modelo
    documentos.close()
else:
    print("Não há aniversariantes hoje.")

