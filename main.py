import pandas as pd
from datetime import datetime
from mailmerge import MailMerge
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def parse_date(date_str):
    return pd.to_datetime(date_str, format='%d/%m', errors='coerce')

# Carrega os dados da planilha e converte a coluna "Aniversário" para o tipo datetime
dados = pd.read_excel('C:\\Users\\jaoja\\Downloads\\Aniversáriamtes.xls', parse_dates=['Aniversário'], date_parser=parse_date)

# Filtra os aniversariantes do dia
hoje = '30/06'
aniversariantes_do_dia = dados[dados['Aniversário'].dt.strftime('%d/%m') == hoje]

if not aniversariantes_do_dia.empty:
    # Carrega o modelo do documento de cartão de aniversário
    template = 'C:\\Users\\jaoja\\Downloads\\Document 3.docx'
    documentos = MailMerge(template)

    # Gera os cartões de aniversário e envia e-mail para o chefe
    for index, row in aniversariantes_do_dia.iterrows():
        nome = row['Nome']
        aniversario = row['Aniversário']
        email = row['Email']

        documentos.merge(NOME=nome)  # Substitua NOME pelo nome da variável no modelo do documento

        # Salva cada cartão de aniversário como um documento separado
        documento_salvo = f"C:\\Users\\jaoja\\Documents\\{nome}.docx"
        documentos.write(documento_salvo)

        # Envia o cartão de aniversário por e-mail
        msg = MIMEMultipart()
        msg['From'] = 'jaojao04999@outlook.com'
        msg['To'] = 'jaojao04999@outlook.com'
        msg['Subject'] = 'Cartão de Aniversário'

        mensagem = f"Olá, chefe!\n\nHoje é aniversário de {nome}.\n\nSegue em anexo o cartão de aniversário.\n\nAtenciosamente,\nSua Equipe"

        msg.attach(MIMEText(mensagem, 'plain'))

        with open(documento_salvo, 'rb') as file:
            anexo = MIMEApplication(file.read(), _subtype='docx')
            anexo.add_header('Content-Disposition', 'attachment', filename=documento_salvo)
            msg.attach(anexo)

        server = smtplib.SMTP('smtp.outlook.com', 587)
        server.starttls()
        server.login("jaojao04999@outlook.com", "sucesso15")  # Insira seu e-mail e senha aqui
        server.send_message(msg)
        server.quit()

    # Feche o documento modelo
    documentos.close()
else:
    print("Não há aniversariantes hoje.")
