import pandas as pd
from datetime import datetime
from mailmerge import MailMerge
from docx2pdf import convert
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from pdf2image import convert_from_path


def parse_date(date_str):
    return pd.to_datetime(date_str, format='%d/%m', errors='coerce')


def carrega_planilha(caminho_planilha):
    dados = pd.read_excel(caminho_planilha, parse_dates=['Aniversário'], date_parser=parse_date)
    # Trata todos os emails como listas
    dados['Email'] = dados['Email'].apply(lambda x: x.split(";") if isinstance(x, str) else x)
    dados['Email'] = dados['Email'].apply(lambda x: [email.strip() for email in x] if isinstance(x, list) else x)
    return dados


def filtra_aniversariantes(data, hojee):
    hoje = hojee.strftime('%d/%m')

    # Se o dia for sexta-feira, filtre também os aniversáriantes de sábado e domingo
    if hojee.weekday() == 4:
        amanha = hojee + pd.DateOffset(days=1)
        depois_de_amanha = hojee + pd.DateOffset(days=2)
        aniversariantes_do_dia = data[data['Aniversário'].dt.strftime('%d/%m').isin(
            [hoje, amanha.strftime('%d/%m'), depois_de_amanha.strftime('%d/%m')])]
    else:
        aniversariantes_do_dia = data[data['Aniversário'].dt.strftime('%d/%m') == hoje]

    return aniversariantes_do_dia


def obter_pasta_mes(mes):
    meses = {
        1: 'Janeiro',
        2: 'Fevereiro',
        3: 'Março',
        4: 'Abril',
        5: 'Maio',
        6: 'Junho',
        7: 'Julho',
        8: 'Agosto',
        9: 'Setembro',
        10: 'Outubro',
        11: 'Novembro',
        12: 'Dezembro'
    }
    return meses.get(mes)


def cria_diretorio_se_nao_existir(diretorio):
    if not os.path.exists(diretorio):
        os.makedirs(diretorio)


def gera_cartoes_aniversario(data, template_path, output_dir):
    aniversariantes_notificados = []

    for index, row in data.iterrows():
        documentos = MailMerge(template_path)
        nome = row['Nomeado']
        cargo = row['Cargo']
        comissao = row['Comissão']
        emails = row['Email']
        celular = row['Contato']
        uf = row['UF']
        sexo = row['Sexo']

        if sexo == "M":
            genero = "o"
            abreviacao = ""
        else:
            genero = "a"
            abreviacao = "a"

        documentos.merge(Nome=nome, Apelido=genero, CEP=abreviacao)
        documento_salvo = f"{output_dir}/{nome}.docx"
        documento_pdf = f"{output_dir}/{nome}.pdf"
        documentos.write(documento_salvo)
        convert(documento_salvo, documento_pdf)
        poppler_path = "C:\\Users\\João Lucas\\Downloads\\Release-23.07.0-0\\poppler-23.07.0\\Library\\bin"
        imagem_cartao = f"{output_dir}/{nome}.jpg"
        images = convert_from_path(documento_pdf, poppler_path=poppler_path)
        images[0].save(imagem_cartao, 'JPEG')

        aniversariantes_notificados.append(
            (emails, celular, nome, comissao, cargo, uf, imagem_cartao, documento_pdf, abreviacao))

        documentos.close()

    return aniversariantes_notificados


def notifica_aniversariantes(aniversariantes_notificados):
    if aniversariantes_notificados:
        mensagem = f"\nHoje é aniversário dos seguintes colaboradores:\n\n"
        for i, aniversariante in enumerate(aniversariantes_notificados):
            email, celular, nome, comissao, cargo, uf, url_imagem, documento_pdf, abreviacao = aniversariante
            mensagem += f"Nome: {nome}\nUF: {uf}\nCargo: {cargo}\nComissão: {comissao}\n"
            mensagem += f"E-mail: {email}\nCelular: {celular}\n\n"
        print(mensagem)


def main():
    hojee = datetime.now()
    hoje = hojee.strftime("%d/%m")
    mes_atual = hojee.month
    agora = datetime.now()
    hora = agora.hour

    if hora < 12:
        print("Bom dia!\n")
        saudacao = "Bom dia"
    elif hora < 18:
        print("Boa tarde!\n")
        saudacao = "Boa Tarde"
    else:
        print("Boa noite!\n")
        saudacao = "Boa noite"

    dados = carrega_planilha('C:\\Users\\João Lucas\\Downloads\\Nova Planilha Aniversariantes.xlsx')
    aniversariantes_do_dia = filtra_aniversariantes(dados, hojee)
    mes = obter_pasta_mes(mes_atual)
    dia = hoje.replace("/", "-")
    output_dir = f'C:\\Users\\João Lucas\\Documents\\Aniversáriantes\\{mes}\\{dia}'
    cria_diretorio_se_nao_existir(f'C:\\Users\\João Lucas\\Documents\\Aniversáriantes\\{mes}')
    cria_diretorio_se_nao_existir(output_dir)
    aniversariantes_notificados = gera_cartoes_aniversario(aniversariantes_do_dia,
                                                           'C:\\Users\\João Lucas\\Downloads\\Cartão de Aniversário.docx',
                                                           output_dir)
    notifica_aniversariantes(aniversariantes_notificados)
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    email_user = 'comissoesgaccfoab2023@gmail.com'
    email_password = ''

    nomes_notificados = set()
    nomes_notificados_emails = {}

    for aniversariante in aniversariantes_notificados:
        emails, celular, nome, comissao, cargo, uf, imagem_cartao, documento_pdf, abreviacao = aniversariante
        try:
            if nome not in nomes_notificados:
                for email in emails:
                    email = email.strip()
                    msg = MIMEMultipart()
                    msg['From'] = email_user
                    msg['To'] = email
                    msg['Subject'] = f"Feliz Aniversário {nome}!"
                    msg['Bcc'] = "daniel.barros@oab.org.br"

                    mail_body = f"""<html>
                        <body>
                        <p>{saudacao} Dr{abreviacao}. {nome},</p>
                        <p>Desejamos a você um Feliz Aniversário! Que este seja um dia especial e repleto de alegria.</p>
                        <p>Segue abaixo e em anexo o respectivo cartão de aniversário:</p>
                        <img src="cid:image" alt="Cartão de Aniversário" width="700"/>
                        </body>
                        </html>
                        """

                    part = MIMEText(mail_body, 'html')
                    msg.attach(part)

                    with open(imagem_cartao, "rb") as image_file:
                        image = MIMEImage(image_file.read())
                        image.add_header('Content-ID', '<image>')
                        msg.attach(image)

                    with open(documento_pdf, "rb") as pdf_file:
                        attachment = MIMEApplication(pdf_file.read())
                        attachment.add_header('Content-Disposition', f'attachment',
                                              filename=os.path.basename(documento_pdf))
                        msg.attach(attachment)

                    with smtplib.SMTP(smtp_server, smtp_port) as server:
                        server.starttls()
                        server.login(email_user, email_password)
                        try:
                            server.sendmail(email_user, email, msg.as_string())
                        except Exception as e:
                            print(f"Email não enviado devido ao erro: {e}")
                        print(f"Email enviado para {nome} - {email}")
                        nomes_notificados.add(nome)
                        if nome not in nomes_notificados_emails:
                            nomes_notificados_emails[nome] = []
                        nomes_notificados_emails[nome].append(email)
            else:
                print(f"Email já enviado para {nome}")
        except Exception as e:
            print(f"O email não foi enviado pelo erro: {e}")


if __name__ == "__main__":
    main()
