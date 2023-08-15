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


def convert_email(value):
    if isinstance(value, str):
        return value.strip()
    return value


def carrega_planilha(caminho_planilha):
    # Carrega os dados da planilha e converte a coluna "Aniversário" para o tipo datetime
    dados = pd.read_excel(caminho_planilha, parse_dates=['Aniversário'], date_parser=parse_date)
    return dados


def filtra_aniversariantes(data, hoje):
    # filtra os aniversáriantes do dia
    aniversariantes_do_dia = data[data['Aniversário'].dt.strftime('%d/%m') == hoje]

    # se for sexta-feira, filtra também os aniversáriantes de sábado e domingo
    if pd.to_datetime(hoje, format='%d/%m').day_name() == 'Friday':
        amanha = (pd.to_datetime(hoje, format='%d/%m') + pd.Timedelta(days=1)).strftime('%d/%m')
        depois_de_amanha = (pd.to_datetime(hoje, format='%d/%m') + pd.Timedelta(days=2)).strftime('%d/%m')
        aniversariantes_do_dia = data[data['Aniversário'].dt.strftime('%d/%m').isin([hoje, amanha, depois_de_amanha])]

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
    aniversariantes_notificados = []  # Lista para armazenar os aniversariantes notificados

    # Gera os cartões de aniversário e armazena os aniversariantes notificados
    for index, row in data.iterrows():
        documentos = MailMerge(template_path)
        nome = row['Nomeado']
        cargo = row['Cargo']
        comissão = row['Comissão']
        email = row['Email']
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

        # Salva cada cartão de aniversário como um documento separado
        documento_salvo = f"{output_dir}/{nome}.docx"
        documento_pdf = f"{output_dir}/{nome}.pdf"

        documentos.write(documento_salvo)

        # Converte o documento do Word em PDF
        convert(documento_salvo, documento_pdf)

        poppler_path = "C:\\Users\\João Lucas\\Downloads\\Release-23.07.0-0\\poppler-23.07.0\\Library\\bin"

        # Salva também o cartão de aniversário como imagem (no formato JPEG)
        imagem_cartao = f"{output_dir}/{nome}.jpg"
        images = convert_from_path(documento_pdf, poppler_path=poppler_path)
        images[0].save(imagem_cartao, 'JPEG')

        aniversariantes_notificados.append((email, celular, nome, comissão, cargo, uf, imagem_cartao, documento_pdf, abreviacao))

        # Feche o documento modelo
        documentos.close()

    return aniversariantes_notificados


def notifica_aniversariantes(aniversariantes_notificados):
    if aniversariantes_notificados:

        # Adiciona o conteúdo da mensagem
        mensagem = f"\nHoje é aniversário dos seguintes colaboradores:\n\n"

        for i, aniversariante in enumerate(aniversariantes_notificados):
            email, celular, nome, comissao, cargo, uf, url_imagem, documento_pdf, abreviacao = aniversariante

            mensagem += f"Nome: {nome}\nUF: {uf}\nCargo: {cargo}\nComissão: {comissao}\n"
            mensagem += f"E-mail: {email}\nCelular: {celular}\n\n"

        print(mensagem)


def main():

    hojee = datetime(2023, 8, 18) # Defina manualmente a data de hoje
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

    # Carrega os dados da planilha
    dados = carrega_planilha('C:\\Users\\João Lucas\\Downloads\\Nova Planilha Aniversariantes.xlsx')

    # Filtra os aniversariantes do dia
    aniversariantes_do_dia = filtra_aniversariantes(dados, hoje)

    mes = obter_pasta_mes(mes_atual)
    dia = hoje.replace("/", "-")
    output_dir = f'C:\\Users\\João Lucas\\Documents\\Aniversáriantes\\{mes}\\{dia}'

    # Cria o diretório do mês se ele não existir
    cria_diretorio_se_nao_existir(f'C:\\Users\\João Lucas\\Documents\\Aniversáriantes\\{mes}')

    # Cria o diretório do dia dentro do diretório do mês
    cria_diretorio_se_nao_existir(output_dir)

    # Gera os cartões de aniversário e armazena os aniversariantes notificados
    aniversariantes_notificados = gera_cartoes_aniversario(aniversariantes_do_dia,
                                                           'C:\\Users\\João Lucas\\Downloads\\Cartão de Aniversário.docx',
                                                           output_dir)

    # Notifica os aniversariantes do dia
    notifica_aniversariantes(aniversariantes_notificados)

    smtp_server = 'smtp.gmail.com'  
    smtp_port = 587
    email_user = 'comissoesgaccfoab2023@gmail.com'
    email_password = 'fsgecxglspcsjhzi'

    nomes_notificados = {}

    for aniversariante in aniversariantes_notificados:
        email, celular, nome, comissao, cargo, uf, imagem_cartao, documento_pdf, abreviacao = aniversariante

        # Verifica se o e-mail é um valor válido (não é float)
        if isinstance(email, str):
            if nome not in nomes_notificados:
                # Cria o objeto MIMEMultipart para enviar o e-mail
                msg = MIMEMultipart()
                msg['From'] = email_user
                msg['To'] = "joao.lima@oab.org.br"
                msg['Subject'] = f"Feliz Aniversário {nome}!"

                # Corpo do e-mail em formato HTML com a imagem
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

                # Incorpora a imagem codificada no corpo do e-mail
                with open(imagem_cartao, "rb") as image_file:
                    image = MIMEImage(image_file.read())
                    image.add_header('Content-ID', '<image>')
                    msg.attach(image)

                    # Abre o arquivo do cartão PDF e adiciona-o como anexo
                with open(documento_pdf, "rb") as pdf_file:
                    attachment = MIMEApplication(pdf_file.read())
                    attachment.add_header('Content-Disposition', f'attachment', filename=os.path.basename(documento_pdf))
                    msg.attach(attachment)

                    # Envia o e-mail
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(email_user, email_password)
                    server.sendmail(email_user, "joao.lima@oab.org.br", msg.as_string())

                print(f"Email enviado para {email}")
                nomes_notificados[nome] = True
            else:
                print(f"Email inválido para {nome}: {email}")


if __name__ == "__main__":
    main()
