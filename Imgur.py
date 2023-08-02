import pandas as pd
from datetime import datetime
from mailmerge import MailMerge
from docx2pdf import convert
import subprocess
import os
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pdf2image import convert_from_path
import requests


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

        documentos.merge(Nome=nome, Apelido=genero, CEP=abreviacao)  # Substitua NOME pelo nome da variável no modelo do documento

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

        aniversariantes_notificados.append((email, celular, nome, comissão, cargo, uf, imagem_cartao, documento_pdf))

        # Feche o documento modelo
        documentos.close()

    return aniversariantes_notificados


def notifica_aniversariantes(aniversariantes_notificados):
    if aniversariantes_notificados:

        # Adiciona o conteúdo da mensagem
        mensagem = f"\nHoje é aniversário dos seguintes colaboradores:\n\n"

        for i, aniversariante in enumerate(aniversariantes_notificados):
            email, celular, nome, comissao, cargo, uf, url_imagem, documento_pdf = aniversariante

            mensagem += f"Nome: {nome}\nUF: {uf}\nCargo: {cargo}\nComissão: {comissao}\n"
            mensagem += f"E-mail: {email}\nCelular: {celular}\n\n"

        print(mensagem)


def send_email(email_user, to_address, subject, body, attachment_path, smtp_server, smtp_port, email_password):
    try:
            with open(attachment_path, "rb") as attachment:
                encoded_pdf = base64.b64encode(attachment.read()).decode('utf-8')

            # Crie o objeto de e-mail com cabeçalhos apropriados
            message = f"""From: {email_user}
    To: {to_address}
    Subject: {subject}
    MIME-Version: 1.0
    Content-Type: multipart/mixed; boundary="BOUNDARY"

    --BOUNDARY
    Content-Type: text/html

    {body}

    --BOUNDARY
    Content-Type: application/pdf; name="{os.path.basename(attachment_path)}"
    Content-Disposition: attachment; filename="{os.path.basename(attachment_path)}"
    Content-Transfer-Encoding: base64

    {encoded_pdf}
    --BOUNDARY--
    """

            # Conecte-se ao servidor SMTP e envie o e-mail
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(email_user, email_password)
                server.sendmail(email_user, to_address, message)

            print(f"Email enviado para {to_address}")

            return encoded_pdf  # Retorna o valor de encoded_pdf
    except Exception as e:
        print(f"Falha ao enviar o email para {to_address}: {e}")
        return None


def upload_imagem_para_imgur(image_path):
    url = "https://api.imgur.com/3/upload"
    client_id = "cee22ea0dbec690"  # Substitua pelo seu Client-ID
    headers = {
        "Authorization": f"Client-ID {client_id}"
    }

    with open(image_path, "rb") as image_file:
        payload = {
            "image": image_file.read(),
            "type": "file"
        }

    response = requests.post(url, headers=headers, data=payload)
    data = response.json()

    if data["success"]:
        link = data["data"]["link"]
        print(f"Imagem enviada com sucesso! Link: {link}")
    else:
        print("Erro ao enviar a imagem para o Imgur.")


def main():

    hoje = "01/08"  # Defina manualmente a data de hoje

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

    mes = "Agosto"
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

    resposta = input("Deseja enviar o email?(s/n): ")

    if resposta.upper() == "S":
        smtp_server = 'smtp.gmail.com'  # Substitua pelo servidor SMTP que você está usando
        smtp_port = 587  # Porta padrão para o servidor SMTP (pode variar de acordo com o provedor de e-mail)
        email_user = 'joao.plima@sempreceub.com'  # Seu endereço de e-mail
        email_password = 'Sucesso15'  # Sua senha de e-mail

        # Upload das imagens para o Imgur e obtenção das URLs
        aniversariantes_com_imagens = []
        for aniversariante in aniversariantes_notificados:
            email, celular, nome, comissao, cargo, uf, imagem_cartao, documento_pdf = aniversariante
            url_da_imagem_enviada = upload_imagem_para_imgur(imagem_cartao)
            if url_da_imagem_enviada:
                aniversariantes_com_imagens.append(
                    (email, celular, nome, comissao, cargo, uf, url_da_imagem_enviada, documento_pdf))
            else:
                # Caso ocorra algum erro ao fazer o upload da imagem, continue sem a imagem no corpo do e-mail
                aniversariantes_com_imagens.append((email, celular, nome, comissao, cargo, uf, "", documento_pdf))


        for aniversariante in aniversariantes_com_imagens:
            email, celular, nome, comissao, cargo, uf, url_imagem, documento_pdf = aniversariante

            # Cria o objeto MIMEMultipart para enviar o e-mail
            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['To'] = str("jaolucasssp@gmail.com")
            msg['Subject'] = f"Feliz Aniversário {nome}!"

            # Incorpora a imagem codificada no corpo do e-mail, se disponível
            if url_imagem:
                mail_body = f"""<p>{saudacao} {nome},</p>
                    <p>Desejamos a você um Feliz Aniversário! Que este seja um dia especial e repleto de alegria.</p>
                    <p>Aqui está o seu cartão de aniversário:</p>
                    <p><img src="{url_imagem}" alt="Cartão de Aniversário" /></p>
                    <p>Atenciosamente,<br/>GAC Conselho Federal OAB</p>"""
            else:
                # Caso não haja imagem disponível, enviar apenas o texto do e-mail
                mail_body = f"""<p>{saudacao} {nome},</p>
                    <p>Desejamos a você um Feliz Aniversário! Que este seja um dia especial e repleto de alegria.</p>
                    <p>Atenciosamente,<br/>GAC Conselho Federal OAB</p>"""

            part = MIMEText(mail_body, 'html')
            msg.attach(part)

            # Abre o arquivo do cartão PDF e adiciona-o como anexo
            with open(documento_pdf, "rb") as pdf_file:
                attachment = MIMEApplication(pdf_file.read())
                attachment.add_header('Content-Disposition', f'attachment', filename=os.path.basename(documento_pdf))
                msg.attach(attachment)

            # Envia o e-mail
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(email_user, email_password)
            server.sendmail(email_user, 'jaolucasssp@gmail.com', msg.as_string())
            server.quit()
    else:
        print("Email não enviado")

    input("Pressione Enter para sair...")


if __name__ == "__main__":
    main()
