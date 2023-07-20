import pandas as pd
from datetime import datetime
from mailmerge import MailMerge
from docx2pdf import convert
import subprocess
import os


def parse_date(date_str):
    return pd.to_datetime(date_str, format='%d/%m/%Y', errors='coerce')


def convert_email(value):
    if isinstance(value, str):
        return value.strip()
    return value


def carrega_planilha(caminho_planilha):
    # Carrega os dados da planilha e converte a coluna "Aniversário" para o tipo datetime
    dados = pd.read_excel(caminho_planilha, parse_dates=['Aniversario'], date_parser=parse_date)
    return dados


def filtra_aniversariantes(data, hoje):
    # filtra os aniversáriantes do dia
    aniversariantes_do_dia = data[data['Aniversario'].dt.strftime('%d/%m') == hoje]
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


def gera_cartoes_aniversario(data, data2, template_path, output_dir):
    aniversariantes_notificados = []  # Lista para armazenar os aniversariantes notificados

    # Gera os cartões de aniversário e armazena os aniversariantes notificados
    for index, row in data.iterrows():
        documentos = MailMerge(template_path)
        nome = row['Nomeado']
        cargo = row['Cargo']
        comissão = row['Comissão']
        aniversario = row['Aniversario']
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
        name = str(nome)
        # Procura os dados de contato e e-mail do aniversariante na segunda planilha (data2)
        aniversariante_info = data2[data2['Nome'] == name.upper()]
        if not aniversariante_info.empty:
            email = aniversariante_info['Emails'].iloc[0]
            celular = aniversariante_info['Telefones'].iloc[0]
        else:
            email = ''
            celular = ''

        aniversariantes_notificados.append((email, celular, nome, comissão, cargo, uf, documento_pdf))

        # Feche o documento modelo
        documentos.close()

    return aniversariantes_notificados


def notifica_aniversariantes(aniversariantes_notificados, saudacao, output_dir):
    if aniversariantes_notificados:
        # Cria um novo documento em formato Markdown
        documento = []

        # Adiciona o cabeçalho com a saudação
        documento.append(f" {saudacao}!")

        # Adiciona o conteúdo da mensagem
        mensagem = f"\nHoje é aniversário dos seguintes colaboradores:\n\n"

        for i, aniversariante in enumerate(aniversariantes_notificados):
            email, celular, nome, comissao, cargo, uf, documento_pdf = aniversariante

            mensagem += f"Nome: {nome}\nUF: {uf}\nCargo: {cargo}\nComissão: {comissao}\n"
            mensagem += f"E-mail: {email}\nCelular: {celular}\n\n"

        documento.append(mensagem)

        # Adiciona a assinatura
        assinatura = "Parabéns a todos!\n\nAtenciosamente,\nSua Equipe"
        documento.append(assinatura)

        # Salva o conteúdo do documento em um arquivo temporário em formato Markdown
        conteudo = '\n'.join(documento)
        dia = "31/07".replace("/", "-")
        arquivo_temporario = f"{output_dir}/{dia}.md"
        with open(arquivo_temporario, 'w', encoding='utf-8') as f:
            f.write(conteudo)

        # Converte o arquivo temporário para o formato do Word usando unoconv
        documento_salvo = f"{output_dir}/{dia}.docx"
        subprocess.run(['python', 'C:/Windows/unoconv.py', '-f', 'docx', '-o', documento_salvo, arquivo_temporario])

        print(f"Documento Word salvo em: {documento_salvo}")


def main():
    hoje = "31/07"  # Defina manualmente a data de hoje

    dados2 = pd.read_excel('C:\\Users\\João Lucas\\Downloads\\Aniversáriantes Julho3.xlsx', parse_dates=['Aniversãrio'], date_parser=parse_date)

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
    dados = carrega_planilha('C:\\Users\\João Lucas\\Downloads\\Cópia de Aniversariantes - Andamento 2 (1).xlsx')

    # Filtra os aniversariantes do dia
    aniversariantes_do_dia = filtra_aniversariantes(dados, hoje)

    mes = agora.month
    pasta_mes = obter_pasta_mes(mes)
    dia = hoje.replace("/", "-")
    output_dir = f'C:\\Users\\João Lucas\\Documents\\Aniversáriantes\\{pasta_mes}\\{dia}'

    # Cria o diretório do mês se ele não existir
    cria_diretorio_se_nao_existir(f'C:\\Users\\João Lucas\\Documents\\Aniversáriantes\\{pasta_mes}')

    # Cria o diretório do dia dentro do diretório do mês
    cria_diretorio_se_nao_existir(output_dir)

    # Gera os cartões de aniversário e armazena os aniversariantes notificados
    aniversariantes_notificados = gera_cartoes_aniversario(aniversariantes_do_dia, dados2, 'C:\\Users\\João Lucas\\Downloads\\Cartão de Aniversário.docx', output_dir)

    # Notifica os aniversáriantes do dia
    notifica_aniversariantes(aniversariantes_notificados, saudacao, output_dir)

    input("Pressione Enter para sair...")


if __name__ == "__main__":
    main()
