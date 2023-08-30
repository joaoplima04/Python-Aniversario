import pandas as pd
from datetime import datetime, timedelta
from Cartoes_aniversario import parse_date

data = pd.read_excel("C:\\Users\\João Lucas\\Downloads\\Nova Planilha Aniversariantes.xlsx", parse_dates=['Aniversário'], date_parser=parse_date)

# Obtém a data atual
data_atual = datetime.now()

# Calcula o dia da semana (0 = segunda-feira, 6 = domingo)
hoje = data_atual.weekday()

hojee = data_atual.strftime('%d/%m')

terça = data_atual + pd.DateOffset(days=1)
quarta = data_atual + pd.DateOffset(days=2)
quinta = data_atual + pd.DateOffset(days=3)
sexta = data_atual + pd.DateOffset(days=4)
sabado = data_atual + pd.DateOffset(days=5)
domingo = data_atual + pd.DateOffset(days=6)


# Filtra os aniversariantes da semana atual
aniversariantes_semana_atual = data[data['Aniversário'].dt.strftime('%m/%d').isin(
            [data_atual.strftime('%m/%d'), terça.strftime('%m/%d'), quarta.strftime('%m/%d'), quinta.strftime('%m/%d'), sexta.strftime('%m/%d'), sabado.strftime('%m/%d'), domingo.strftime('%m/%d')])]
aniversariantes_notificados = []  # Lista para armazenar os aniversariantes notificados

# Gera os cartões de aniversário e armazena os aniversariantes notificados
for index, row in aniversariantes_semana_atual.iterrows():
    nome = row['Nomeado']
    cargo = row['Cargo']
    comissão = row['Comissão']
    email = row['Email']
    celular = row['Contato']
    uf = row['UF']
    sexo = row['Sexo']
    aniversario = row['Aniversário']

    # Formata a data como uma string no formato desejado
    data_formatada = aniversario.strftime('%d/%m')

    aniversariantes_notificados.append((email, celular, nome, comissão, cargo, uf, aniversario, data_formatada))

# Ordena a lista de aniversariantes notificados
if aniversariantes_notificados:
    aniversariantes_notificados.sort(key=lambda x: x[6])

# Adiciona o conteúdo da mensagem
mensagem = f"\nOs aniversariantes da semana são:\n\n"

for i, aniversariante in enumerate(aniversariantes_notificados):
    email, celular, nome, comissao, cargo, uf, aniversario, data_formatada = aniversariante

    mensagem += f"Nome: {nome}\nUF: {uf}\nCargo: {cargo}\nComissão: {comissao}\n"
    mensagem += f"E-mail: {email}\nCelular: {celular}\n Aniversário: {data_formatada}\n\n"

print(mensagem)
