# Aniversário de Colaboradores

Este script Python automatiza o envio de cartões de aniversário para colaboradores a partir de uma planilha. Ele gera cartões personalizados em formato PDF e imagem, e envia e-mails com esses anexos para os aniversariantes.

## Funcionalidades

- **Carrega uma planilha** com informações dos colaboradores, incluindo nome, e-mail, telefone, e data de aniversário.
- **Filtra** os aniversariantes do dia, considerando também os aniversariantes dos próximos dias se o dia atual for uma sexta-feira.
- **Gera cartões de aniversário** em formato DOCX e converte-os para PDF e imagem JPEG.
- **Envia e-mails** com os cartões de aniversário anexados para os aniversariantes.

## Requisitos

- **Python 3.x**
- **Bibliotecas Python**:
  - `pandas`
  - `mailmerge`
  - `docx2pdf`
  - `pdf2image`
  - `smtplib` (parte da biblioteca padrão do Python)
  - `email.mime` (parte da biblioteca padrão do Python)
  - `os` (parte da biblioteca padrão do Python)

1. Você pode instalar as bibliotecas necessárias usando `pip`:
```
pip install pandas mailmerge docx2pdf pdf2image
```

## Dependências Adicionais

Para converter PDFs em imagens, o script usa o `poppler-utils`. Instale-o com o seguinte comando:

```
sudo apt-get install poppler-utils
```
## Uso

1. Prepare a Planilha: A planilha deve estar no formato Excel (.xlsx) e deve conter as seguintes colunas: Nomeado, Cargo, Comissão, Email, Contato, UF, Sexo, e Aniversário.

2. Prepare o Modelo: Crie um modelo de cartão de aniversário no formato DOCX. O modelo deve ter campos que podem ser preenchidos pelo MailMerge, como Nome, Apelido, e CEP.

3. Configuração do Script:

* Atualize o caminho para a planilha de entrada na função carrega_planilha.
* Atualize o caminho para o modelo de cartão de aniversário na função gera_cartoes_aniversario.
* Ajuste o caminho para onde os arquivos de saída serão salvos.
* Configure as informações do servidor SMTP (endereço do servidor e credenciais de e-mail).

4. Execute o Script: Execute o script Python no seu ambiente local.
  ```
  python Main.py
  ```

## Considerações

* Segurança: Não compartilhe suas credenciais de e-mail. Use variáveis de ambiente ou arquivos de configuração para armazenar informações sensíveis.

* Testes: Antes de usar em produção, teste o script com um conjunto de dados menor para garantir que tudo funcione conforme o esperado.

## Contribuições

Se você tiver sugestões ou melhorias para o script, sinta-se à vontade para contribuir. Envie uma solicitação de pull ou abra uma issue para discutir melhorias.

## Licença
Este projeto está licenciado sob a Licença MIT.

## Contribuições
Se você tiver sugestões ou melhorias para o script, sinta-se à vontade para contribuir. Envie uma solicitação de pull ou abra uma issue para discutir melhorias.

## Contato
* Nome: João Lucas Pereira Lima
* Email: jaolucasssp@gmail.com
* linkedin: https://www.linkedin.com/in/joao-lucas-90219a273/
