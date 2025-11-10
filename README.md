Automação de E-mails & Relatório de Disparo

Este repositório contém dois projetos Python voltados para automação de e-mails via Outlook e tratamento dos resultados do último disparo, facilitando o acompanhamento e análise de campanhas de e-mail marketing.
 
Projeto1: Disparo de E-mails (Automacao_Email_Mkt.ipynb)
Descrição

Automatiza o envio de e-mails para uma lista de contatos presente em uma planilha Excel (Automacao.xlsx) usando Python e Outlook.

O e-mail é personalizado com o primeiro nome do contato.

Registra mensagens de sucesso e erro no console.

Tecnologias

Python 3.13

pandas

pywin32

Excel (via openpyxl, opcional para manipulação da planilha)

Estrutura do arquivo

Automacao_Email_Mkt.ipynb → notebook que faz a automação e mostra mensagens de envio no console.

Automacao.xlsx → planilha com os contatos (colunas obrigatórias: Empresa, Nome, E-mail).

Como usar

Coloque a planilha Automacao.xlsx na mesma pasta do notebook.

Abra o notebook no Jupyter Notebook.

Execute as células de cima para baixo.

Os e-mails serão enviados via Outlook, e o console exibirá quais foram enviados ou deram erro.

Projeto 2: Relatório do Último Disparo (Relatorio_Ultimo_Disparo.ipynb)
Descrição

Analisa os outputs do notebook de disparo de e-mails e gera um Excel resumido com os resultados:

Aba Sucesso → contatos que receberam o e-mail.

Aba Erros → contatos que não receberam, com a mensagem de erro.

Isso não envia e-mails novamente, apenas processa o histórico já registrado no notebook do disparo.

Tecnologias

Python 3.13

pandas

json (para ler o notebook .ipynb)

openpyxl (para gerar o Excel)

Estrutura do arquivo

Relatorio_Ultimo_Disparo.ipynb → notebook que lê Automacao_Email_Mkt.ipynb e gera relatório Excel.

Como usar

Certifique-se que Automacao_Email_Mkt.ipynb está na mesma pasta do notebook de relatório.

Abra Relatorio_Ultimo_Disparo.ipynb no Jupyter Notebook.

Execute a célula.

Um arquivo Relatorio_Ultimo_Disparo.xlsx será criado com abas de Sucesso e Erros.

Opcional: veja um resumo no console com total de e-mails enviados e com erro.

Observações

Esses notebooks não funcionam sem Outlook configurado e contatos válidos no Excel.

A análise do relatório não dispara e-mails novamente, garantindo segurança ao processar dados históricos.

Pode personalizar os textos de e-mail e planilha conforme sua campanha.

