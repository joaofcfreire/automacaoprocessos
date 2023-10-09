# Automação de Processos - Aplicação Mercado de Trabalho
### Primeiro projeto de automação de processos - aplicação no mercado de trabalho

Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas por todo o Brasil.
Todo dia, pela manhã, a equipe de análise de dados calcula os chamados OnePages, e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.
Um OnePage é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de cada loja para saber os principais indicadores.
Além de enviar os e-mails para cada gerente de cada loja, é enviado um Ranking para a diretoria e é criado um backup em excel das databases utilizadas para o cálculo dos indicadores.

### O script em Python faz todo este processo automaticamente.

Primeiramente o código importa todas as bibliotecas necessárias e as databases utilizando o Pandas.
Depois disso, trata as databases e cria um dataframe para cada loja em um dicionário, onde a chave é o nome da loja e o valor o dataframe.
Após isso, é implementado uma lógica que verifica se a pasta backup da Loja já existe, caso não exista, se cria a pasta.
E então, salva todos os arquivos com a data e nome da loja em suas respectivas pastas, calcula os indicadores e envia os e-mails.

O corpo do e-mail, foi realizado em HTML utilizando lógicas de if/else para saber se os indicadores atigiram as metas, caso tenha atingido o cenário fica verde, caso contrário, fica vermelho.

### As metas são:

####Meta faturamento do dia por loja: R$1.000,00
####Meta faturamento do ano por loja: R$1.650.000,00

###Meta Diversidade de Produto do dia por loja: 4
####Meta Diversidade de Produto do ano por loja: 120

####Meta ticket médio tanto do ano quanto do dia: R$500,00

#### Antes de utilizar o script

####1- Altere na planilha databases/Emails.xlsx para seu e-mail ou um e-mail válido que corresponda ao nome do gerente.
####2- Altere o password para um password válido do seu g-mail para o envio dos e-mails.
####3- Realize o pip install de todas as biblioteas.

