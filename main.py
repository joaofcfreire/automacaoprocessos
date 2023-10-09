import pandas as pd
import pathlib, smtplib, email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#importando arquivos da database

emails = pd.read_excel(r'databases/Emails.xlsx')
lojas = pd.read_csv(r'databases/Lojas.csv', encoding = 'latin1', sep = ';')
vendas = pd.read_excel(r'databases/Vendas.xlsx')

#incluir nome da loja em vendas

vendas = vendas.merge(lojas, on='ID Loja')

#criando um dataframe para cada loja armazenado em um dicionário

dicionario_lojas = {}

for loja in lojas['Loja']:

    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

#definindo dia do indicador (a data mais recente da database de vendas)

dia_indicador = vendas['Data'].max()

#identificar se a pasta de backup de cada loja já existe
#caso não exista, cria a pasta com nome da loja

caminho_backup = pathlib.Path(r'backups')
arquivos_backup = caminho_backup.iterdir()

lista_arquivos = []

for arquivo in arquivos_backup:
    lista_arquivos.append(arquivo.name)

for loja in dicionario_lojas:
    if loja not in lista_arquivos:
        nova_pasta = (caminho_backup / loja).mkdir()
    #salvando arquivo nas pastas criadas
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.day, dia_indicador.month, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    #criando excel para cada loja na pasta criada da loja
    dicionario_lojas[loja].to_excel(local_arquivo)

#definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_medio_ano = 500

#calculando os indicadores (faturamento, diversidade de produtos e ticket médio) e enviando por e-mail

for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]
    vendas_loja = vendas_loja[['Código Venda', 'ID Loja', 'Produto', 'Quantidade', 'Valor Unitário', 'Valor Final']]
    vendas_loja_dia = vendas_loja_dia[['Código Venda', 'ID Loja', 'Produto', 'Quantidade', 'Valor Unitário', 'Valor Final']]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    #diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    #ticket medio
    valor_venda_ano = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda_ano['Valor Final'].mean()
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    #enviando e-mail para o gerente de cada loja

    nome_gerente = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtde_produtos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtde_produtos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticket_medio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticket_medio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'


    corpo_email = f'''
    <p>Bom dia, {nome_gerente}.</p>
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>
    
    <h2>{loja}</h2>
    
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtde_produtos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticket_medio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtde_produtos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticket_medio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>
    <br>
    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
    
    <p>Qualquer dúvida estou à disposição.</p>
    <p>att,</p>
    <p>João Freire</p>
    '''

    msg = MIMEMultipart()
    msg['From'] = 'joaofcfreire@gmail.com'
    msg['To'] = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    msg['Subject'] = "OnePage Dia {}/{} - Loja {}".format(dia_indicador.day, dia_indicador.month, loja)
    password = 'SEU_PASSWORD_GMAIL'
    msg.attach(MIMEText(corpo_email, 'html'))
    anexo = pathlib.Path.cwd() / caminho_backup / loja / '{}_{}_{}.xlsx'.format(dia_indicador.day, dia_indicador.month, loja)

    with open(anexo, 'rb') as file:
        attachment = MIMEApplication(file.read(), _subtype="xlsx")
        attachment.add_header('Content-Disposition', 'attachment', filename='{}_{}_{}.xlsx'.format(dia_indicador.day, dia_indicador.month, loja))
        msg.attach(attachment)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print(f'E-mail da Loja {loja} enviado!')

#Criando ranking para diretoria

faturamento_lojas_ano = vendas.groupby("Loja")[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas_ano.sort_values(by='Valor Final', ascending=False)

vendas_dia = vendas.loc[vendas['Data'] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby("Loja")[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

#Enviando rankings para diretoria por e-mail

corpo_email = f'''
    <p>Prezados, bom dia,</p>
    <br>
    <p><strong>Melhor loja do dia</strong> em faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0,0]:.2f}</p>
    <p><strong>Pior loja do dia</strong> em faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1,0]:.2f}</p>
    <br>
    <p><strong>Melhor loja do ano</strong> em faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0,0]:.2f}</p>
    <p><strong>Pior loja do ano</strong> em faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1,0]:.2f}</p>
    <br>
    <p>Segue Ranking do Dia {dia_indicador.day}/{dia_indicador.month} de todas as lojas:</p>
    {faturamento_lojas_dia.to_html()}
    <br>
    <p>Segue Ranking <strong>ANUAL</strong> de todas as lojas:</p>
    {faturamento_lojas_ano.to_html()}
    <br>
    <p>Qualquer dúvida estou à disposição.</p>
    <p>att,</p>
    <p>João Freire</p>'''

msg = email.message.Message()
msg['Subject'] = "Ranking do Dia/Ano {}/{} - Todas as Lojas".format(dia_indicador.day, dia_indicador.month)
msg['From'] = 'joaofcfreire@gmail.com'
msg['To'] = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
password = 'SEU_PASSWORD_GMAIL'
msg.add_header('Content-Type', 'text/html')
msg.set_payload(corpo_email)

s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('E-mail da Diretoria enviado!')