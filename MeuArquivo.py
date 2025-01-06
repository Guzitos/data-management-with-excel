import pandas as pd
import smtplib
import email.message
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

# faturamento por loja

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# quantidade de produtos vendidos

qnt_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qnt_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# enviar um email com o relatório

def enviar_email():
    corpo_email = f"""
    <p>Prezados,</p>
    <p>Segue o Relatório de vendas por cada loja.</p>
    
    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    
    <p>Quantidade:</p>
    {qnt_produtos.to_html()}
    
    <p>Ticket médio por cada loja:</p>
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
    
    <p>Qualquer dúvida estou a disposição.</p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Relatório de vendas por loja."
    msg['From'] = 'email'
    msg['To'] = 'email'
    password = 'password'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()