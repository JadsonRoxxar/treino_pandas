import pandas as pd
import smtplib
import email.message

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar basse de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('-' *50)
# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' *50)
# quantidade de produtos vendidos por loja
quantidade =  tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' *50)
# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório
def enviar_email():
    corpo_email = f"""
    <p>Prezados,</p>
    
    <p>Segue o Relatório de Vendas por cada loja.</p>
    
    <p>Faturamento:
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}</p>
    
    <p>Quantidade Vendida:
    {quantidade.to_html()}</p>
    
    <p>Ticket Médio dos Produtos em cada Loja:
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}</p>
    <p>Att..,</p>
    <p>Jadson</p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Relatório de Vendas por Loja"
    msg['From'] = 'jadsonrocha.jrda@gmail.com'
    msg['To'] = 'jadsonrocha.jrda@gmail.com'
    password = 'dmqh feks fpwt supa'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()