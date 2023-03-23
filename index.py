import pandas as pd
import smtplib
import email.message as emessage

# importar a base de dados
tabela_vendas = pd.read_excel(r'C:\Users\medei\OneDrive\Documentos\Raspagem de Dados\Vendas.xlsx')

# visualizar a base de dados
"""pd.set_option(opção, valor)"""
pd.set_option('display.max_columns', None)
#print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
#print(faturamento)
# qnt de produtos vendidos por loja
qnt_produtos = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
# print(qnt_produtos)

# ticket médio por produto em cada loja
#print('-'*50)
ticket_medio = (faturamento['Valor Final'] / qnt_produtos['Quantidade']).to_frame()
# print(ticket_medio)

# enviar um email com o relatório 
def enviar_email():  
    corpo_email = f'''
    <p>Prezados,<p>

    <p>Segue em anexo o Relatório de Vendas por cada loja.<p>

    <p>Faturamento:<p>
    {faturamento.to_html()}

    <p>Quantidade vendida:<p>
    {qnt_produtos.to_html()}

    <p>Ticket Médio dos Produtos em cada loja:<p>
    {ticket_medio.to_html()}

    <p>Qualquer dúvidas estou à disposição.<p>

    <p>Att..<p>
    <p>Hércules<p>

    '''


    msg = emessage.Message()
    msg['Subject'] = "Relatorio de Vendas"
    msg['From'] = 'medeiroshercules23@gmail.com'
    msg['To'] = 'medeiroshercules23@gmail.com'
    password = 'imunowrrkciooldp' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


enviar_email()
