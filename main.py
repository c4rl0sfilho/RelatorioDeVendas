import pandas as pd
import win32com.client as win32

def importar_dados(caminho_arquivo):
    tabela = pd.read_excel(caminho_arquivo)
    pd.set_option('display.max_columns', None)
    return tabela

def calcular_faturamento(tabela):
    return tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

def calcular_quantidade(tabela):
    return tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

def calcular_ticket_medio(faturamento, quantidade):
    ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
    ticket_medio.columns = ['Ticket Médio']
    return ticket_medio

def enviar_email(destinatario, faturamento, quantidade, ticket_medio):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.Subject = 'Relatório de Vendas por Loja'
    mail.HTMLBody = f'''
    <p>Prezados,</p>

    <p>Segue o Relatório de Vendas por cada Loja.</p>

    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    <hr>
    <p>Quantidade Vendida:</p>
    {quantidade.to_html()}
    <hr>
    <p>Ticket Médio dos Produtos em cada Loja:</p>
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

    <p>Qualquer dúvida estou à disposição.</p>

    <p>Att.,<br>
    Lira</p>
    '''

    mail.Send()
    print('Email enviado com sucesso!')

def main():
    # Caminho do arquivo de vendas
    caminho_arquivo = 'Vendas.xlsx'
    destinatario = 'ce04435@gmail.com'

    # Processar dados
    tabela_vendas = importar_dados(caminho_arquivo)
    faturamento = calcular_faturamento(tabela_vendas)
    quantidade = calcular_quantidade(tabela_vendas)
    ticket_medio = calcular_ticket_medio(faturamento, quantidade)

    # Enviar relatório por e-mail
    enviar_email(destinatario, faturamento, quantidade, ticket_medio)

if __name__ == "__main__":
    main()
