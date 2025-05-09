import os
import pandas as pd
import win32com.client
import tkinter as tk
from tkinter import messagebox

# Caminho da pasta input
pasta_input = r"C:\Users\conbric2\OneDrive - Waters Corporation\Documentos\CONTAS A RECEBER\Projeto Aging Semáforo\Sistema_Cobranca\input"

# Encontrar o arquivo mais recente na pasta input
arquivos = [os.path.join(pasta_input, f) for f in os.listdir(pasta_input) if f.endswith(".xlsx")]
if not arquivos:
    print("Nenhum arquivo encontrado na pasta input.")
    exit()

arquivo_mais_recente = max(arquivos, key=os.path.getctime)

# Ler a aba "AGING_LIST" com cabeçalho na quarta linha
df_aging = pd.read_excel(arquivo_mais_recente, sheet_name="AGING_LIST", header=3)

# Ler a aba "MASTER_DATA" com cabeçalho na primeira linha
df_emails = pd.read_excel(arquivo_mais_recente, sheet_name="MASTER_DATA", header=0)

# Verificar se as colunas necessárias existem
if "OVERDUE DAYS" not in df_aging.columns or "CUSTOMER" not in df_aging.columns or "NOTA FISCAL" not in df_aging.columns:
    print("Erro: Colunas necessárias não encontradas na aba AGING_LIST.")
    exit()

if "CUSTOMER" not in df_emails.columns or "E-MAIL" not in df_emails.columns:
    print("Erro: Colunas necessárias na aba MASTER_DATA não encontradas.")
    exit()

# Conectar ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application")

# Lista de e-mails de teste
emails_teste = [
    "conbric2@partner.waters.com",
    "Juliana_Silva@waters.com",
    "condalm@partner.waters.com",
    "confoli@partner.waters.com",
    "Solange_Cavanha@waters.com",
    "Alan_Carvalho@waters.com",
    "jessica_silva@waters.com"
]

def enviar_cobranca(df_filtrado, titulo_email, corpo_email_base):
    """Função para enviar e-mails com base no DataFrame filtrado, título e corpo do e-mail."""
    emails_enviados = 0
    for index, linha in df_filtrado.iterrows():
        cliente = linha["CUSTOMER"]
        nota_fiscal = linha["NOTA FISCAL"]

        for email_destinatario in emails_teste:
            # Criar e-mail
            mail = outlook.CreateItem(0)
            mail.To = email_destinatario
            mail.Subject = titulo_email

            corpo_email = corpo_email_base.format(número_da_nota=nota_fiscal)
            mail.HTMLBody = corpo_email
            mail.SentOnBehalfOfName = "contasareceber@waters.com"
            mail.Send()
            emails_enviados += 1
            print(f"E-mail enviado para {email_destinatario} sobre a NF {nota_fiscal} do cliente {cliente}")
    return emails_enviados

# Corpo do e-mail da primeira cobrança (para -10)
corpo_email_primeira_cobranca = """
<p>Prezado(a),</p>
<p>Verificamos que consta em aberto em nosso sistema o pagamento da NF <strong>{número_da_nota}</strong>.</p>
<p>Como não recebemos nenhuma notificação com o motivo do atraso, solicitamos que entre em contato conosco para regularizarmos o pagamento.</p>
<p>Caso o atraso se deve a problemas circunstanciais, chegaremos logo a um acordo de negociação.</p>
<p>Caso o pagamento já tenha sido efetuado, por favor, nos envie o comprovante bancário.</p>
<p>Em caso de dúvidas, entre em contato com o nosso time de <strong>contasareceber@waters.com</strong>.</p>

<hr>
<p><span style="color:#0073e6; font-weight: bold;">Accounts Receivable</span><br>
Waters Technologies do Brasil<br>
<a href="http://www.waters.com" style="color: black; text-decoration: none;">www.waters.com</a><br>
Alameda Tocantins, 125 – 27º andar<br>
Alphaville – Barueri/SP<br>
CEP: 06455-020</p>
"""

# Corpo do e-mail da segunda cobrança (para -20)
corpo_email_segunda_cobranca = """
<p>Prezado(a),</p>
<p>Verificamos que ainda consta em aberto em nosso sistema o pagamento da NF <strong>{número_da_nota}</strong>.</p>
<p>Como não recebemos nenhuma notificação referente ao motivo do atraso, solicitamos que entre em contato conosco para regularizarmos o pagamento.</p>
<p>Lembramos que a manutenção dessa pendência pode, em breve, interferir em novos atendimentos e aquisições.</p>
<p>Caso o pagamento já tenha sido efetuado, por gentileza, nos envie o comprovante bancário.</p>
<p>Em caso de dúvidas, entre em contato com o nosso time de <strong>contasareceber@waters.com</strong>.</p>

<hr>
<p><span style="color:#0073e6; font-weight: bold;">Accounts Receivable</span><br>
Waters Technologies do Brasil<br>
<a href="http://www.waters.com" style="color: black; text-decoration: none;">www.waters.com</a><br>
Alameda Tocantins, 125 – 27º andar<br>
Alphaville – Barueri/SP<br>
CEP: 06455-020</p>
"""

# Corpo do e-mail da terceira cobrança (para -30)
corpo_email_terceira_cobranca = """
<p>Prezado(a),</p>
<p>Esta é nossa terceira notificação formal referente à pendência de pagamento da Nota Fiscal nº <strong>{número_da_nota}</strong>, vencida em nosso sistema.</p>
<p>Apesar dos contatos anteriores, não identificamos o recebimento nem qualquer posicionamento quanto à regularização.</p>
<p>Informamos que, caso o pagamento não seja efetuado imediatamente, poderemos adotar medidas administrativas, como protesto em cartório e/ou suspensão de novos fornecimentos.</p>
<p>Nosso objetivo é a conciliação amigável, e estamos à disposição para eventuais esclarecimentos ou negociação.</p>
<p>Caso o pagamento já tenha sido efetuado, pedimos a gentileza de nos enviar o comprovante bancário para registro.</p>
<p>Em caso de dúvidas, entre em contato com o nosso time de <strong>contasareceber@waters.com</strong>.</p>

<hr>
<p><span style="color:#0073e6; font-weight: bold;">Accounts Receivable</span><br>
Waters Technologies do Brasil<br>
<a href="http://www.waters.com" style="color: black; text-decoration: none;">www.waters.com</a><br>
Alameda Tocantins, 125 – 27º andar<br>
Alphaville – Barueri/SP<br>
CEP: 06455-020</p>
"""

total_emails_enviados = 0

# Filtrar clientes com OVERDUE DAYS igual a -10 e enviar a primeira cobrança
df_filtrado_primeira_cobranca = df_aging[df_aging["OVERDUE DAYS"] == -10]
if not df_filtrado_primeira_cobranca.empty:
    titulo_primeira_cobranca = "Primeira Cobrança de Nota Fiscal Vencida - Waters Technologies do Brasil"
    emails_enviados_primeira = enviar_cobranca(df_filtrado_primeira_cobranca, titulo_primeira_cobranca, corpo_email_primeira_cobranca)
    total_emails_enviados += emails_enviados_primeira
    print(f"Primeira cobrança enviada para {emails_enviados_primeira} destinatários.")

# Filtrar clientes com OVERDUE DAYS igual a -20 e enviar a segunda cobrança
df_filtrado_segunda_cobranca = df_aging[df_aging["OVERDUE DAYS"] == -20]
if not df_filtrado_segunda_cobranca.empty:
    titulo_segunda_cobranca = "Segunda Cobrança de Nota Fiscal Vencida - Waters Technologies do Brasil"
    emails_enviados_segunda = enviar_cobranca(df_filtrado_segunda_cobranca, titulo_segunda_cobranca, corpo_email_segunda_cobranca)
    total_emails_enviados += emails_enviados_segunda
    print(f"Segunda cobrança enviada para {emails_enviados_segunda} destinatários.")

# Filtrar clientes com OVERDUE DAYS igual a -30 e enviar a terceira cobrança
df_filtrado_terceira_cobranca = df_aging[df_aging["OVERDUE DAYS"] == -30]
if not df_filtrado_terceira_cobranca.empty:
    titulo_terceira_cobranca = "Terceira Cobrança de Nota Fiscal Vencida – Waters Technologies do Brasil"
    emails_enviados_terceira = enviar_cobranca(df_filtrado_terceira_cobranca, titulo_terceira_cobranca, corpo_email_terceira_cobranca)
    total_emails_enviados += emails_enviados_terceira
    print(f"Terceira cobrança enviada para {emails_enviados_terceira} destinatários.")

# Exibir pop-up de conclusão
root = tk.Tk()
root.withdraw()
messagebox.showinfo("Cobrança Reativa", f"E-mails de cobrança reativa enviados com sucesso!\nTotal de e-mails enviados: {total_emails_enviados}")
