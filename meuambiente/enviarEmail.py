import smtplib
import pandas as pd
import os
import sys
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# Descobrindo o diretório atual do executável

global caminho_imagem

def verifica_ambiente_execucao():

    if getattr(sys, 'frozen', False):
        dir_path = sys._MEIPASS
        print("Caminho da imagem frozen:", dir_path)
    else:
        dir_path = os.path.dirname(os.path.abspath(__file__))
        print("Caminho da imagem normal:", dir_path)

    # Criando o caminho absoluto para a imagem
    caminho_imagem = os.path.join(dir_path, 'LogoGranaPix.jpeg')

    # Verificando e imprimindo o caminho da imagem
    print("Caminho da imagem:", caminho_imagem)

    return caminho_imagem

    # Verificando e imprimindo o caminho da imagem
    print("Caminho da imagem:", caminho_imagem)

# # Verificando se o arquivo realmente existe
# if os.path.exists(caminho_imagem):
#     print("O arquivo da imagem existe.")
# else:
#     print("O arquivo da imagem não existe.")

# Definindo Datas
data_atual = datetime.date.today()

dias_desde_segunda = data_atual.weekday()  # Retorna o dia da semana (0 = segunda-feira, 1 = terça-feira, ..., 6 = domingo)
segunda_da_semana_passada = data_atual - datetime.timedelta(days=dias_desde_segunda) - datetime.timedelta(7)
domingo_da_semana_passada = segunda_da_semana_passada + datetime.timedelta(days=6)
segunda_da_semana_atual = segunda_da_semana_passada + datetime.timedelta(days=7) 
terca_da_semana_atual = segunda_da_semana_atual + datetime.timedelta(days=1)

segunda_da_semana_passada = segunda_da_semana_passada.strftime('%d/%m/%Y')
domingo_da_semana_passada = domingo_da_semana_passada.strftime('%d/%m/%Y')
segunda_da_semana_atual = segunda_da_semana_atual.strftime('%d/%m/%Y')
terca_da_semana_atual = terca_da_semana_atual.strftime('%d/%m/%Y')

print("Data da segunda-feira da semana passada:", segunda_da_semana_passada)
print("Data de domingo da semana passada:", domingo_da_semana_passada)
print("Data da segunda-feira da semana atual:", segunda_da_semana_atual)
print("Data da terça-feira da semana atual:", terca_da_semana_atual)

def enviar_email(destinatario, nome_corban, anexos):
    print("Entrou função enviar e-mail")
    caminho_imagem = verifica_ambiente_execucao() 
    # Configuração do servidor de e-mail
    servidor_email = "smtp.office365.com"  
    porta = 587
    remetente_email = "relatorios@granapix.com.br"
    senha = "Grana@comissao2023"

    # Configuração da mensagem
    msg = MIMEMultipart()
    msg['From'] = remetente_email
    msg['To'] = destinatario
    msg['Subject'] = f'Relatório de Comissões - Corban {nome_corban} - Semana {segunda_da_semana_passada} a {domingo_da_semana_passada}'

    corpo = f"""
        <p> Prezado Parceiro GranaPix,</p>
        <p> Segue o relatório de comissão referente ao pagamento de {terca_da_semana_atual}, abrangendo a produção da última semana, de {segunda_da_semana_passada} a {domingo_da_semana_passada}.</p>
        <p>Em caso de dúvidas, por favor, entre em contato conosco pelo e-mail relatorios@granapix.com.br.</p>
        <p>Agradecemos a parceria!</p>
        <p>Atenciosamente, Equipe Grana Pix</p>

    """
    msg.attach(MIMEText(corpo, 'html'))


    for anexo in anexos:
        with open(anexo, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(anexo)}",
            )
            msg.attach(part)

    # Conecta ao servidor de e-mail e envia a mensagem
    with smtplib.SMTP(servidor_email, porta) as server:
        server.starttls()
        server.login(remetente_email, senha)
        text = msg.as_string()
        server.sendmail(remetente_email, destinatario, text)

    print("E-mail enviado com sucesso!")

def processar_diretorio(url):
    for pasta_corban in os.listdir(url): #cria o caminho completo para cada item
        caminho_pasta_corban = os.path.join(url, pasta_corban)
        print("URL pasta corban:", pasta_corban)
        print("Caminho relativo pasta corban:",  caminho_pasta_corban)
    
        if os.path.isdir(caminho_pasta_corban):  #verifica se é um diretório
            # Encontra os arquivos Excel e PDF na pasta do usuário
            arquivos = os.listdir(caminho_pasta_corban)
            arquivo_excel = None
            arquivo_pdf = None
            
            for arquivo in arquivos:
                if arquivo.endswith('.xlsx'):
                    arquivo_excel = os.path.join(caminho_pasta_corban, arquivo)
                    print("Caminho relativo excel:",  arquivo_excel)
                elif arquivo.endswith('.pdf'):
                    arquivo_pdf = os.path.join(caminho_pasta_corban, arquivo)
                    print("Caminho relativo pdf:",  arquivo_pdf)
            
            if arquivo_excel and arquivo_pdf:
                # Lê os dados do arquivo Excel
                df = pd.read_excel(arquivo_excel)
                destinatario = "felipe.moura@granapix.com.br"
                nome_corban = "teste bot"
                # celular = df.at[2, 3]

                # Envia o e-mail
                enviar_email(destinatario, nome_corban, [arquivo_excel, arquivo_pdf])