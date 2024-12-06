import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from config.email_config import SMTP_SERVER, SMTP_PORT, EMAIL_FROM, EMAIL_PASSWORD, EMAIL_LIST
import os

def send_emails(email_list, filename):
    # Verifica se o arquivo existe antes de tentar enviá-lo
    if not os.path.isfile(filename):
        print(f"Arquivo não encontrado: {filename}")
        return  # Sai da função se o arquivo não existir

    for person in email_list:
        body = """
        Prezados, bom dia!


        Segue em anexo base de itens faturados.

        Atenciosamente.

        Paulo Cardoso





        Obs.: E-mail enviado individualmente, de forma automática.
        """

        msg = MIMEMultipart()
        msg['From'] = EMAIL_FROM
        msg['To'] = person
        msg['Subject'] = "[SAE] Jundiaí - Faturamento"

        msg.attach(MIMEText(body, 'plain'))

        # Debug: Mostrando o caminho do arquivo a ser anexado
        print(f"Tentando anexar o arquivo: {filename}")

        try:
            # Abre o arquivo como um binário
            with open(filename, 'rb') as attachment:
                attachment_package = MIMEBase('application', 'octet-stream')
                attachment_package.set_payload(attachment.read())
                encoders.encode_base64(attachment_package)
                
                # Debug: Exibir o nome do arquivo que será anexado
                attachment_name = os.path.basename(filename)  # Extrai apenas o nome do arquivo
                print(f"Anexando o arquivo: {attachment_name}")
                attachment_package.add_header('Content-Disposition', f"attachment; filename={attachment_name}")
                msg.attach(attachment_package)

        except Exception as e:
            print(f"Erro ao abrir o arquivo para anexar: {e}")
            continue  # Se houver um erro ao abrir o arquivo, continue para o próximo destinatário

        text = msg.as_string()

        # Conecta-se ao servidor SMTP e envia o e-mail
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()  # Inicia a conexão segura
                server.login(EMAIL_FROM, EMAIL_PASSWORD)  # Faz login
                server.sendmail(EMAIL_FROM, person, text)  # Envia o e-mail
                print(f"Email enviado para: {person}")
        except Exception as e:
            print(f"Erro ao enviar o e-mail para {person}: {e}")

# Exemplo de uso
if __name__ == "__main__":
    email_list = EMAIL_LIST  # Adicione mais e-mails conforme necessário
    filename = "C:.\data\FaturadosDia.xlsx"  # Altere para o caminho correto do seu arquivo
    send_emails(email_list, filename)
