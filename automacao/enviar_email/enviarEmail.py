import os
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
from mensagem_email import EmailBody


class ProcessarEmail:

    def __init__(self, server, port):
        self.contador_emails = 0
        self.emails_enviados = []
        self.smtp_server = server
        self.smtp_port = port

    def enviar_email(
        self, remetente, senha, destinatarios, conteudo_email,
        assunto=None
    ) -> None:

        text_email = conteudo_email.text_email

        for destinatario in destinatarios:
            destinatario = destinatario.strip()  # Remover espaços em branco
            # Transformar a mensagem em MIMEMultipart
            mime_multipart = MIMEMultipart()
            mime_multipart["From"] = remetente
            mime_multipart["To"] = destinatario
            mime_multipart["Subject"] = assunto or "Email sem assunto"

            mime_multipart.attach(MIMEText(text_email, 'html', 'utf-8'))

            try:
                with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:

                    self.contador_emails += 1
                    print(f'\nDestinatário {self.contador_emails}: {destinatario}')

                    server.ehlo()
                    server.starttls()
                    server.login(remetente, senha)
                    server.sendmail(remetente, destinatario,
                                    mime_multipart.as_string())
                    print("Verificações concluídas")
                self.emails_enviados.append(destinatario)

            except smtplib.SMTPAuthenticationError:
                print(
                    "Erro de autenticação: verifique o email e a senha do "
                    "remetente."
                )
            except smtplib.SMTPConnectError:
                print(
                    "Erro de conexão: não foi possível conectar ao servidor "
                    "SMTP."
                )
            except smtplib.SMTPRecipientsRefused:
                print(
                    f"Erro: {destinatario} foi recusado pelo servidor SMTP."
                )
            except smtplib.SMTPSenderRefused:
                print("Erro: O remetente foi recusado pelo servidor SMTP.")
            except smtplib.SMTPDataError:
                print("Erro: O servidor SMTP retornou um erro ao enviar os "
                      "dados.")
            except smtplib.SMTPException as e:
                print(f"Erro SMTP: {e}")
            except Exception as e:
                print(f"Erro não identificado ao enviar email: {e}")

    def display_email_result(self):
        if self.emails_enviados:
            print(f'\nEmail enviado com sucesso para:\n\n'
                  f'{"\n".join(
                      [email.strip() for email in self.emails_enviados])}')
        else:
            print('\nNenhum email enviado')


if __name__ == '__main__':

    load_dotenv()

    email_teste = ProcessarEmail('smtp.gmail.com', 587)
    email_teste.enviar_email(
        remetente=os.getenv("FROM_EMAIL"), senha=os.getenv("FROM_PASSWORD"),
        destinatarios=os.getenv("TO_EMAIL").split(','),
        conteudo_email=EmailBody(), assunto="Teste de envio de email"
    )
