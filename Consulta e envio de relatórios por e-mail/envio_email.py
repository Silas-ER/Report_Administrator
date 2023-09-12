#############################################################################################################
########################### FUNÇÃO PARA ENVIO DE E-MAIL COM A PLANILHA COMO ANEXO ###########################
#############################################################################################################

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def anexar_e_enviar(mes_atual, ano_atual):
    #CONFIGURAÇÕES DO SERVIDOR SMTP
    smtp_server = 'smtp.exemplo.com.br'
    smtp_port = 587  # Use a porta correta para o servidor SMTP do Terra
    smtp_username = 'email-envio@exemplo.com'
    smtp_password = 'senha@password'

    #CRIA O TITULO DO EMAIL
    msg = MIMEMultipart()
    msg['From'] = 'email-envio@exemplo.com'
    msg['To'] = 'email-destino@exemplo.com'
    msg['Subject'] = 'Relatorio de Lagostas para secretaria de aquicultura e pesca'

    #CORPO DO E-MAIL
    body = "Segue em anexo os dados referentes ao mês de {}".format(mes_atual)
    msg.attach(MIMEText(body, 'plain'))

    #ADICIONAR O ANEXO DO E-MAIL
    with open('relatorios/LAGOSTA - RELATÓRIO SECRETARIA DE AQUICULTURA E PESCA - MES - {} - ANO - {}.xlsx'.format(mes_atual, ano_atual), 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name='relatorios/LAGOSTA - RELATÓRIO SECRETARIA DE AQUICULTURA E PESCA - MES - {} - ANO - {}.xlsx'.format(mes_atual, ano_atual))

    #CABEÇALHOS DO ANEXO
    part['Content-Disposition'] = f'attachment; filename={"relatorios/LAGOSTA - RELATÓRIO SECRETARIA DE AQUICULTURA E PESCA - MES - {} - ANO - {}.xlsx".format(mes_atual, ano_atual)}'
    msg.attach(part)

    #CONECTAR AO SERVIDOR SMTP E ENVIAR E-MAIL
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  
        server.login(smtp_username, smtp_password)

        server.sendmail(smtp_username, 'email-destino@exemplo.com', msg.as_string()) #ENVIO DO E-MAIL
        print('Email enviado com sucesso')
    except Exception as e:
        print(f'Erro ao enviar email: {str(e)}')
    finally:
        server.quit()
