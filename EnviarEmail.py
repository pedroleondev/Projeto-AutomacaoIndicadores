import smtplib
import mimetypes
from email.message import EmailMessage
from datetime import date, datetime
import configura

############################
# CÓDIGO FEITO PELO ~LEON~ #
############################

def enviarEmail(titulo_mail,corpo_mail):

    try:
        #data de hoje
        data_atual = date.today()
        data_em_texto = ("{}/{}/{}".format(data_atual.day, data_atual.month,data_atual.year))
        #hora do momento
        dataAtual = datetime.now()
        horaAtual = dataAtual.strftime('%H:%M')

        # declarando variável da biblioteca p/ uso  + quem envia e quem receberá
        msg = EmailMessage()
        sender = configura.EMAIL_LOGIN
        recipient = configura.EMAIL_LOGIN

        #Configuração - De quem, Para quem, Título e Corpo do e-mail
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = '{} - {}, ({})'.format(titulo_mail,data_em_texto, horaAtual)
        body = corpo_mail
        msg.set_content(body)


        #Configurando anexo a se colocar no e-mail
        mime_type, _ = mimetypes.guess_type('C:\\PYTHON\\Projeto AutomacaoIndicadores\\Backup Arquivos Lojas\\Norte Shopping\\12_26_Norte Shopping.xlsx')
        mime_type, mime_subtype = mime_type.split('/')
        with open('C:\\PYTHON\\Projeto AutomacaoIndicadores\\Backup Arquivos Lojas\\Norte Shopping\\12_26_Norte Shopping.xlsx','rb') as file:
            msg.add_attachment(file.read(),
            maintype=mime_type,
            subtype=mime_subtype,
            filename='12_26_Norte Shopping.xlsx')

        #print(msg)

        #Configuração do SMTP ----------GMAIL----------
        mail_server = smtplib.SMTP_SSL('smtp.gmail.com', port = 465)
        mail_server.set_debuglevel(1)
        mail_server.login(configura.EMAIL_LOGIN, configura.EMAIL_PASSWORD)
        mail_server.send_message(msg)
        mail_server.quit()
        print('email enviado !')
        
    except:
        print('Não foi possível enviar e-mail, por favor, validar informações do código')

enviarEmail('testando o título','finalmente funcionou !')