import smtplib
from email.mime.multipart                       import MIMEMultipart                    # Para formatação de mensagens
from email.mime.text                            import MIMEText 
from oauth2client.service_account               import ServiceAccountCredentials
import gspread, time

scope = ["https://spreadsheets.google.com/feeds",
         'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    '---credencial---.json', scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(
    '---credencial---').worksheet('Página1')
dados = sheet.get_all_values()

def email(destinatarios, title, aviso, user, senha):    
    
    server = smtplib.SMTP('smtp.outlook.com', 587)
    server.starttls()
    username = user
    password = senha

    from_addr = user

    server.login(username, password)
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['Subject'] = title

    if dadosPlanilha[2] == 'Nível alto':
        body = '''
        Olá!

        Esse e-mail é de nível alto!
        '''

    if dadosPlanilha[2] == 'Nível baixo':
        body = '''
        Olá!

        Esse e-mail é de nível baixo!
        ''' 
    
    msg.attach(MIMEText(body))
    text = msg.as_string()

    emails = destinatarios
    usuario = []
    cont = 0
    contSimbolo = 0    

    # Quantidade de ponto-vírgula
    for i in range(0,len(emails)):
        if emails[i] == ";":
            contSimbolo += 1
    
    # Destinatário de email único
    if contSimbolo == 0:
        posMin = 0
        posMax = len(emails)
        usuario.append(emails[posMin:posMax])
    
    # Destinatário com mais de um email
    for i in range(0,len(emails)):
        if emails[i] == ";":
            pos = i        
            contMax = contSimbolo

            if cont == 0:
                posMin = 0
                posMax = pos
                usuario.append(emails[posMin:posMax])

            if cont >= 1 and cont <= contMax-1:
                posMin = posMax + 1
                posMax = pos
                usuario.append(emails[posMin:posMax])
            
            if cont == contMax-1:
                posMin = posMax + 1
                posMax = len(emails)
                usuario.append(emails[posMin:posMax])

            cont += 1
            
    # Exibe os emails dos destinatários que estão sendo enviados
    print(f'Email(s): {usuario}')
    
    # Envio dos emails
    for i in range(0,len(usuario)):
        to_addrs = usuario[i]
        msg['To'] = to_addrs
        server.sendmail(user, to_addrs, text)

    # Fecha a conexão dos envios
    server.quit()

    print('Email(s) enviado(s) com sucesso!')
    

for i in range(1,3):
    dadosPlanilha = dados[i]
    email(dadosPlanilha[0], dadosPlanilha[1], dadosPlanilha[2], "meu_email@hotmail.com", open('password.txt').read().strip())

