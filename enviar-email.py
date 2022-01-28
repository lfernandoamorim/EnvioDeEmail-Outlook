#Faz a importação da biblioteca Win32
import win32com.client as Win32
import csv

#utiliza o proglrama Outlook
outlook = Win32.Dispatch('outlook.application')

#Criar e-mail
email = outlook.CreateItem(0)


#Configuração do e-mail
nomeDestino = ''
emailDestino = ''
with open('email.csv','r') as arquivo:

    #envia e-mail para todos que estão no arquivo email.csv
    arquivo_csv = csv.reader(arquivo, delimiter=";")
    for linha in enumerate(arquivo_csv):
        nomeDestino = str(linha[1][0])
        emailDestino = str(linha[1][1])
        print("E-mail enviado para ",nomeDestino,"no e-mail:",emailDestino)

        #Email de Destino
        email.To = emailDestino
        #Assunto do Email
        email.Subject = "E-mail automático - NÃO RESPONDER"
        #Corpo da mensagem do Email. Para fucionar as variáveis dentro do {}, precisa solocar o f antes das """
        email.HTMLBody = f"""
        <p>Olá! <b>{nomeDestino}</b>, sou o assistente da <b>Advance Sistemas</b></p>

        <p>O sistema de automação foi concluído com sucesso e por isso, não se preocupe.</p>

        <p>Até a próxima...</p>
        """
        #Envia o Email
        email.Send()
        print("E-mail enviado para ",nomeDestino,"no e-mail:",emailDestino)

#Mensagem de confirmação de envio dos e-mails
print("E-mails enviados com sucesso!!!")
