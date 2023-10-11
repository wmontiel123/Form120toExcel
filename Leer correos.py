import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
donebox = outlook.GetDefaultFolder(6).Folders("Postulaciones")
donebox2 = outlook.GetDefaultFolder(6).Folders("Seminarios")
donebox3 = outlook.GetDefaultFolder(6).Folders("Solicitud de Cotizaciones")

messages = inbox.Items

message = messages.GetLast()
body_json = message.body

unread_messages = []
unread_messages2 = []
unread_messages3 = []
for message in messages:
 if message.Unread == True:
    if "cv" in message.subject.lower():
     unread_messages.append(message)
    if "seminario" in message.subject.lower():
        unread_messages2.append(message)
    if "cotiza" in message.subject.lower():
     unread_messages3.append(message)


for message in unread_messages:
    message.Move(donebox)
    
for message in unread_messages2:
    message.Move(donebox2)
    
for message in unread_messages3:
    message.Move(donebox3)