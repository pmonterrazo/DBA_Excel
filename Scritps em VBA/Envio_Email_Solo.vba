Sub enviar_email()

set objeto_outlook = CreateObject("Outlook.Application")

Set Email = objeto_outlook.createitem(0)

Email.display

Email.to = "creditoecobranca@meggal.com.br"
Email.cc = "ti@meggal.com.br; camila.melo@meggal.com.br"
Email.bcc = "philipe.monterrazo@meggal.com.br"

Email.Subject = "Planilha de Cobrança Duttra MA"

Email.Body = "Olá," & Chr(10) & Chr(10) _
& "Segue planilha de Cobrança Duttra MA em Anexo" & Chr(10) & chr(10) _
& "Este email foi enviado de forma automatica" & Chr(10) & "Versão 2022.1"

Email.Attachments.Add (ThisWorkbook.Path & "\Planilha de Cobrança.xlsm")
Application.Wait (Now + TimeValue("0:00:02"))
Email.send


End Sub