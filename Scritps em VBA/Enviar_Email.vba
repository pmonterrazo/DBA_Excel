Sub enviar_email()

set objeto_outlook = CreateObject("Outlook.Application")

For linha = 2 to 4

Set Email = objeto_outlook.createitem(0)

Email.display

Email.to = Cells(linha, 1).Value
Email.cc = "ti@meggal.com.br"
Email.bcc = "philipe.monterrazo@meggal.com.br"

Email.Subject = "Planilha de Cobrança Duttra MA"

Email.Body = Cells(linha, 2).Value & "," & Chr(10) & Chr(10) _
& Cells(linha, 3).Value & Chr(10) & chr(10) _
& "Teste de Automação da Planilha de Cobrança" & Chr(10) & "Versão 2022.1"

Email.Attachments.Add (ThisWorkbook.Path & "Planilha de Cobrança.xlsm")
Application.Wait (Now + TimeValue("0:00:02"))
Email.send

Next

End Sub