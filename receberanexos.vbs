Sub GetAttachments()
On Error GoTo GetAttachments_err
' Declaração de variáveis
Dim ns As Namespace
Dim Inbox As MAPIFolder
Dim Item As Object
Dim Atmt As Attachment
Dim FileName As String
Dim i As Integer
Set ns = GetNamespace("MAPI")
Set Inbox = ns.GetDefaultFolder(olFolderInbox)
i = 0
'Verifica no seu inbox se existe algum anexo de acordo com o assunto especificado
If Inbox.Items.Count = 0 Then
MsgBox "There are no messages in the Inbox.", vbInformation, _
"Nothing Found"
Exit Sub
End If
' Check each message for attachments
For Each Item In Inbox.Items
'Este nome entre aspas duplas é o assunto do email que contém o anexo, MUDE DE ACORDO COM SEU EMAIL
If Item = "[DOU]" Then
' Save any attachments found
For Each Atmt In Item.Attachments
' Em filename você irá inserir o caminho de onde quer salvar seu anexo, MUDE DE ACORDO COM SEU AMBIENTE.
FileName = "C:\Users\alexandre.borges\Documents\TCU\" & Atmt.FileName
Atmt.SaveAsFile FileName
i = i + 1
Next Atmt
End If
Next Item
' Show summary message
If i > 0 Then
MsgBox "I found " & i & " attached files." _
& vbCrLf & "I have saved them into the C:\Users\carlos\Desktop\anexos." _
& vbCrLf & vbCrLf & "Have a nice day.", vbInformation, "Finished!"
Else
MsgBox "I didn't find any attached files in your mail.", vbInformation, "Finished!"
End If
' Clear memory
GetAttachments_exit:
Set Atmt = Nothing
Set Item = Nothing
Set ns = Nothing
Exit Sub
' Handle errors
GetAttachments_err:
'MsgBox "An unexpected error has occurred." _
& vbCrLf & "Please note and report the following information." _
& vbCrLf & "Macro Name: GetAttachments" _
& vbCrLf & "Error Number: " & Err.Number _
& vbCrLf & "Error Description: " & Err.Description _
, vbCritical, "Error!"
Resume GetAttachments_exit
End Sub