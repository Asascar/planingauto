Sub RunAScriptRuleRoutine(MyMail As MailItem)
Dim strID As String
Dim olNS As Outlook.NameSpace
Dim msg As Outlook.MailItem
strID = MyMail.EntryID
Set olNS = Application.GetNamespace("MAPI")
Set msg = olNS.GetItemFromID(strID)
' do stuff with msg, e.g.
For Each att in msg.Attachments
att.SaveAsFile "C:\Users\alexandre.borges\Documents\TCU" & att.FileName
Next

Set att = Nothing
Set msg = Nothing
Set olNS = Nothing
End Sub