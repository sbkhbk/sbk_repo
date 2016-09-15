'Sub sendmail()
Const olFolderContacts = 10
sDistName = "Bangalore Team"

Set objoutlook = CreateObject("Outlook.application")
Set objnamespace = objoutlook.getnamespace("MAPI")

Set colcontacts = objnamespace.getdefaultfolder(olFolderContacts).items
intcount = colcontacts.Count

MsgBox intcount
For i = 1 To intcount
    If TypeName(colcontacts.Item(i)) = "DistListItem" Then
        Set objdistlist = colcontacts.Item(i)
            sEmails = "email id"
                If objdistlist.DLName = sDistName Then
                    For j = 1 To objdistlist.MemberCount
                        sEmails = sEmails & ";" & objdistlist.getMember(j).Address
                    Next
                    MsgBox sEmails
                End If
            End If
Next
        
Dim objoutlookmsg
Set objoutlookmsg = objoutlook.createItem(0)
With objoutlookmsg
    '.to = sEmails
    .to = "email id"
    .Subject = "Test mail sent thru script"
    '.body = "Hello," & vbCrLf & <B>"    This is just a test mail."</B> & vbCrLf & "Regards," & vbCrLf & "Bharath Kumar S"
    .HTMLBody = "Hello,<BR>This is just a test mail."
    .send
End With

Set objoutlookmsg = Nothing
Set objoutlook = Nothing
Set objnamespace = Nothing
'End Sub
