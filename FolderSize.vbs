dim oFS, oFolder, oFolder1
dim oFSub, s
Set oFS = wscript.createobject("Scripting.FileSystemObject")
'Set oFS1 = wscript.createobject("Scripting.FileSystemObject")
Set oFolder = oFS.GetFolder("folder path")
'Set oFolder = oFS.GetFolder("folder path")
Set oFSub = oFolder.SubFolders

if oFS.FolderExists(oFolder) Then
	msgbox "Exist"
	s =""
	for each f1 in oFSub
		's = s & " " & f1.name & "  " & f1.Size /1024\1024 & " MB" & vbCrLf
		s = s & " " & f1.name &  " " & f1.size /1024\1024 & vbCrLf
	Next
msgbox s
End If
set oFolder1 = oFS1.GetFolder("folder path")

Set myApp = CreateObject("Outlook.Application")
Set myItem = myApp.CreateItem(0)

myItem.To = "email id"
myItem.Subject = "F Drive info"

msgbox oFolder.Size
'msgbox 31 - oFolder1.Size /1024\1024\1024

ShowFolderDetails oFolder, oFolder1
'ShowFolderDetails oFolder

sub ShowFolderDetails(oF, oF1)
'sub ShowFolderDetails(oF)
	dim F
	'	wscript.echo oF.Name & "Size = " & oF.Size /1024\1024\1024 & " GB"
	'	wscript.echo oF.Name & "No. of Files = " & oF.Files.count
	'	wscript.echo oF.Name & "No. of Folders = " & oF.SubFolders.count
	'	for each F in oF.subfolders
	'		showfolderdetails(F)
	'	next
	myItem.HTMLBody = "<H3> PFB F Drive Usage </H3>"
	myItem.HTMLBody = myItem.HTMLBody + "Used Space = " & oF.Size /1024\1024\1024 & " GB <BR>"
	myItem.HTMLBody = myItem.HTMLBody + "Free Space = " & 170 - (oF.Size /1024\1024\1024) & " GB <BR><BR><BR>"
		
	myItem.HTMLBody = myItem.HTMLBody + "<H3> PFB H Drive Usage </H3>"
	myItem.HTMLBody = myItem.HTMLBody + "Used Space = " & oF1.Size /1024\1024\1024 & " GB <BR>"
	myItem.HTMLBody = myItem.HTMLBody + "Free Space = " & 31 - (oF1.Size /1024\1024\1024) & " GB <BR><BR><BR>"
	
end sub

myItem.HTMLBody = myItem.HTMLBody + "Regards," & "<BR>"
myItem.HTMLBody = myItem.HTMLBody + "Bharath Kumar S" & "<BR>"
myItem.HTMLBody = myItem.HTMLBody + "This is auto generated mail."

myItem.Send

