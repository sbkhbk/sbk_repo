set objFSO = CreateObject("Scripting.FileSystemObject")

ZipFile="Folder Path"

set objFolder = objFSO.GetFolder(ZipFile)
set colFiles = objFolder.Files
ExtractTo="Folder Path\" & ptfname & "\"

'If the extraction location does not exist create it.
Set fso = CreateObject("Scripting.FileSystemObject")
If NOT fso.FolderExists(ExtractTo) Then
   fso.CreateFolder(ExtractTo)
End If

'Extract the contants of the zip file.
set objShell = CreateObject("Shell.Application")
set FilesInZip=objShell.NameSpace(ZipFile).items
objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)
Set fso = Nothing
Set objShell = Nothing
