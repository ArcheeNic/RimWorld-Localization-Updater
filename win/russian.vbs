' Получение пути
Function getCurrentPath()
	strPath = Wscript.ScriptFullName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(strPath)
	strFolder = objFSO.GetParentFolderName(objFile)
	getCurrentPath = strFolder & "\Mods\Core\Languages" 
End Function

' Скачивание локализации
Function downloadLocalization(InetFile,localFile)
	 
	Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	WinHttp.Open "GET", InetFile, False
	WinHttp.Send
	WinHttp.waitForResponse = True
	 
	Set oADOStream = CreateObject("ADODB.Stream")
	oADOStream.Mode = 3
	oADOStream.Type = 1
	oADOStream.Open
	oADOStream.Write WinHttp.responseBody
	 
	oADOStream.SaveToFile localFile, 2
	 
	Set WinHttp = Nothing
	Set oADOStream = Nothing
End Function

' Удаление архива
Function DeleteArchive(FileName)
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(main_FolderLang) Then 
fso.DeleteFolder main_FolderLang
End If

fso = Nothing
End Function


main_InetFile = "https://github.com/Ludeon/RimWorld-ru/archive/master.zip"
main_strFolder = getCurrentPath()
main_Lang="Russian"
main_zipFile = main_strFolder & "\" & main_Lang & ".zip"
main_FolderLang= main_strFolder & "\" & main_Lang

Call downloadLocalization(main_InetFile,main_zipFile)

Set fso = CreateObject("Scripting.FileSystemObject")
If NOT fso.FolderExists(main_FolderLang) Then
   fso.CreateFolder(main_FolderLang)
End If

'Extract the contants of the zip file.
Set objShell = CreateObject("Shell.Application")
Set ObjectInZip = objShell.NameSpace(main_zipFile)
Set FilesInZip=ObjectInZip.Items()

objShell.NameSpace(main_strFolder).copyHere FilesInZip, 16

If fso.FolderExists(main_FolderLang) Then 
fso.DeleteFolder main_FolderLang
End If


If fso.FileExists(main_zipFile) Then 
fso.DeleteFile main_zipFile
End If
