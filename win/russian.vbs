' Получение пути
Function Custom_GetCurrentPath()
	strPath = Wscript.ScriptFullName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(strPath)
	strFolder = objFSO.GetParentFolderName(objFile)
	Custom_GetCurrentPath = strFolder & "\Mods\Core\Languages" 
	Set objFSO = Nothing
End Function

' Скачивание локализации
Function Custom_DownloadLocalization(InetFile,localFile)
	 
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

' Удаление файла
Function Custom_DeleteFile(FileName)

	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(FileName) Then 
	fso.DeleteFile FileName
	End If

	Set fso = Nothing
End Function


' Удаление папки
Function Custom_DeleteFolder(FolderName)
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FolderExists(FolderName) Then 
	fso.DeleteFolder FolderName
	End If

	Set fso = Nothing

End Function

' Создание папки
Function Custom_CreateFolder(FolderName)
	Set fso = CreateObject("Scripting.FileSystemObject")

	If NOT fso.FolderExists(FolderName) Then
	fso.CreateFolder(FolderName)
	End If

	Set fso = Nothing

End Function

' Распаковать архив
Function Custom_ExtractArchive(zipFile)
	Set objShell = CreateObject("Shell.Application")
	Set ObjectInZip = objShell.NameSpace(zipFile)
	Set FilesInZip= ObjectInZip.Items()
	objShell.NameSpace(main_strFolder).copyHere FilesInZip, 16
	
	Set objShell = Nothing
	Set ObjectInZip = Nothing
	Set FilesInZip = Nothing
End Function

' Проверить скачивание
Function Custom_CheckDownload(FileName)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If(fso.FileExists(FileName)) Then 
	Else
		WScript.Echo("Не удалось скачать перевод :-(")
		WScript.Quit()
	End If
	Set fso = Nothing
End Function

' Проверить рапаковку перевода
Function Custom_CheckTranslateFolder(Folder)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If(fso.FolderExists(Folder)) Then 
	Else
		WScript.Echo("Папка неправильно распаковалась :-(")
		WScript.Quit()
	End If
	Set fso = Nothing
End Function

Function Custom_getRights()
	If Not WScript.Arguments.Named.Exists("elevate") Then
	CreateObject("Shell.Application").ShellExecute WScript.FullName _
		, """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
	WScript.Quit
	End If
End Function

main_InetFile = "https://github.com/Ludeon/RimWorld-ru/archive/master.zip"
main_strFolder = Custom_GetCurrentPath()
main_Lang="Russian"
main_TranslateFolder=main_strFolder & "\" & "RimWorld-ru-master"
main_zipFile = main_strFolder & "\" & main_Lang & ".zip"
main_FolderLang= main_strFolder & "\" & main_Lang


Call Custom_getRights()
Call Custom_DeleteFile(main_zipFile)
Call Custom_DownloadLocalization(main_InetFile,main_zipFile)
Call Custom_CheckDownload(main_zipFile)
Call Custom_ExtractArchive(main_zipFile)
Call Custom_CheckTranslateFolder(main_TranslateFolder)
Call Custom_DeleteFolder(main_FolderLang)
Call Custom_DeleteFile(main_zipFile)