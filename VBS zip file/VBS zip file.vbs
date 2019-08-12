'压缩文件到zip
'--------------------------以下调用----------------

'On Error Resume Next

MyZip "file.txt", "file.zip"

Msgbox "Done!"

'--------------------------以下调用----------------

Sub MyZip(Scr_, Dst_ )

	Path_ = createobject("Scripting.FileSystemObject").GetFolder(".").Path
	Scr_ = Path_ & "\" & Scr_
	Dst_ = Path_ & "\" & Dst_
	HisZip Scr_, Dst_
	
End Sub



'--------------------------以下不是我写，纯调用，需要绝对路径------------


Sub HisZip(ByVal mySourceDir, ByVal myZipFile) 
	Set fso = CreateObject("Scripting.FileSystemObject") 
	
	If fso.GetExtensionName(myZipFile) <> "zip" Then 
		Exit Sub 
	ElseIf fso.FolderExists(mySourceDir) Then 
		FType = "Folder" 
	ElseIf fso.FileExists(mySourceDir) Then 
		FType = "File" 
	FileName = fso.GetFileName(mySourceDir) 
	FolderPath = Left(mySourceDir, Len(mySourceDir) - Len(FileName)) 
	Else 
		Exit Sub 
	End If
	
	Set f = fso.CreateTextFile(myZipFile, True) 
	f.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0)) 
	f.Close 
	Set objShell = CreateObject("Shell.Application") 
	Select Case Ftype 
	Case "Folder" 
	Set objSource = objShell.NameSpace(mySourceDir) 
	Set objFolderItem = objSource.Items() 
	Case "File" 
	Set objSource = objShell.NameSpace(FolderPath) 
	Set objFolderItem = objSource.ParseName(FileName) 
	End Select 
	Set objTarget = objShell.NameSpace(myZipFile) 
	intOptions = 256 
	objTarget.CopyHere objFolderItem, intOptions 
	Do 
	WScript.Sleep 1000 
	Loop Until objTarget.Items.Count > 0 
End Sub

