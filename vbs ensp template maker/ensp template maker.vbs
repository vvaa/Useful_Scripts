'�ű����ܣ���ģ���е�@@�滻�����֣�Ȼ�󱣴浽��Ӧ�ļ��У�ѹ����zip�ļ������������޸�ensp�豸��
'------------------------------------------------------------------


SrcTxt_ = "@@"
SrcFile_ = "template.cfg"

'------------------------------------------------------------------


'On Error Resume Next

For n = 1 to 39
	if n < 10 then
		NewFolder_ = "00000000-0000-0000-0000-00000000000" & n
	else
		NewFolder_ = "00000000-0000-0000-0000-0000000000" & n
	end if
	CreateFolder_ NewFolder_
	Replace_ SrcTxt_, n, SrcFile_, NewFolder_ & "\vrpcfg.cfg"
	MyZip NewFolder_ & "\vrpcfg.cfg", NewFolder_ & "\vrpcfg.zip"
	DeleteFile_ NewFolder_ & "\vrpcfg.cfg"	
Next

Msgbox "Done!"


'------------------------------------------------------------------


Sub Replace_(SrcTxt_, DstTxt_, SrcFile_, DstFile_)

	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(SrcFile_) 
	S_ = FILE_.ReadAll	
	
	S_ = Replace(S_, SrcTxt_, DstTxt_)
	Set FILE_ = FSO_.CreateTextFile(DstFile_)
	FILE_.Write S_
	
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing
	
End Sub


Sub CreateFolder_(Fname_)
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.CreateFolder(Fname_)
	Set FSO_ = Nothing
End Sub

Sub DeleteFile_(Fname_)			'ɾ���ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.GetFile(Fname_)
	FILE_.attributes = 0
	FILE_.delete
End Sub


'------------------------------------------------------------------


Sub MyZip(Scr_, Dst_ )

	Path_ = createobject("Scripting.FileSystemObject").GetFolder(".").Path
	Scr_ = Path_ & "\" & Scr_
	Dst_ = Path_ & "\" & Dst_
	HisZip Scr_, Dst_
	
End Sub



'--------------------------���²�����д�������ã���Ҫ����·��------------


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

