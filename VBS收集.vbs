˵�����ռ�һЩ�Լ����õ�vbs���������㸴��
Ϊ�˲�Ū�죬Լ����
������sub_��ͷ��
������fun_��ͨ
������var_��ͷ

==============================
��ͷ
On Error Resume Next


==============================

���ı��ļ�

Set var_fso = CreateObject("Scripting.FileSystemObject")
Set var_file = var_fso.OpenTextFile(var_txt) 
var_s = var_file.ReadAll

==============================
д���ı��ļ�

Set var_file = var_fso.CreateTextFile(var_txt)
var_file.Write var_s
var_file.Close

==============================
�滻�ı�

var_s = Replace(var_s, var_txt1, var_txt2)

==============================
�����滻�������滻������滻
��һ��1��ʾ��1��ʼ���ڶ���1��ʾÿ���滻1��

For var_i = 1 to 100
	var_s = Replace(var_s, var_txt1, var_i, 1, 1)
Next

==============================
������ʽ
Set var_re = New RegExp
var_re.Global = True
var_re.IgnoreCase = False
var_re.Pattern = "\d+\n"
For Each var_m in var_re.Execute(var_s)
	Msgbox var_m
Next

==============================
ת�����˫����

Msgbox """"

==============================
ѹ���ļ���zip��
�������������·��


Sub sub_hiszip(ByVal mySourceDir, ByVal sub_myzipFile) 
	Set fso = CreateObject("Scripting.FileSystemObject") 
	
	If fso.GetExtensionName(sub_myzipFile) <> "zip" Then 
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
	
	Set f = fso.CreateTextFile(sub_myzipFile, True) 
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
	Set objTarget = objShell.NameSpace(sub_myzipFile) 
	intOptions = 256 
	objTarget.CopyHere objFolderItem, intOptions 
	Do 
	WScript.Sleep 1000 
	Loop Until objTarget.Items.Count > 0 
End Sub

����Ϊ�İ�Ϊ���·��

Sub sub_myzip(var_src, var_dst )

	var_path = createobject("Scripting.FileSystemObject").GetFolder(".").Path
	var_src = var_path & "\" & var_src
	var_dst = var_path & "\" & var_dst
	sub_hiszip var_src, var_dst
	
End Sub


==============================
��ȡ��ǰ·��
var_path = Createobject("Scripting.FileSystemObject").GetFolder(".").Path
Msgbox var_path


==============================
ɾ���ļ�
Sub sub_deletefile(var_filename)			'ɾ���ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.GetFile(var_filename)
	var_file.attributes = 0
	var_file.delete
End Sub

==============================
�����ļ���
Sub sub_createfolder(var_filename)
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.CreateFolder(var_filename)
	Set var_fso = Nothing
End Sub

==============================
����������д�뵽�ļ�
Sub sub_writefile(var_arrin(), var_filename)		'������д�뵽�ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 2, True) 
	For i = 0 To UBound(var_arrin)
		var_file.WriteLine var_arrin(i)
	Next
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
End Sub

==============================
���ж�ȡ�ı��ļ�������
Function Readvar_file(var_filename)		'��ȡ�ļ�������
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 1)
	i = 0
	Do Until var_file.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = var_file.ReadLine
		i = i + 1
	Loop
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
 	Readvar_file = Arr_
End Function
==============================
ʾ�����룺�ָ��ı��ļ�Ϊ����ļ�

'���ı��ļ��ָ�������ָ���ļ��ֳ�������ȵ��ļ���
'------------------------------------------------------------------

SplitNuvar_m = 6		'Ҫ�ָ������
Srcvar_file = "Source.txt"		'Ҫ�ָ��Դ�ļ�



'On Error Resume Next
Arr1_ = Readvar_file(Srcvar_file)
For n = 1 to SplitNuvar_m
	sub_writefile GetPartArr_(Arr1_,SplitNuvar_m,n) , n & ".txt"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GetPartArr_(var_arrin(),NumAll_,NumPart_)	
	BlockSize_ = (Ubound(var_arrin)+1) \ NumAll_
	If NumAll_ = NumPart_ Then
		j = 0
		For k = BlockSize_ * (NumPart_-1)  to Ubound(var_arrin)
			redim preserve ArrTmp_(j)
			ArrTmp_(j) = var_arrin(k)
			j = j + 1
		Next
	Else
		j = 0
		For k = BlockSize_ * (NumPart_-1) to BlockSize_ * NumPart_
			redim preserve ArrTmp_(j)
			ArrTmp_(j) = var_arrin(k)
			j = j + 1
		Next
	End If
 	GetPartArr_ = ArrTmp_
End Function

Function Readvar_file(var_filename)		'��ȡ�ļ�������
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 1)
	i = 0
	Do Until var_file.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = var_file.ReadLine
		i = i + 1
	Loop
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
 	Readvar_file = Arr_
End Function

Sub sub_writefile(var_arrin(), var_filename)		'������д�뵽�ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 2, True) 
	For i = 0 To UBound(var_arrin)
		var_file.WriteLine var_arrin(i)
	Next
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
End Sub



==============================
�ϲ��ı��ļ�

Sub Mergevar_file(Srcvar_filename, Dstvar_filename)		'�ϲ��ı��ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set FileSrc_ = var_fso.OpenTextFile(Srcvar_filename, 1, True, -2)
	Set Filevar_dst = var_fso.OpenTextFile(Dstvar_filename, 8, True)
	Do Until FileSrc_.AtEndOfStream
		Filevar_dst.WriteLine(FileSrc_.Readline)
	Loop
	FileSrc_.Close
	Filevar_dst.Close
End Sub

==============================
��������ʾ����������������
'һ���򵥵������ֵ����ɽű����������գ����֣����õ���Ϣ��������ϳɿ��ܵ����롣
'������ϵĶ��峤�ȣ������ɶ���м��ļ�����1��2��3�����������ϲ���һ���ļ���
'------------------------------------------------------------------

SrcStr_ = "Source.txt"		'������ȡ��Դ�ļ�
DstStr_ = "pwd.txt"			'���������ļ�
CombLenth_ = 3		'��ϳ���
DeleteTempOrNot_ = 0		'�Ƿ�ɾ����;������1.txt��2.txt���ļ�


On Error Resume Next

Arr1_ = Readvar_file(SrcStr_)
For n = 1 To CombLenth_		'����1.txt��2.txt���ļ�
	If n = 1 Then
		sub_writefile Arr1_, "1.txt"
	Else
		sub_writefile PwdGen( Arr1_ , Readvar_file((n-1) & ".txt") ) , n & ".txt"
	End If
Next

sub_deletefile DstStr_
For n = 1 To CombLenth_		'�ϲ���һ���ļ���ɾ����ʱ�ļ�
	Mergevar_file n & ".txt" , DstStr_
	If DeleteTempOrNot_ Then
		sub_deletefile	n & ".txt"
	End If
Next

msgbox "���ɳɹ�����ǰ�����ϳ��ȣ�" & CombLenth_ 




'------------------------------------------------------------------

Function PwdGen(Arr1_(), Arr2_())		'�������ɺ�����������������ϳ�������
	k=0
	for i = 0 to UBound(Arr1_)
		for j = 0 to UBound(Arr2_)
			ReDim Preserve Arr3_(k)
			Arr3_(k) = Arr1_(i) & Arr2_(j)
			k = k + 1
		next
	next
	PwdGen = Arr3_
End Function

Function Readvar_file(var_filename)		'��ȡ�ļ�������
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 1)
	i = 0
	Do Until var_file.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = var_file.ReadLine
		i = i + 1
	Loop
	var_file.Close
 	Readvar_file = Arr_
End Function

Sub sub_writefile(Arr_(), var_filename)		'������д�뵽�ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 2, True)
	For i = 0 To UBound(Arr_)
		var_file.Write Arr_(i) & vbCrLf
	Next
End Sub

Sub Mergevar_file(Srcvar_filename, Dstvar_filename)		'�ϲ��ı��ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set FileSrc_ = var_fso.OpenTextFile(Srcvar_filename, 1, True, -2)
	Set Filevar_dst = var_fso.OpenTextFile(Dstvar_filename, 8, True)
	Do Until FileSrc_.AtEndOfStream
		Filevar_dst.WriteLine(FileSrc_.Readline)
	Loop
	FileSrc_.Close
	Filevar_dst.Close
End Sub

Sub sub_deletefile(var_filename)			'ɾ���ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.GetFile(var_filename)
	var_file.attributes = 0
	var_file.delete
	var_file.Close
End Sub

==============================

'��IP�������ָ�ʽ
'ע�⣺��Ҫ�пո�������ַ���
'------------------------------------------------------------------

SrcIpvar_file = "from.txt"	
DstIpvar_file = "to.txt"


ArrIpAll_ = Readvar_file(SrcIpvar_file)

Set var_fso = CreateObject("Scripting.FileSystemObject")
Set var_file = var_fso.OpenTextFile(DstIpvar_file, 2, True) '�ڶ�������8��ʾ׷��

For m = 0 to UBound(ArrIpAll_)
	var_file.WriteLine GenIpNuvar_m(ArrIpAll_(m))
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GenIpNuvar_m(StrIn_)
'����IP��ʽ���ɺ���
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNuvar_m = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3) 
End Function


Function Readvar_file(var_filename)		'��ȡ�ļ�������
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 1)
	i = 0
	Do Until var_file.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = var_file.ReadLine
		i = i + 1
	Loop
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
 	Readvar_file = Arr_
End Function


==============================

'������IPת��X.X.X.X��ʽ
'------------------------------------------------------------------
SrcIpvar_file = "from.txt"	
DstIpvar_file = "to.txt"


ArrIpAll_ = Readvar_file(SrcIpvar_file)

Set var_fso = CreateObject("Scripting.FileSystemObject")
Set var_file = var_fso.OpenTextFile(DstIpvar_file, 2, True) '�ڶ�������8��ʾ׷��

For m = 0 to UBound(ArrIpAll_)
	var_file.WriteLine GenStdIp_(ArrIpAll_(m))
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GenStdIp_(NumIn_)
'��׼IP��ʽ���ɺ���
	a = Int(NumIn_ / 16777216)	
	b = Int((NumIn_ - Int(NumIn_ / 16777216)*16777216) / 65536)
	'���ﲻ��mod����Ϊmod��������ֻ֧��long���ͣ�����������
	'���������\ͬ������ֻ����/������intȡ��
	c = Int((NumIn_ - Int(NumIn_ / 65536)*65536) / 256)
	d = Int(NumIn_ - Int(NumIn_ / 256)*256)
	GenStdIp_ = a & "." & b & "." & c & "." & d 
End Function


Function Readvar_file(var_filename)		'��ȡ�ļ�������
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 1)
	i = 0
	Do Until var_file.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = var_file.ReadLine
		i = i + 1
	Loop
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
 	Readvar_file = Arr_
End Function


==============================

IP��������

'�����θ�ʽ��IP���ɵ���IP�б�����nmap��-sL����
'���ǣ�nmap��֧��X.X.X.X-X.X.X.X���ָ�ʽ��ֻ֧��X.X.X.1-255����
'���ű���Ϊ���ֹ��ܵĲ���
'ע�⣺��Ҫ�пո�������ַ���������д���������û�пհ��л��ʽ����
'������������ʽ����ɾ��������û�������ַ�����
'\n\d+\.\d+\.\d+\.\d+-\d+\.\d+\.\d+\.\d+\n
'------------------------------------------------------------------

SrcIpvar_file = "from.txt"	
DstIpvar_file = "to.txt"


Createvar_file DstIpvar_file
ArrStrAll_ = Readvar_file(SrcIpvar_file)
for m = 0 to UBound(ArrStrAll_)
	sub_writefile GenIP_(ArrStrAll_(m)) , DstIpvar_file
next

Msgbox "Done!"

'------------------------------------------------------------------


Function GenIP_(StrIn_)
'����IP���ɺ��������Ϊ����
	IpMin_ = left(StrIn_,instr(StrIn_,"-")-1 )
	IpMax_ = right(StrIn_,len(StrIn_)-instr(StrIn_,"-"))
	i = 0
	for a = GenIpNuvar_m(IpMin_) to GenIpNuvar_m(IpMax_)
		redim preserve ArrTmp_(i)
		ArrTmp_(i) = GenStdIp_(a)
		i = i + 1
	next
	GenIP_ = ArrTmp_
	Set i = Nothing
	Set IpMin_ = Nothing
	Set IpMax_ = Nothing
End Function


Function GenIpNuvar_m(StrIn_)
'����IP��ʽ���ɺ���
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNuvar_m = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3)
End Function


Function GenStdIp_(NumIn_)
'��׼IP��ʽ���ɺ���
	a = Int(NumIn_ / 16777216)	
	b = Int((NumIn_ - Int(NumIn_ / 16777216)*16777216) / 65536)
	'���ﲻ��mod����Ϊmod��������ֻ֧��long���ͣ�����������
	'���������\ͬ������ֻ����/������intȡ��
	c = Int((NumIn_ - Int(NumIn_ / 65536)*65536) / 256)
	d = Int(NumIn_ - Int(NumIn_ / 256)*256)
	GenStdIp_ = a & "." & b & "." & c & "." & d 
End Function


Function Readvar_file(var_filename)		'��ȡ�ļ�������
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 1)
	i = 0
	Do Until var_file.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = var_file.ReadLine
		i = i + 1
	Loop
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
 	Readvar_file = Arr_
End Function

Sub sub_writefile(Arr_(), var_filename)		'������д�뵽�ļ�
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 8, True) '�ڶ�������8��ʾ׷��
	For i = 0 To UBound(Arr_)
		var_file.Write Arr_(i) & vbCrLf
	Next
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
End Sub

Sub Createvar_file(var_filename)		
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 2, True) '�ڶ�������2��ʾ����
	var_file.Write ""
	Set var_fso = Nothing
	Set var_file = Nothing	
End Sub

==============================

httpҳ����������
'ԭ���Ŀ�ģ�Ϊ��ע��һ��github��ϡȱ�û��������Ǿ��������ע����ˣ�����д���ű���������
'github���û���ҳ:https://github.com/XXX��Ȼ����FileLocator�ȹ����ҳ�404��Ӧ�����֡���ʵ������
'��ҳ����ļ���С�Ϳ��Կ������ˡ�
'�ű����к󣬻���ResultĿ¼������html�ļ���source.txt�ļ���Ҫ���ص�ҳ�档
'------------------------------------------------------------------

BaseUrl_ = "https://github.com/"
SrcStr_ = "Source.txt"		'������ȡ��Դ�ļ�
DstFolder_ = "Result"		'��������Ŀ¼


On Error Resume Next
sub_createfolder DstFolder_
Arr1_ = Readvar_file(SrcStr_)
For n = 0 to Ubound(Arr1_)
	HttpDwnLoader_ BaseUrl_ & Arr1_(n) , DstFolder_ & "\" & Arr1_(n) & ".html"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Sub HttpDwnLoader_(HttpUrl_,var_filename)		'����http���ļ�
	Set xaPost = CreateObject("MSXML2.ServerXMLHTTP")
	Set sGet = CreateObject("ADODB.Stream")
	sGet.Mode = 3
	sGet.Type = 1
	xaPost.Open "GET", HttpUrl_ , False
	xaPost.Send()
	sGet.Open()
	sGet.Write(xaPost.responseBody)
	sGet.SaveToFile var_filename, 2
	sGet.Close
End Sub

Function Readvar_file(var_filename)		'��ȡ�ļ�������
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 1)
	i = 0
	Do Until var_file.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = var_file.ReadLine
		i = i + 1
	Loop
	var_file.Close
	Set var_fso = Nothing
 	Readvar_file = Arr_
End Function

Sub sub_createfolder(var_filename)
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.CreateFolder(var_filename)
	Set var_fso = Nothing
End Sub


==============================

�����ļ�������

'��target�ļ�����ƥ��1.txt���ļ�������Ϊ2.txt

On Error Resume Next

t1 = "from.txt"
t2 = "to.txt"

s1=getarr(t1)
s2=getarr(t2)

for n=0 to UBound(s1)
	rname "target\" & s1(n) , s2(n)
next
msgbox "done!"


'read the text file gived line by line and set the value to a array, then return the array.
function getarr(fname)
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set fSeed = fso.OpenTextFile(fname,1)
	i = 0
	Do Until fSeed.AtEndOfStream
	redim preserve ArrTemp(i)
	ArrTemp(i) = fSeed.ReadLine 
	i=i+1
	Loop 
	fSeed.Close
	getarr=ArrTemp
end function

sub rname(od,nw)
	Set fso = CreateObject("Scripting.FileSystemObject") 
	set f=fso.getfile(od)
	f.name=nw
end sub

==============================

����exe

set ws = CreateObject("WScript.Shell")

a = wscript.ScriptFullName
b = left(a,instrrev(a,"\")-1)

ws.CurrentDirectory = b

c = b & "\AutoHotkey.exe"
c = chr(34) & c & chr(34)

d = b & "\script\main.ahk"
d = chr(34) & d & chr(34)

e = c & " " & d

ws.run e


==============================
telnet�����Ͱ���

set sh=WScript.CreateObject("WScript.Shell") 
sh.run "telnet 192.168.8.200"
WScript.Sleep 500
sh.SendKeys "admin~"
WScript.Sleep 500 
sh.SendKeys "admin~"
WScript.Sleep 500
sh.SendKeys "~"

==============================
==============================
==============================
==============================
==============================
==============================
==============================
==============================
==============================
==============================


