˵�����ռ�һЩ�Լ����õ�vbs���������㸴��

==============================
��ͷ
On Error Resume Next


==============================

���ı��ļ�

Set FSO_ = CreateObject("Scripting.FileSystemObject")
Set FILE_ = FSO_.OpenTextFile(Txt_) 
S_ = FILE_.ReadAll

==============================
д���ı��ļ�

Set FILE_ = FSO_.CreateTextFile(Txt_)
FILE_.Write S_
FILE_.Close

==============================
�滻�ı�

S_ = Replace(S_, Txt1_, Txt2_)

==============================
�����滻�������滻������滻
��һ��1��ʾ��1��ʼ���ڶ���1��ʾÿ���滻1��

For I_ = 1 to 100
	S_ = Replace(S_, Txt1_, I_, 1, 1)
Next

==============================
������ʽ
Set Re_ = New RegExp
Re_.Global = True
Re_.IgnoreCase = False
Re_.Pattern = "\d+\n"
For Each M_ in Re_.Execute(S_)
	Msgbox M_
Next

==============================
ת�����˫����

Msgbox """"

==============================
ѹ���ļ���zip��
�������������·��


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

����Ϊ�İ�Ϊ���·��

Sub MyZip(Scr_, Dst_ )

	Path_ = createobject("Scripting.FileSystemObject").GetFolder(".").Path
	Scr_ = Path_ & "\" & Scr_
	Dst_ = Path_ & "\" & Dst_
	HisZip Scr_, Dst_
	
End Sub


==============================
��ȡ��ǰ·��
Path_ = Createobject("Scripting.FileSystemObject").GetFolder(".").Path
Msgbox Path_
==============================
ɾ���ļ�
Sub DeleteFile_(Fname_)			'ɾ���ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.GetFile(Fname_)
	FILE_.attributes = 0
	FILE_.delete
End Sub

==============================
�����ļ���
Sub CreateFolder_(Fname_)
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.CreateFolder(Fname_)
	Set FSO_ = Nothing
End Sub

==============================
����������д�뵽�ļ�
Sub WriteFile_(ArrIn_(), Fname_)		'������д�뵽�ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 2, True) 
	For i = 0 To UBound(ArrIn_)
		FILE_.WriteLine ArrIn_(i)
	Next
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
End Sub

==============================
���ж�ȡ�ı��ļ�������
Function ReadFile_(Fname_)		'��ȡ�ļ�������
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 1)
	i = 0
	Do Until FILE_.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = FILE_.ReadLine
		i = i + 1
	Loop
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
 	ReadFile_ = Arr_
End Function
==============================
ʾ�����룺�ָ��ı��ļ�Ϊ����ļ�

'���ı��ļ��ָ�������ָ���ļ��ֳ�������ȵ��ļ���
'------------------------------------------------------------------

SplitNum_ = 6		'Ҫ�ָ������
SrcFile_ = "Source.txt"		'Ҫ�ָ��Դ�ļ�



'On Error Resume Next
Arr1_ = ReadFile_(SrcFile_)
For n = 1 to SplitNum_
	WriteFile_ GetPartArr_(Arr1_,SplitNum_,n) , n & ".txt"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GetPartArr_(ArrIn_(),NumAll_,NumPart_)	
	BlockSize_ = (Ubound(ArrIn_)+1) \ NumAll_
	If NumAll_ = NumPart_ Then
		j = 0
		For k = BlockSize_ * (NumPart_-1)  to Ubound(ArrIn_)
			redim preserve ArrTmp_(j)
			ArrTmp_(j) = ArrIn_(k)
			j = j + 1
		Next
	Else
		j = 0
		For k = BlockSize_ * (NumPart_-1) to BlockSize_ * NumPart_
			redim preserve ArrTmp_(j)
			ArrTmp_(j) = ArrIn_(k)
			j = j + 1
		Next
	End If
 	GetPartArr_ = ArrTmp_
End Function

Function ReadFile_(Fname_)		'��ȡ�ļ�������
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 1)
	i = 0
	Do Until FILE_.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = FILE_.ReadLine
		i = i + 1
	Loop
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
 	ReadFile_ = Arr_
End Function

Sub WriteFile_(ArrIn_(), Fname_)		'������д�뵽�ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 2, True) 
	For i = 0 To UBound(ArrIn_)
		FILE_.WriteLine ArrIn_(i)
	Next
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
End Sub



==============================
�ϲ��ı��ļ�

Sub MergeFile_(SrcFname_, DstFname_)		'�ϲ��ı��ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FileSrc_ = FSO_.OpenTextFile(SrcFname_, 1, True, -2)
	Set FileDst_ = FSO_.OpenTextFile(DstFname_, 8, True)
	Do Until FileSrc_.AtEndOfStream
		FileDst_.WriteLine(FileSrc_.Readline)
	Loop
	FileSrc_.Close
	FileDst_.Close
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

Arr1_ = ReadFile_(SrcStr_)
For n = 1 To CombLenth_		'����1.txt��2.txt���ļ�
	If n = 1 Then
		WriteFile_ Arr1_, "1.txt"
	Else
		WriteFile_ PwdGen( Arr1_ , ReadFile_((n-1) & ".txt") ) , n & ".txt"
	End If
Next

DeleteFile_ DstStr_
For n = 1 To CombLenth_		'�ϲ���һ���ļ���ɾ����ʱ�ļ�
	MergeFile_ n & ".txt" , DstStr_
	If DeleteTempOrNot_ Then
		DeleteFile_	n & ".txt"
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

Function ReadFile_(Fname_)		'��ȡ�ļ�������
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 1)
	i = 0
	Do Until FILE_.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = FILE_.ReadLine
		i = i + 1
	Loop
	FILE_.Close
 	ReadFile_ = Arr_
End Function

Sub WriteFile_(Arr_(), Fname_)		'������д�뵽�ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 2, True)
	For i = 0 To UBound(Arr_)
		FILE_.Write Arr_(i) & vbCrLf
	Next
End Sub

Sub MergeFile_(SrcFname_, DstFname_)		'�ϲ��ı��ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FileSrc_ = FSO_.OpenTextFile(SrcFname_, 1, True, -2)
	Set FileDst_ = FSO_.OpenTextFile(DstFname_, 8, True)
	Do Until FileSrc_.AtEndOfStream
		FileDst_.WriteLine(FileSrc_.Readline)
	Loop
	FileSrc_.Close
	FileDst_.Close
End Sub

Sub DeleteFile_(Fname_)			'ɾ���ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.GetFile(Fname_)
	FILE_.attributes = 0
	FILE_.delete
	FILE_.Close
End Sub

==============================

'��IP�������ָ�ʽ
'ע�⣺��Ҫ�пո�������ַ���
'------------------------------------------------------------------

SrcIpFile_ = "from.txt"	
DstIpFile_ = "to.txt"


ArrIpAll_ = ReadFile_(SrcIpFile_)

Set FSO_ = CreateObject("Scripting.FileSystemObject")
Set FILE_ = FSO_.OpenTextFile(DstIpFile_, 2, True) '�ڶ�������8��ʾ׷��

For m = 0 to UBound(ArrIpAll_)
	FILE_.WriteLine GenIpNum_(ArrIpAll_(m))
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GenIpNum_(StrIn_)
'����IP��ʽ���ɺ���
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNum_ = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3) 
End Function


Function ReadFile_(Fname_)		'��ȡ�ļ�������
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 1)
	i = 0
	Do Until FILE_.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = FILE_.ReadLine
		i = i + 1
	Loop
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
 	ReadFile_ = Arr_
End Function


==============================

'������IPת��X.X.X.X��ʽ
'------------------------------------------------------------------
SrcIpFile_ = "from.txt"	
DstIpFile_ = "to.txt"


ArrIpAll_ = ReadFile_(SrcIpFile_)

Set FSO_ = CreateObject("Scripting.FileSystemObject")
Set FILE_ = FSO_.OpenTextFile(DstIpFile_, 2, True) '�ڶ�������8��ʾ׷��

For m = 0 to UBound(ArrIpAll_)
	FILE_.WriteLine GenStdIp_(ArrIpAll_(m))
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


Function ReadFile_(Fname_)		'��ȡ�ļ�������
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 1)
	i = 0
	Do Until FILE_.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = FILE_.ReadLine
		i = i + 1
	Loop
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
 	ReadFile_ = Arr_
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

SrcIpFile_ = "from.txt"	
DstIpFile_ = "to.txt"


CreateFile_ DstIpFile_
ArrStrAll_ = ReadFile_(SrcIpFile_)
for m = 0 to UBound(ArrStrAll_)
	WriteFile_ GenIP_(ArrStrAll_(m)) , DstIpFile_
next

Msgbox "Done!"

'------------------------------------------------------------------


Function GenIP_(StrIn_)
'����IP���ɺ��������Ϊ����
	IpMin_ = left(StrIn_,instr(StrIn_,"-")-1 )
	IpMax_ = right(StrIn_,len(StrIn_)-instr(StrIn_,"-"))
	i = 0
	for a = GenIpNum_(IpMin_) to GenIpNum_(IpMax_)
		redim preserve ArrTmp_(i)
		ArrTmp_(i) = GenStdIp_(a)
		i = i + 1
	next
	GenIP_ = ArrTmp_
	Set i = Nothing
	Set IpMin_ = Nothing
	Set IpMax_ = Nothing
End Function


Function GenIpNum_(StrIn_)
'����IP��ʽ���ɺ���
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNum_ = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3)
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


Function ReadFile_(Fname_)		'��ȡ�ļ�������
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 1)
	i = 0
	Do Until FILE_.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = FILE_.ReadLine
		i = i + 1
	Loop
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
 	ReadFile_ = Arr_
End Function

Sub WriteFile_(Arr_(), Fname_)		'������д�뵽�ļ�
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 8, True) '�ڶ�������8��ʾ׷��
	For i = 0 To UBound(Arr_)
		FILE_.Write Arr_(i) & vbCrLf
	Next
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
End Sub

Sub CreateFile_(Fname_)		
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 2, True) '�ڶ�������2��ʾ����
	FILE_.Write ""
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
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
CreateFolder_ DstFolder_
Arr1_ = ReadFile_(SrcStr_)
For n = 0 to Ubound(Arr1_)
	HttpDwnLoader_ BaseUrl_ & Arr1_(n) , DstFolder_ & "\" & Arr1_(n) & ".html"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Sub HttpDwnLoader_(HttpUrl_,Fname_)		'����http���ļ�
	Set xaPost = CreateObject("MSXML2.ServerXMLHTTP")
	Set sGet = CreateObject("ADODB.Stream")
	sGet.Mode = 3
	sGet.Type = 1
	xaPost.Open "GET", HttpUrl_ , False
	xaPost.Send()
	sGet.Open()
	sGet.Write(xaPost.responseBody)
	sGet.SaveToFile Fname_, 2
	sGet.Close
End Sub

Function ReadFile_(Fname_)		'��ȡ�ļ�������
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 1)
	i = 0
	Do Until FILE_.AtEndOfStream
		redim preserve Arr_(i)
		Arr_(i) = FILE_.ReadLine
		i = i + 1
	Loop
	FILE_.Close
	Set FSO_ = Nothing
 	ReadFile_ = Arr_
End Function

Sub CreateFolder_(Fname_)
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.CreateFolder(Fname_)
	Set FSO_ = Nothing
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


