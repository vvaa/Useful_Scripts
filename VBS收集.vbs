说明：收集一些自己常用的vbs函数，方便复用

==============================
开头
On Error Resume Next


==============================

打开文本文件

Set FSO_ = CreateObject("Scripting.FileSystemObject")
Set FILE_ = FSO_.OpenTextFile(Txt_) 
S_ = FILE_.ReadAll

==============================
写入文本文件

Set FILE_ = FSO_.CreateTextFile(Txt_)
FILE_.Write S_
FILE_.Close

==============================
替换文本

S_ = Replace(S_, Txt1_, Txt2_)

==============================
递增替换，增量替换，序号替换
第一个1表示从1开始，第二个1表示每次替换1个

For I_ = 1 to 100
	S_ = Replace(S_, Txt1_, I_, 1, 1)
Next

==============================
正则表达式
Set Re_ = New RegExp
Re_.Global = True
Re_.IgnoreCase = False
Re_.Pattern = "\d+\n"
For Each M_ in Re_.Execute(S_)
	Msgbox M_
Next

==============================
转义符，双引号

Msgbox """"

==============================
压缩文件成zip包
传入参数需完整路径


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

以下为改版为相对路径

Sub MyZip(Scr_, Dst_ )

	Path_ = createobject("Scripting.FileSystemObject").GetFolder(".").Path
	Scr_ = Path_ & "\" & Scr_
	Dst_ = Path_ & "\" & Dst_
	HisZip Scr_, Dst_
	
End Sub


==============================
获取当前路径
Path_ = Createobject("Scripting.FileSystemObject").GetFolder(".").Path
Msgbox Path_
==============================
删除文件
Sub DeleteFile_(Fname_)			'删除文件
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.GetFile(Fname_)
	FILE_.attributes = 0
	FILE_.delete
End Sub

==============================
创建文件夹
Sub CreateFolder_(Fname_)
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.CreateFolder(Fname_)
	Set FSO_ = Nothing
End Sub

==============================
将数组逐行写入到文件
Sub WriteFile_(ArrIn_(), Fname_)		'将数组写入到文件
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
逐行读取文本文件到数组
Function ReadFile_(Fname_)		'读取文件到数组
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
示例代码：分割文本文件为多个文件

'将文本文件分割，按行数分割！将文件分成行数相等的文件。
'------------------------------------------------------------------

SplitNum_ = 6		'要分割的数量
SrcFile_ = "Source.txt"		'要分割的源文件



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

Function ReadFile_(Fname_)		'读取文件到数组
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

Sub WriteFile_(ArrIn_(), Fname_)		'将数组写入到文件
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
合并文本文件

Sub MergeFile_(SrcFname_, DstFname_)		'合并文本文件
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
完整功能示例：密码生成器：
'一个简单的密码字典生成脚本，根据生日，名字，爱好等信息，排列组合成可能的密码。
'根据组合的定义长度，会生成多个中间文件，以1、2、3等命名，最后合并成一个文件。
'------------------------------------------------------------------

SrcStr_ = "Source.txt"		'用来读取的源文件
DstStr_ = "pwd.txt"			'保存结果的文件
CombLenth_ = 3		'组合长度
DeleteTempOrNot_ = 0		'是否删除中途产生的1.txt、2.txt等文件


On Error Resume Next

Arr1_ = ReadFile_(SrcStr_)
For n = 1 To CombLenth_		'生成1.txt、2.txt等文件
	If n = 1 Then
		WriteFile_ Arr1_, "1.txt"
	Else
		WriteFile_ PwdGen( Arr1_ , ReadFile_((n-1) & ".txt") ) , n & ".txt"
	End If
Next

DeleteFile_ DstStr_
For n = 1 To CombLenth_		'合并到一个文件并删除临时文件
	MergeFile_ n & ".txt" , DstStr_
	If DeleteTempOrNot_ Then
		DeleteFile_	n & ".txt"
	End If
Next

msgbox "生成成功！当前最大组合长度：" & CombLenth_ 




'------------------------------------------------------------------

Function PwdGen(Arr1_(), Arr2_())		'密码生成函数，将两个数组组合成新数组
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

Function ReadFile_(Fname_)		'读取文件到数组
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

Sub WriteFile_(Arr_(), Fname_)		'将数组写入到文件
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 2, True)
	For i = 0 To UBound(Arr_)
		FILE_.Write Arr_(i) & vbCrLf
	Next
End Sub

Sub MergeFile_(SrcFname_, DstFname_)		'合并文本文件
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FileSrc_ = FSO_.OpenTextFile(SrcFname_, 1, True, -2)
	Set FileDst_ = FSO_.OpenTextFile(DstFname_, 8, True)
	Do Until FileSrc_.AtEndOfStream
		FileDst_.WriteLine(FileSrc_.Readline)
	Loop
	FileSrc_.Close
	FileDst_.Close
End Sub

Sub DeleteFile_(Fname_)			'删除文件
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.GetFile(Fname_)
	FILE_.attributes = 0
	FILE_.delete
	FILE_.Close
End Sub

==============================

'把IP生成数字格式
'注意：不要有空格等其它字符！
'------------------------------------------------------------------

SrcIpFile_ = "from.txt"	
DstIpFile_ = "to.txt"


ArrIpAll_ = ReadFile_(SrcIpFile_)

Set FSO_ = CreateObject("Scripting.FileSystemObject")
Set FILE_ = FSO_.OpenTextFile(DstIpFile_, 2, True) '第二个参数8表示追加

For m = 0 to UBound(ArrIpAll_)
	FILE_.WriteLine GenIpNum_(ArrIpAll_(m))
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GenIpNum_(StrIn_)
'数字IP格式生成函数
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNum_ = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3) 
End Function


Function ReadFile_(Fname_)		'读取文件到数组
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

'把数字IP转成X.X.X.X格式
'------------------------------------------------------------------
SrcIpFile_ = "from.txt"	
DstIpFile_ = "to.txt"


ArrIpAll_ = ReadFile_(SrcIpFile_)

Set FSO_ = CreateObject("Scripting.FileSystemObject")
Set FILE_ = FSO_.OpenTextFile(DstIpFile_, 2, True) '第二个参数8表示追加

For m = 0 to UBound(ArrIpAll_)
	FILE_.WriteLine GenStdIp_(ArrIpAll_(m))
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GenStdIp_(NumIn_)
'标准IP格式生成函数
	a = Int(NumIn_ / 16777216)	
	b = Int((NumIn_ - Int(NumIn_ / 16777216)*16777216) / 65536)
	'这里不用mod是因为mod函数本身只支持long类型，会出现溢出，
	'整除运算符\同理，所以只能用/函数再int取整
	c = Int((NumIn_ - Int(NumIn_ / 65536)*65536) / 256)
	d = Int(NumIn_ - Int(NumIn_ / 256)*256)
	GenStdIp_ = a & "." & b & "." & c & "." & d 
End Function


Function ReadFile_(Fname_)		'读取文件到数组
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

IP生成器：

'把网段格式的IP生成单个IP列表，类似nmap的-sL功能
'但是，nmap不支持X.X.X.X-X.X.X.X这种格式，只支持X.X.X.1-255这种
'本脚本作为这种功能的补充
'注意：不要有空格等其它字符！如果运行错误检查下有没有空白行或格式不对
'用以下正则表达式搜索删除，看有没有其它字符留下
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
'单个IP生成函数，结果为数组
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
'数字IP格式生成函数
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNum_ = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3)
End Function


Function GenStdIp_(NumIn_)
'标准IP格式生成函数
	a = Int(NumIn_ / 16777216)	
	b = Int((NumIn_ - Int(NumIn_ / 16777216)*16777216) / 65536)
	'这里不用mod是因为mod函数本身只支持long类型，会出现溢出，
	'整除运算符\同理，所以只能用/函数再int取整
	c = Int((NumIn_ - Int(NumIn_ / 65536)*65536) / 256)
	d = Int(NumIn_ - Int(NumIn_ / 256)*256)
	GenStdIp_ = a & "." & b & "." & c & "." & d 
End Function


Function ReadFile_(Fname_)		'读取文件到数组
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

Sub WriteFile_(Arr_(), Fname_)		'将数组写入到文件
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 8, True) '第二个参数8表示追加
	For i = 0 To UBound(Arr_)
		FILE_.Write Arr_(i) & vbCrLf
	Next
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
End Sub

Sub CreateFile_(Fname_)		
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(Fname_, 2, True) '第二个参数2表示覆盖
	FILE_.Write ""
	Set FSO_ = Nothing
	Set FILE_ = Nothing	
End Sub

==============================

http页面下载器：
'原因和目的：为了注册一个github的稀缺用户名，但是绝大多数被注册掉了，所以写个脚本批量下载
'github的用户主页:https://github.com/XXX，然后用FileLocator等工具找出404对应的名字。其实不存在
'的页面从文件大小就可以看出来了。
'脚本运行后，会在Result目录下生成html文件，source.txt文件是要下载的页面。
'------------------------------------------------------------------

BaseUrl_ = "https://github.com/"
SrcStr_ = "Source.txt"		'用来读取的源文件
DstFolder_ = "Result"		'保存结果的目录


On Error Resume Next
CreateFolder_ DstFolder_
Arr1_ = ReadFile_(SrcStr_)
For n = 0 to Ubound(Arr1_)
	HttpDwnLoader_ BaseUrl_ & Arr1_(n) , DstFolder_ & "\" & Arr1_(n) & ".html"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Sub HttpDwnLoader_(HttpUrl_,Fname_)		'下载http到文件
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

Function ReadFile_(Fname_)		'读取文件到数组
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

批量文件重命名

'把target文件夹中匹配1.txt的文件重命名为2.txt

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

运行exe

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
telnet并发送按键

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


