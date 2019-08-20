说明：收集一些自己常用的vbs函数，方便复用
为了不弄混，约定：
过程用sub_开头，
函数用fun_开通
变量用var_开头

==============================
开头
On Error Resume Next


==============================

打开文本文件

Set var_fso = CreateObject("Scripting.FileSystemObject")
Set var_file = var_fso.OpenTextFile(var_txt) 
var_s = var_file.ReadAll

==============================
写入文本文件

Set var_file = var_fso.CreateTextFile(var_txt)
var_file.Write var_s
var_file.Close

==============================
替换文本

var_s = Replace(var_s, var_txt1, var_txt2)

==============================
递增替换，增量替换，序号替换
第一个1表示从1开始，第二个1表示每次替换1个

For var_i = 1 to 100
	var_s = Replace(var_s, var_txt1, var_i, 1, 1)
Next

==============================
正则表达式
Set var_re = New RegExp
var_re.Global = True
var_re.IgnoreCase = False
var_re.Pattern = "\d+\n"
For Each var_m in var_re.Execute(var_s)
	Msgbox var_m
Next

==============================
转义符，双引号

Msgbox """"

==============================
压缩文件成zip包
传入参数需完整路径


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

以下为改版为相对路径

Sub sub_myzip(var_src, var_dst )

	var_path = createobject("Scripting.FileSystemObject").GetFolder(".").Path
	var_src = var_path & "\" & var_src
	var_dst = var_path & "\" & var_dst
	sub_hiszip var_src, var_dst
	
End Sub


==============================
获取当前路径
var_path = Createobject("Scripting.FileSystemObject").GetFolder(".").Path
Msgbox var_path


==============================
删除文件
Sub sub_deletefile(var_filename)			'删除文件
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.GetFile(var_filename)
	var_file.attributes = 0
	var_file.delete
End Sub

==============================
创建文件夹
Sub sub_createfolder(var_filename)
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.CreateFolder(var_filename)
	Set var_fso = Nothing
End Sub

==============================
将数组逐行写入到文件
Sub sub_writefile(var_arrin(), var_filename)		'将数组写入到文件
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
逐行读取文本文件到数组
Function Readvar_file(var_filename)		'读取文件到数组
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
示例代码：分割文本文件为多个文件

'将文本文件分割，按行数分割！将文件分成行数相等的文件。
'------------------------------------------------------------------

SplitNuvar_m = 6		'要分割的数量
Srcvar_file = "Source.txt"		'要分割的源文件



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

Function Readvar_file(var_filename)		'读取文件到数组
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

Sub sub_writefile(var_arrin(), var_filename)		'将数组写入到文件
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
合并文本文件

Sub Mergevar_file(Srcvar_filename, Dstvar_filename)		'合并文本文件
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
完整功能示例：密码生成器：
'一个简单的密码字典生成脚本，根据生日，名字，爱好等信息，排列组合成可能的密码。
'根据组合的定义长度，会生成多个中间文件，以1、2、3等命名，最后合并成一个文件。
'------------------------------------------------------------------

SrcStr_ = "Source.txt"		'用来读取的源文件
DstStr_ = "pwd.txt"			'保存结果的文件
CombLenth_ = 3		'组合长度
DeleteTempOrNot_ = 0		'是否删除中途产生的1.txt、2.txt等文件


On Error Resume Next

Arr1_ = Readvar_file(SrcStr_)
For n = 1 To CombLenth_		'生成1.txt、2.txt等文件
	If n = 1 Then
		sub_writefile Arr1_, "1.txt"
	Else
		sub_writefile PwdGen( Arr1_ , Readvar_file((n-1) & ".txt") ) , n & ".txt"
	End If
Next

sub_deletefile DstStr_
For n = 1 To CombLenth_		'合并到一个文件并删除临时文件
	Mergevar_file n & ".txt" , DstStr_
	If DeleteTempOrNot_ Then
		sub_deletefile	n & ".txt"
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

Function Readvar_file(var_filename)		'读取文件到数组
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

Sub sub_writefile(Arr_(), var_filename)		'将数组写入到文件
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 2, True)
	For i = 0 To UBound(Arr_)
		var_file.Write Arr_(i) & vbCrLf
	Next
End Sub

Sub Mergevar_file(Srcvar_filename, Dstvar_filename)		'合并文本文件
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set FileSrc_ = var_fso.OpenTextFile(Srcvar_filename, 1, True, -2)
	Set Filevar_dst = var_fso.OpenTextFile(Dstvar_filename, 8, True)
	Do Until FileSrc_.AtEndOfStream
		Filevar_dst.WriteLine(FileSrc_.Readline)
	Loop
	FileSrc_.Close
	Filevar_dst.Close
End Sub

Sub sub_deletefile(var_filename)			'删除文件
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.GetFile(var_filename)
	var_file.attributes = 0
	var_file.delete
	var_file.Close
End Sub

==============================

'把IP生成数字格式
'注意：不要有空格等其它字符！
'------------------------------------------------------------------

SrcIpvar_file = "from.txt"	
DstIpvar_file = "to.txt"


ArrIpAll_ = Readvar_file(SrcIpvar_file)

Set var_fso = CreateObject("Scripting.FileSystemObject")
Set var_file = var_fso.OpenTextFile(DstIpvar_file, 2, True) '第二个参数8表示追加

For m = 0 to UBound(ArrIpAll_)
	var_file.WriteLine GenIpNuvar_m(ArrIpAll_(m))
Next

Msgbox "Done!"

'------------------------------------------------------------------

Function GenIpNuvar_m(StrIn_)
'数字IP格式生成函数
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNuvar_m = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3) 
End Function


Function Readvar_file(var_filename)		'读取文件到数组
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

'把数字IP转成X.X.X.X格式
'------------------------------------------------------------------
SrcIpvar_file = "from.txt"	
DstIpvar_file = "to.txt"


ArrIpAll_ = Readvar_file(SrcIpvar_file)

Set var_fso = CreateObject("Scripting.FileSystemObject")
Set var_file = var_fso.OpenTextFile(DstIpvar_file, 2, True) '第二个参数8表示追加

For m = 0 to UBound(ArrIpAll_)
	var_file.WriteLine GenStdIp_(ArrIpAll_(m))
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


Function Readvar_file(var_filename)		'读取文件到数组
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

IP生成器：

'把网段格式的IP生成单个IP列表，类似nmap的-sL功能
'但是，nmap不支持X.X.X.X-X.X.X.X这种格式，只支持X.X.X.1-255这种
'本脚本作为这种功能的补充
'注意：不要有空格等其它字符！如果运行错误检查下有没有空白行或格式不对
'用以下正则表达式搜索删除，看有没有其它字符留下
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
'单个IP生成函数，结果为数组
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
'数字IP格式生成函数
	ArrTmp_ = split(StrIn_, ".") 
	Redim Preserve ArrTmp_(3)
	GenIpNuvar_m = ArrTmp_(0) * 16777216 + ArrTmp_(1) * 65536 + ArrTmp_(2) * 256 + ArrTmp_(3)
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


Function Readvar_file(var_filename)		'读取文件到数组
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

Sub sub_writefile(Arr_(), var_filename)		'将数组写入到文件
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 8, True) '第二个参数8表示追加
	For i = 0 To UBound(Arr_)
		var_file.Write Arr_(i) & vbCrLf
	Next
	var_file.Close
	Set var_fso = Nothing
	Set var_file = Nothing	
End Sub

Sub Createvar_file(var_filename)		
	Set var_fso = CreateObject("Scripting.FileSystemObject")
	Set var_file = var_fso.OpenTextFile(var_filename, 2, True) '第二个参数2表示覆盖
	var_file.Write ""
	Set var_fso = Nothing
	Set var_file = Nothing	
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
sub_createfolder DstFolder_
Arr1_ = Readvar_file(SrcStr_)
For n = 0 to Ubound(Arr1_)
	HttpDwnLoader_ BaseUrl_ & Arr1_(n) , DstFolder_ & "\" & Arr1_(n) & ".html"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Sub HttpDwnLoader_(HttpUrl_,var_filename)		'下载http到文件
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

Function Readvar_file(var_filename)		'读取文件到数组
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


