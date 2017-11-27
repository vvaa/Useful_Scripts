'把网段格式的IP生成单个IP列表，类似nmap的-sL功能
'但是，nmap不支持X.X.X.X-X.X.X.X这种格式，只支持X.X.X.1-255这种
'本脚本作为这种功能的补充
'注意：不要有空格等其它字符！运行前删除to.txt，否则结果会追加到ip.txt
'------------------------------------------------------------------

SrcIpFile_ = "from.txt"	
DstIpFile_ = "to.txt"


ArrStrAll_ = ReadFile_(SrcIpFile_)

for m = 0 to UBound(ArrStrAll_)
	WriteFile_ GenIP_(ArrStrAll_(m)) , DstIpFile_
next

Msgbox "Done!"

'------------------------------------------------------------------


Function GenIP_(StrIn_)
'单个IP生成函数，结果为数组
	IpMin_ = left(StrIn_,instr(StrIn_,"-")-1 )
	IpMax_= right(StrIn_,len(StrIn_)-instr(StrIn_,"-"))
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