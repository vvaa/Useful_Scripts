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
