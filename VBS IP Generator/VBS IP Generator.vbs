'�����θ�ʽ��IP���ɵ���IP�б�����nmap��-sL����
'���ǣ�nmap��֧��X.X.X.X-X.X.X.X���ָ�ʽ��ֻ֧��X.X.X.1-255����
'���ű���Ϊ���ֹ��ܵĲ���
'ע�⣺��Ҫ�пո�������ַ�������ǰɾ��to.txt����������׷�ӵ�ip.txt
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
'����IP���ɺ��������Ϊ����
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