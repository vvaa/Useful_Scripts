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

