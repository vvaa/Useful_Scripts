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