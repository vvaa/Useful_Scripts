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