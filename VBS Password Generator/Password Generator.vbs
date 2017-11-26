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