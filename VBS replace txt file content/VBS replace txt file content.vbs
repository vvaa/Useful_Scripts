'替换文本文件内容，保存到新的文件
'------------------------------------------------------------------

SrcTxt_ = "111"
DstTxt_ = "@@@@@"
SrcFile_ = "Old.txt"
DstFile_ = "New.txt"

'------------------------------------------------------------------


'On Error Resume Next

Replace_ SrcTxt_, DstTxt_, SrcFile_, DstFile_

Msgbox "Done!"



'------------------------------------------------------------------


Sub Replace_(SrcTxt_, DstTxt_, SrcFile_, DstFile_)

	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.OpenTextFile(SrcFile_) 
	S_ = FILE_.ReadAll	
	
	S_=Replace(S_,SrcTxt_,DstTxt_)
	Set FILE_ = FSO_.CreateTextFile(DstFile_)
	FILE_.Write S_
	
	FILE_.Close
	Set FSO_ = Nothing
	Set FILE_ = Nothing
	
End Sub

