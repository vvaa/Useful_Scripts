'ʹ��������ʽ�������滻index���֣����ű�Ŀ�ģ�����enspÿ�����߶��ڲ�ͬ�ӿڡ�
'------------------------------------------------------------------

NumStart_ = 3
SrcFile_ = "a.txt"

'------------------------------------------------------------------


'On Error Resume Next




Set FSO_ = CreateObject("Scripting.FileSystemObject")
Set FILE_ = FSO_.OpenTextFile(SrcFile_) 
S_ = FILE_.ReadAll	

Set Re_ = New RegExp
Re_.Global = True
Re_.IgnoreCase = False


Re_.Pattern = "srcIndex=" & """" & "\d+"
N_ = NumStart_
For Each M_ in Re_.Execute(S_)
    S_ = Replace(S_, M_, "srcIndex=" & """" & "@@@@@")
    N_ = N_ + 1
Next

For I_ = NumStart_ to N_ -1
	S_ = Replace(S_, "srcIndex=" & """" & "@@@@@", "srcIndex=" & """" & I_,1,1)
Next




Re_.Pattern = "tarIndex=" & """" & "\d+"
N_ = NumStart_
For Each M_ in Re_.Execute(S_)
    S_ = Replace(S_, M_, "tarIndex=" & """" & "@@@@@")
    N_ = N_ + 1
Next


For I_ = NumStart_ to N_ -1
	S_ = Replace(S_, "tarIndex=" & """" & "@@@@@", "tarIndex=" & """" & I_,1,1)
Next





Set FILE_ = FSO_.CreateTextFile(SrcFile_)
FILE_.Write S_
FILE_.Close
Set FSO_ = Nothing
Set FILE_ = Nothing

Msgbox "Done!"






