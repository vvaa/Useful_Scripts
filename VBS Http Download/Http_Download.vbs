'ԭ���Ŀ�ģ�Ϊ��ע��һ��github��ϡȱ�û��������Ǿ��������ע����ˣ�����д���ű���������
'github���û���ҳ:https://github.com/XXX��Ȼ����FileLocator�ȹ����ҳ�404��Ӧ�����֡���ʵ������
'��ҳ����ļ���С�Ϳ��Կ������ˡ�
'�ű����к󣬻���ResultĿ¼������html�ļ���source.txt�ļ���Ҫ���ص�ҳ�档
'------------------------------------------------------------------

BaseUrl_ = "https://github.com/"
SrcStr_ = "Source.txt"		'������ȡ��Դ�ļ�
DstFolder_ = "Result"		'��������Ŀ¼


On Error Resume Next
CreateFolder_ DstFolder_
Arr1_ = ReadFile_(SrcStr_)
For n = 0 to Ubound(Arr1_)
	HttpDwnLoader_ BaseUrl_ & Arr1_(n) , DstFolder_ & "\" & Arr1_(n) & ".html"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Sub HttpDwnLoader_(HttpUrl_,Fname_)		'����http���ļ�
	Set xaPost = CreateObject("MSXML2.ServerXMLHTTP")
	Set sGet = CreateObject("ADODB.Stream")
	sGet.Mode = 3
	sGet.Type = 1
	xaPost.Open "GET", HttpUrl_ , False
	xaPost.Send()
	sGet.Open()
	sGet.Write(xaPost.responseBody)
	sGet.SaveToFile Fname_, 2
	sGet.Close
End Sub

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
 	ReadFile_ = Arr_
End Function

Sub CreateFolder_(Fname_)
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.CreateFolder(Fname_)
	Set FSO_ = Nothing
End Sub
