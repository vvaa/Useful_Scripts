'原因和目的：为了注册一个github的稀缺用户名，但是绝大多数被注册掉了，所以写个脚本批量下载
'github的用户主页:https://github.com/XXX，然后用FileLocator等工具找出404对应的名字。其实不存在
'的页面从文件大小就可以看出来了。
'脚本运行后，会在Result目录下生成html文件，source.txt文件是要下载的页面。
'------------------------------------------------------------------

BaseUrl_ = "https://github.com/"
SrcStr_ = "Source.txt"		'用来读取的源文件
DstFolder_ = "Result"		'保存结果的目录


On Error Resume Next
CreateFolder_ DstFolder_
Arr1_ = ReadFile_(SrcStr_)
For n = 0 to Ubound(Arr1_)
	HttpDwnLoader_ BaseUrl_ & Arr1_(n) , DstFolder_ & "\" & Arr1_(n) & ".html"
Next

Msgbox "Done!"

'------------------------------------------------------------------

Sub HttpDwnLoader_(HttpUrl_,Fname_)		'下载http到文件
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
 	ReadFile_ = Arr_
End Function

Sub CreateFolder_(Fname_)
	Set FSO_ = CreateObject("Scripting.FileSystemObject")
	Set FILE_ = FSO_.CreateFolder(Fname_)
	Set FSO_ = Nothing
End Sub
