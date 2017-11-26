'把target文件夹中匹配1.txt的文件重命名为2.txt

On Error Resume Next

t1 = "from.txt"
t2 = "to.txt"

s1=getarr(t1)
s2=getarr(t2)

for n=0 to UBound(s1)
	rname "target\" & s1(n) , s2(n)
next
msgbox "done!"


'read the text file gived line by line and set the value to a array, then return the array.
function getarr(fname)
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set fSeed = fso.OpenTextFile(fname,1)
	i = 0
	Do Until fSeed.AtEndOfStream
	redim preserve ArrTemp(i)
	ArrTemp(i) = fSeed.ReadLine 
	i=i+1
	Loop 
	fSeed.Close
	getarr=ArrTemp
end function

sub rname(od,nw)
	Set fso = CreateObject("Scripting.FileSystemObject") 
	set f=fso.getfile(od)
	f.name=nw
end sub