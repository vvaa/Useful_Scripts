set ws = CreateObject("WScript.Shell")

a = wscript.ScriptFullName
b = left(a,instrrev(a,"\")-1)

ws.CurrentDirectory = b

c = b & "\AutoHotkey.exe"
c = chr(34) & c & chr(34)

d = b & "\script\main.ahk"
d = chr(34) & d & chr(34)

e = c & " " & d

ws.run e

