set sh=WScript.CreateObject("WScript.Shell") 
sh.run "telnet 192.168.8.200"
WScript.Sleep 500
sh.SendKeys "admin~"
WScript.Sleep 500 
sh.SendKeys "admin~"
WScript.Sleep 500
sh.SendKeys "~"