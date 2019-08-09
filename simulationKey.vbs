Set cloner = CreateObject("WScript.Shell")
cloner.run"cmd"
'打開命令提示視窗
WScript.Sleep 500
'等500豪秒

cloner.SendKeys"telnet 192.168.1.1"  
'telnet 的 ip 和 port
cloner.SendKeys("{Enter}")
WScript.Sleep 500

cloner.SendKeys"UserName"
cloner.SendKeys("{Enter}")
WScript.Sleep 500

cloner.SendKeys"password"
cloner.SendKeys("{Enter}")
WScript.Sleep 500

cloner.SendKeys"sys reboot"
'Vigor Telnet 重開機的語法 form JanusLin
cloner.SendKeys("{Enter}")
WScript.Sleep 500


cloner.SendKeys"exit"
cloner.SendKeys("{Enter}") 