do
Set wshshell = wscript.CreateObject("WScript.Shell")
Wshshell.run "Notepad"
wscript.sleep 100
wshshell.sendkeys "S"
wscript.sleep 100
wshshell.sendkeys "U"
wscript.sleep 100
wshshell.sendkeys "B"
wscript.sleep 100
wshshell.sendkeys "S"
wscript.sleep 100
wshshell.sendkeys "C"
wscript.sleep 100
wshshell.sendkeys "R"
wscript.sleep 100
wshshell.sendkeys "I"
wscript.sleep 100
wshshell.sendkeys "B"
wscript.sleep 100
wshshell.sendkeys "E"
wscript.sleep 100
wshshell.sendkeys "!"
loop