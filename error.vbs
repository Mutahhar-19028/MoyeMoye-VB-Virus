dim count
set object = wscript.CreateObject("wscript.shell")
do
main = msgbox("moye moye", 0+16, "Moye Moye")
wscript.Sleep(4000)
object.run "new.vbs"
count = count + 1
loop until count = 2
