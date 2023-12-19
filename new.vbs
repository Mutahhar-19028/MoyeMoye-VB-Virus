Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

' Play audio
oPlayer.URL = "moye.mp3"
oPlayer.controls.play
Set XPlayer = CreateObject("WScript.Shell")
XPlayer.Run "error.vbs"
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend

' Release the audio file
oPlayer.close