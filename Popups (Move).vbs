Option Explicit

' Frage anzeigen
Dim result
result = MsgBox("Do you really want to run this program? It will create many popups and may crash your computer.", vbYesNo + vbExclamation, "Warning!")

' Wenn der Benutzer "Yes" auswählt
If result = vbYes Then
    ' Shell-Objekt erstellen
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    
    ' Unendliche Schleife, die sich bewegende Pop-ups erstellt
    Do
        objShell.Run "mshta.exe ""about:<html><head><title>hacked</title><hta:application showintaskbar=""no""/><script>var x = 0;var y = 0;var dx = 20;var dy = 20;function moveWindow(){x += dx;y += dy;if(x + 100 > screen.width || x < 0){dx *= -1;}if(y + 100 > screen.height || y < 0){dy *= -1;}self.moveTo(x,y);setTimeout(moveWindow, 10);}moveWindow();</script></head><body bgcolor=""white""><center><font color=""RED"" size=""20""><b>YOYO</b></font></center></body></html>""", 1, False
    Loop
Else
    ' Wenn der Benutzer "No" auswählt
    MsgBox "The script has been terminated.", vbInformation, "Goodbye"
    WScript.Quit
End If
