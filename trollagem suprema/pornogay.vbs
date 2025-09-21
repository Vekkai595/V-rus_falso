Set shell = CreateObject("WScript.Shell")

For i = 1 To 1000
    shell.Popup "SEU PC ESTÁ SENDO BRUTALMENTE DANIFICADO", 1, "ALERTA CRÍTICO", 16 + 4096
Next

WScript.Sleep 10000 ' Espera 10 segundos

shell.Run "shutdown -s -t 60 -c ""Seu PC vai desligar em 1 minuto..."""