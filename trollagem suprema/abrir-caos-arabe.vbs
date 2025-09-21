Set shell = CreateObject("WScript.Shell")
For i = 1 To 200
    shell.Run "mshta.exe mensagem-arabe.hta", 1, False
    WScript.Sleep 50
Next