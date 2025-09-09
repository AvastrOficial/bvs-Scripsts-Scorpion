' === archivo: esperar_desbloqueo.vbs ===
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

Do
    WScript.Sleep 1000
Loop Until fso.FileExists("C:\Temp\unlock.key")

' Cerrar Internet Explorer
shell.Run "taskkill /f /im iexplore.exe", 0, True

' Restaurar explorer.exe
shell.Run "explorer.exe", 0, False
