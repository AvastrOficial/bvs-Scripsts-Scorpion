' === archivo: esperar_desbloqueo.vbs ===
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set WMI = GetObject("winmgmts:\\.\root\cimv2")

Do
    WScript.Sleep 1000

    ' Verificar si explorer.exe está corriendo, si sí, matarlo
    Set procs = WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'")
    For Each p In procs
        p.Terminate
    Next

Loop Until fso.FileExists("C:\Temp\unlock.key")

' Cerrar Google Chrome (modo kiosko)
shell.Run "taskkill /f /im chrome.exe", 0, True

' Restaurar explorer.exe
shell.Run "explorer.exe", 0, False
