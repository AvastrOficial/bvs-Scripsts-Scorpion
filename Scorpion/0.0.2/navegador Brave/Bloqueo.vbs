' === bloqueo.vbs ===
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Crear carpeta Temp si no existe
If Not objFSO.FolderExists("C:\Temp") Then
    objFSO.CreateFolder("C:\Temp")
End If

' Ruta del archivo HTML temporal
htmlPath = "C:\Temp\pantalla_bloqueo.html"

' Obtener usuario de Windows
Set objNetwork = CreateObject("WScript.Network")
usuario = objNetwork.UserName

' Crear archivo HTML con la pantalla de bloqueo
Set objFile = objFSO.CreateTextFile(htmlPath, True)
objFile.WriteLine "<!DOCTYPE html>"
objFile.WriteLine "<html><head><meta charset='UTF-8'><title>Acceso Restringido</title>"
objFile.WriteLine "<style>"
objFile.WriteLine "body { background:#000; color:#f00; font-family:Consolas, monospace; text-align:center; margin-top:10%; overflow:hidden; }"
objFile.WriteLine "h1 { font-size:40px; margin-bottom:30px; text-shadow:0 0 15px #f00; }"
objFile.WriteLine "p { font-size:20px; margin-bottom:20px; }"
objFile.WriteLine "input { background:#111; border:1px solid #f00; color:#f00; font-size:20px; padding:10px; text-align:center; outline:none; }"
objFile.WriteLine "button { background:#111; border:1px solid #f00; color:#f00; font-size:20px; padding:10px 25px; cursor:pointer; margin-top:20px; }"
objFile.WriteLine "button:hover { background:#f00; color:#000; text-shadow:0 0 10px #000; }"
objFile.WriteLine "</style>"
objFile.WriteLine "<script>"
objFile.WriteLine "document.onkeydown = function(e) {"
objFile.WriteLine "   var k = e.keyCode;"
objFile.WriteLine "   if ((k >= 112 && k <= 123) || k==91 || k==92) { return false; }"  ' Bloquea F1–F12, WinKeys
objFile.WriteLine "   if (e.ctrlKey && (k==27 || k==9)) { return false; }"             ' Bloquea Ctrl+Esc, Ctrl+Tab
objFile.WriteLine "   if (e.altKey && k==9) { return false; }"                          ' Bloquea Alt+Tab
objFile.WriteLine "   return true;"
objFile.WriteLine "};"
objFile.WriteLine "</script></head>"
objFile.WriteLine "<body>"
objFile.WriteLine "<h1>Usuario: " & usuario & "</h1>"
objFile.WriteLine "<p>Introduce la contraseña para cerrar:</p>"
objFile.WriteLine "<input type='password' id='clave' /> <br>"
objFile.WriteLine "<button onclick=""if(document.getElementById('clave').value=='1234'){window.close();}else{alert('Contraseña incorrecta');}"">Acceder</button>"
objFile.WriteLine "</body></html>"
objFile.Close

' Ruta de Brave
bravePath = """C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"""

If Not objFSO.FileExists("C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe") Then
    MsgBox "Brave no encontrado en la ruta predeterminada.", 16, "Error"
    WScript.Quit
End If

' Ejecutar Brave en modo kiosco (sin barras ni botones)
objShell.Run bravePath & " --kiosk " & Chr(34) & htmlPath & Chr(34), 1, False
