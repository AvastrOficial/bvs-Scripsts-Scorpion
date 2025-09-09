Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Crear carpeta si no existe
Set objFolder = objFSO.GetFolder("C:\")
If Not objFSO.FolderExists("C:\Temp") Then
    objFSO.CreateFolder("C:\Temp")
End If

' Ruta del archivo HTML
htmlPath = "C:\Temp\pantalla_bloqueo.html"

' Contenido del HTML
htmlContent = "<!DOCTYPE html>" & vbCrLf & _
"<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"  <meta charset='UTF-8'>" & vbCrLf & _
"  <title>Acceso Restringido</title>" & vbCrLf & _
"  <style>" & vbCrLf & _
"    html, body {" & vbCrLf & _
"      margin: 0;" & vbCrLf & _
"      padding: 0;" & vbCrLf & _
"      background-color: #000;" & vbCrLf & _
"      color: #f00;" & vbCrLf & _
"      font-family: Consolas, monospace;" & vbCrLf & _
"      overflow: hidden;" & vbCrLf & _
"      user-select: none;" & vbCrLf & _
"    }" & vbCrLf & _
"    .centered {" & vbCrLf & _
"      position: absolute;" & vbCrLf & _
"      top: 50%;" & vbCrLf & _
"      left: 50%;" & vbCrLf & _
"      transform: translate(-50%, -50%);" & vbCrLf & _
"      text-align: center;" & vbCrLf & _
"    }" & vbCrLf & _
"    h1 {" & vbCrLf & _
"      font-size: 40px;" & vbCrLf & _
"      text-shadow: 0 0 15px #ff0000;" & vbCrLf & _
"    }" & vbCrLf & _
"    p {" & vbCrLf & _
"      font-size: 20px;" & vbCrLf & _
"      margin-bottom: 20px;" & vbCrLf & _
"    }" & vbCrLf & _
"    input {" & vbCrLf & _
"      background: #111;" & vbCrLf & _
"      border: 1px solid #ff0000;" & vbCrLf & _
"      color: #ff0000;" & vbCrLf & _
"      font-size: 20px;" & vbCrLf & _
"      padding: 10px;" & vbCrLf & _
"      text-align: center;" & vbCrLf & _
"      outline: none;" & vbCrLf & _
"    }" & vbCrLf & _
"    button {" & vbCrLf & _
"      background: #111;" & vbCrLf & _
"      border: 1px solid #ff0000;" & vbCrLf & _
"      color: #ff0000;" & vbCrLf & _
"      font-size: 20px;" & vbCrLf & _
"      padding: 10px 25px;" & vbCrLf & _
"      cursor: pointer;" & vbCrLf & _
"      margin-top: 20px;" & vbCrLf & _
"    }" & vbCrLf & _
"    button:hover {" & vbCrLf & _
"      background: #ff0000;" & vbCrLf & _
"      color: #000;" & vbCrLf & _
"    }" & vbCrLf & _
"  </style>" & vbCrLf & _
"  <script>" & vbCrLf & _
"    document.addEventListener('keydown', function (e) {" & vbCrLf & _
"      const k = e.keyCode;" & vbCrLf & _
"      if ((k >= 112 && k <= 123) || k === 27 || k === 91 || k === 92 || e.ctrlKey || e.altKey || e.metaKey) {" & vbCrLf & _
"        e.preventDefault();" & vbCrLf & _
"        return false;" & vbCrLf & _
"      }" & vbCrLf & _
"      const input = document.getElementById('clave');" & vbCrLf & _
"      if (e.target === input) {" & vbCrLf & _
"        if (!((k >= 48 && k <= 57) || (k >= 96 && k <= 105) || [8, 13].includes(k))) {" & vbCrLf & _
"          e.preventDefault();" & vbCrLf & _
"        }" & vbCrLf & _
"      }" & vbCrLf & _
"    });" & vbCrLf & _
"    document.addEventListener('contextmenu', function (e) {" & vbCrLf & _
"      e.preventDefault();" & vbCrLf & _
"    });" & vbCrLf & _
"    function cerrar() {" & vbCrLf & _
"      const clave = document.getElementById('clave').value;" & vbCrLf & _
"      if (clave === '1234') {" & vbCrLf & _
"        alert('Acceso concedido');" & vbCrLf & _
"        window.close();" & vbCrLf & _
"      } else {" & vbCrLf & _
"        alert('Contraseña incorrecta');" & vbCrLf & _
"      }" & vbCrLf & _
"    }" & vbCrLf & _
"  </script>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body>" & vbCrLf & _
"  <div class='centered'>" & vbCrLf & _
"    <h1>ACCESO RESTRINGIDO</h1>" & vbCrLf & _
"    <p>Introduce la contraseña numérica:</p>" & vbCrLf & _
"    <input type='password' id='clave' maxlength='8' autofocus />" & vbCrLf & _
"    <br />" & vbCrLf & _
"    <button onclick='cerrar()'>Acceder</button>" & vbCrLf & _
"  </div>" & vbCrLf & _
"</body>" & vbCrLf & _
"</html>"

' Crear y escribir el archivo HTML
Set objFile = objFSO.CreateTextFile(htmlPath, True)
objFile.Write htmlContent
objFile.Close

' Ruta de Chrome
chromePath = """C:\Program Files\Google\Chrome\Application\chrome.exe"""

' Parámetros para modo kiosko
args = "--kiosk """ & htmlPath & """"

' Ejecutar Chrome
objShell.Run chromePath & " " & args, 0, False
