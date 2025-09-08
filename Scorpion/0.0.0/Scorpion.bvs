
Set objNetwork = CreateObject("WScript.Network")
usuario = objNetwork.UserName


MsgBox "Bienvenido, " & usuario, vbInformation, "Detección de usuario"


Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.ToolBar = 0
objIE.StatusBar = 0
objIE.MenuBar = 0
objIE.FullScreen = True
objIE.Navigate("about:blank")

Do While objIE.Busy
    WScript.Sleep 100
Loop


html = "<!DOCTYPE html>" & _
       "<html>" & _
       "<head>" & _
       "<meta charset='UTF-8'>" & _
       "<title>Acceso Restringido</title>" & _
       "<style>" & _
       "body { background-color: #000000; color: #ff0000; font-family: Consolas, monospace; text-align:center; margin-top:10%; overflow:hidden; }" & _
       "h1 { font-size: 40px; margin-bottom: 30px; text-shadow: 0 0 15px #ff0000; }" & _
       "p { font-size: 20px; margin-bottom: 20px; }" & _
       "input { background:#111; border:1px solid #ff0000; color:#ff0000; font-size:20px; padding:10px; text-align:center; outline:none; }" & _
       "button { background:#111; border:1px solid #ff0000; color:#ff0000; font-size:20px; padding:10px 25px; cursor:pointer; margin-top:20px; }" & _
       "button:hover { background:#ff0000; color:#000; text-shadow:0 0 10px #000; }" & _
       "</style>" & _
       "<script>" & _
       "document.onkeydown = function(e) {" & _
       "   var k = e.keyCode;" & _
       "   if ((k >= 112 && k <= 123) || k==91 || k==92) { return false; }" & _  ' Bloquea F1–F12 y WinKeys
       "   if (e.ctrlKey && (k==27 || k==9)) { return false; }" & _              ' Bloquea Ctrl+Esc, Ctrl+Tab
       "   if (e.altKey && k==9) { return false; }" & _                         ' Bloquea Alt+Tab
       "   return true;" & _
       "};" & _
       "</script>" & _
       "</head>" & _
       "<body>" & _
       "<h1>Usuario: " & usuario & "</h1>" & _
       "<p>Introduce la contrase&ntilde;a para cerrar:</p>" & _
       "<input type='password' id='clave' />" & _
       "<br>" & _
       "<button onclick=""if(document.getElementById('clave').value=='1234'){window.close();}else{alert('Contraseña incorrecta');}"">Acceder</button>" & _
       "</body>" & _
       "</html>"

objIE.Document.Write html
objIE.Document.Close
