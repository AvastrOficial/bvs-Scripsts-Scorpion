# Bvs Scripsts Scorpion - Pantalla de Bloqueo con VBS + IE

Este proyecto implementa una **pantalla de bloqueo estilo "kiosko"** usando `VBScript` e `Internet Explorer`.  
El escritorio de Windows queda inaccesible hasta que el usuario introduce la contraseña correcta.

---

## 📂 Archivos del proyecto

- **`pantalla_bloqueo.vbs`**  
  - Crea la carpeta temporal `C:\Temp`.  
  - Genera un archivo `pantalla_bloqueo.html` con la interfaz de bloqueo.  
  - Mata el proceso `explorer.exe` (para ocultar el escritorio).  
  - Lanza **Internet Explorer en modo kiosko (-k)** con el HTML.  
  - Inicia `esperar_desbloqueo.vbs` en segundo plano.

- **`esperar_desbloqueo.vbs`**  
  - Revisa cada segundo si existe el archivo `C:\Temp\unlock.key`.  
  - Cuando la contraseña correcta se ingresa en el HTML, el archivo se crea.  
  - Entonces:
    - Cierra `iexplore.exe` (Internet Explorer).  
    - Restaura `explorer.exe` → el escritorio vuelve a estar disponible.  

- **`start.bat`**  
  - Lanza automáticamente `pantalla_bloqueo.vbs` con `wscript`.  
  - Ejemplo:
    ```bat
    @echo off
    wscript "C:\Users\admin\Desktop\pantalla_bloqueo.vbs"
    ```

---

## 🔐 Flujo de funcionamiento

1. El usuario ejecuta `start.bat`.  
2. Se crea y abre la pantalla de bloqueo en **modo kiosko**.  
3. El usuario debe ingresar la contraseña correcta (`1234` por defecto).  
4. Si la clave es correcta:
   - El script genera `C:\Temp\unlock.key`.  
   - `esperar_desbloqueo.vbs` detecta el archivo.  
   - Se cierra IE y se restaura el escritorio.  

---

## ⚠️ Advertencias

- Este código **requiere Internet Explorer**, que ya no está disponible en Windows modernos (10/11).  
- Solo funciona en entornos donde **ActiveX y WScript están habilitados**.  
- Es un ejemplo educativo; **no debe usarse como malware** ni para bloquear equipos sin permiso.  

---

## 🚀 Personalización

- Cambiar la contraseña en `pantalla_bloqueo.vbs` → en la función JavaScript:
  ```js
  if(c==='1234'){
      // acción de desbloqueo
  }
 ```