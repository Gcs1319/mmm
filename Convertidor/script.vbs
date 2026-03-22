' Ajuste de tamaño al iniciar
Sub Window_OnLoad
    window.resizeTo 800, 900
End Sub

' Seleccion de carpeta
Sub SeleccionarCarpeta()
    Dim objShell, objFolder
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Selecciona carpeta:", &H1, 17)
    
    If Not objFolder Is Nothing Then 
        document.getElementById("folderInput").value = objFolder.Self.Path
    End If
End Sub

' Descarga usando yt-dlp
Sub Descargar()
    Dim objWShell, url, folder, formato, comando, calidad, resultado
    
    url = document.getElementById("urlInput").value
    folder = document.getElementById("folderInput").value
    formato = document.getElementById("formatSelect").value
    
    If Trim(url) = "" Or Trim(folder) = "" Then
        MsgBox "Error: Falta el enlace o la carpeta.", 48, "Atencion"
        Exit Sub
    End If
    
    If formato = "mp3" Then
        calidad = "-x --audio-format mp3 --audio-quality 0"
    Else
        calidad = "-f ""bv*[ext=mp4]+ba[ext=m4a]/b[ext=mp4]"""
    End If
    
    Set objWShell = CreateObject("WScript.Shell")
    
    comando = "cmd /c yt-dlp " & calidad & " --restrict-filenames -o """ & folder & "/%(title)s.%(ext)s"" " & url
    
    document.getElementById("status").innerHTML = "DESCARGANDO..."
    document.getElementById("status").style.color = "#0078d7"
    document.getElementById("btnMain").disabled = True
    
    resultado = objWShell.Run(comando, 0, True)
    
    If resultado = 0 Then
        document.getElementById("status").innerHTML = "DESCARGA FINALIZADA"
        document.getElementById("status").style.color = "#28a745"
        MsgBox "Descarga completada correctamente. Revisa tu carpeta.", 64, "Finalizado"
    Else
        document.getElementById("status").innerHTML = "Error en la descarga."
        document.getElementById("status").style.color = "#ff0000"
        MsgBox "Ha ocurrido un error durante la descarga.", 48, "Error"
    End If
    
    document.getElementById("btnMain").disabled = False
End Sub