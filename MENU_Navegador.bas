Attribute VB_Name = "MENU_Navegador"
 
Option Explicit

' Macro principal para abrir pestañas
Public Sub AbrirPestañaProceso(ByVal nombreDestino As String, ByRef hojaOrigen As Worksheet)
    Dim wsNueva As Worksheet
    
    ' 1. Si la hoja ya existe, la borramos para que no se acumulen "OLD"
    Call BorrarHojaSiExiste(nombreDestino)
    
    ' 2. Creamos la hoja nueva
    Set wsNueva = ThisWorkbook.Worksheets.Add(After:=hojaOrigen)
    wsNueva.Name = nombreDestino
    wsNueva.Visible = xlSheetVisible
    
    ' 3. Si es una hoja de proceso (VCA_), dibujamos sus botones
    If InStr(nombreDestino, "VCA_") > 0 Then
        Call MENU_Logic.DibujarBotonesVCA(nombreDestino)
    End If
    
    wsNueva.Activate
End Sub

' Procedimiento para borrar sin que Excel pregunte ni deje rastro
Private Sub BorrarHojaSiExiste(ByVal nombre As String)
    Dim ws As Worksheet
    
    ' Intentamos asignar la hoja a la variable
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    
    ' Si la hoja existe realmente
    If Not ws Is Nothing Then
        ' 1. Comprobamos que no sea la única hoja del libro
        If ThisWorkbook.Worksheets.Count > 1 Then
            Application.DisplayAlerts = False
            
            ' 2. IMPORTANTE: La hacemos visible antes de borrar (Excel no deja borrar hojas VeryHidden directamente a veces)
            ws.Visible = xlSheetVisible
            
            ' 3. Intentamos borrar
            On Error Resume Next
            ws.Delete
            If Err.Number <> 0 Then
                MsgBox "No se pudo borrar la hoja '" & nombre & "'. Verifica si el libro está protegido.", vbCritical
            End If
            On Error GoTo 0
            
            Application.DisplayAlerts = True
        End If
    End If
End Sub
