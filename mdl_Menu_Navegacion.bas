
Option Explicit

' ==================================================================================
' MÓDULO: mdl_Menu_Navegacion
' OBJETIVO:
'   Gestionar un menú principal que muestra UNA hoja concreta
'   ocultando todas las demás.
'
' USO:
'   Botones de menú → PASO 1 / PASO 2 / PASO 3
' ==================================================================================

Public Sub MostrarHoja(ByVal nombreHoja As String)

    Dim ws As Worksheet

    ' Ocultar todas las hojas del libro
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetHidden
    Next ws

    ' Comprobar que la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La hoja '" & nombreHoja & "' no existe.", vbCritical, "Error"
        Exit Sub
    End If

    ' Mostrar y activar la hoja seleccionada
    ws.Visible = xlSheetVisible
    ws.Activate

End Sub

Public Sub Menu_Paso_1()
    Call MostrarHoja("1")
End Sub

Public Sub Menu_Paso_2()
    Call MostrarHoja("2")
End Sub

Public Sub Menu_Paso_3()
    Call MostrarHoja("3")
End Sub
