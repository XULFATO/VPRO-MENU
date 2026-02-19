Attribute VB_Name = "mdl_UI_Core"
Option Explicit

Public Sub MostrarMenuInicial()
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
    ' === OCULTAR TODO ===
    OcultarTodasLasHojas
    
    ' === CREAR / MOSTRAR MENU ===
    Set ws = ObtenerHojaMenu
    ws.Visible = xlSheetVisible
    ws.Activate
    
    PrepararAspectoApp ws
    LimpiarShapes ws
    CrearMenuPrincipal ws
    
    Application.ScreenUpdating = True
End Sub

Private Sub OcultarTodasLasHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVeryHidden
    Next ws
End Sub

Private Function ObtenerHojaMenu() As Worksheet
    On Error Resume Next
    Set ObtenerHojaMenu = ThisWorkbook.Worksheets("MENU")
    On Error GoTo 0
    
    If ObtenerHojaMenu Is Nothing Then
        Set ObtenerHojaMenu = ThisWorkbook.Worksheets.Add
        ObtenerHojaMenu.Name = "MENU"
    End If
End Function
