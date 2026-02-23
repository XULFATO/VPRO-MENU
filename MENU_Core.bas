Attribute VB_Name = "MENU_Core"

Option Explicit

Public Sub MostrarMenuInicial()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    ActiveWindow.DisplayWorkbookTabs = True
    
    Call OcultarTodasLasHojas("MENU")
    
    Set ws = ObtenerHojaMenu
    ws.Visible = xlSheetVisible
    ws.Activate
    
    Call MENU_Diseno.PrepararAspectoApp(ws)
    Call MENU_Diseno.LimpiarShapes(ws)
    Call MENU_Diseno.CrearMenuPrincipal(ws)
    
    Application.ScreenUpdating = True
End Sub

Public Sub OcultarTodasLasHojas(Optional ByVal nombreHoja As String = "MENU")
    Dim ws As Worksheet
    On Error Resume Next
    ThisWorkbook.Worksheets(nombreHoja).Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> nombreHoja Then ws.Visible = xlSheetVeryHidden
    Next ws
    On Error GoTo 0
End Sub

Private Function ObtenerHojaMenu() As Worksheet
    On Error Resume Next
    Set ObtenerHojaMenu = ThisWorkbook.Worksheets("MENU")
    On Error GoTo 0
    If ObtenerHojaMenu Is Nothing Then
        Set ObtenerHojaMenu = ThisWorkbook.Worksheets.Add: ObtenerHojaMenu.Name = "MENU"
    End If
End Function
