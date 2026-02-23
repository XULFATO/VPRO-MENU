Attribute VB_Name = "MENU_Logic"
 
Option Explicit

Public Sub DibujarBotonesVCA(ByVal nombreHoja As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    
    ws.Activate
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = True
    End With
    ws.Cells.Interior.Color = RGB(245, 247, 250)
    
    Dim shp As Shape
    For Each shp In ws.Shapes: shp.Delete: Next
    
    Select Case nombreHoja
        Case "VCA_Espana"
            CrearBoton ws, "PASO 1: Importar ESP", "Paso1_Importar_ESP", 60, RGB(255, 204, 0), RGB(0, 0, 0)
            CrearBoton ws, "PASO 2: Generar VCA", "Paso2_Calcular_ESP", 120, RGB(204, 0, 0), RGB(255, 255, 255)
        Case "VCA_Portugal"
            CrearBoton ws, "PASO 1: Importar POR", "Paso1_Importar_POR", 60, RGB(0, 102, 0), RGB(255, 255, 255)
            CrearBoton ws, "PASO 2: Generar VCA", "Paso2_Calcular_POR", 120, RGB(255, 0, 0), RGB(255, 255, 255)
        Case "VCA_Config"
            CrearBoton ws, "AJUSTAR PARÁMETROS", "Paso_Config_A", 60, RGB(255, 150, 0), RGB(255, 255, 255)
            CrearBoton ws, "ACTUALIZAR MAESTROS", "Paso_Config_B", 120, RGB(130, 0, 200), RGB(255, 255, 255)
    End Select
End Sub

Private Sub CrearBoton(ws As Worksheet, texto As String, macro As String, topPos As Double, colorRelleno As Long, colorTexto As Long)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, topPos, 280, 45)
    With btn
        .Fill.ForeColor.RGB = colorRelleno
        .Line.Visible = msoFalse
        .OnAction = macro
        With .TextFrame2.TextRange
            .Text = texto: .Font.Size = 11: .Font.Bold = True
            .Font.Fill.ForeColor.RGB = colorTexto
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
End Sub
