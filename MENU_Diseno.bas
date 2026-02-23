Attribute VB_Name = "MENU_Diseno"

Option Explicit

Public Sub PrepararAspectoApp(ws As Worksheet)
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(245, 247, 250)
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = True
    End With
End Sub

Public Sub LimpiarShapes(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes: shp.Delete: Next
End Sub

Public Sub CrearMenuPrincipal(ws As Worksheet)
    ' Título
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 30, 500, 40)
    shp.TextFrame2.TextRange.Text = "SISTEMA DE GESTIÓN VCA"
    shp.TextFrame2.TextRange.Font.Size = 20
    shp.TextFrame2.TextRange.Font.Bold = True
    shp.Line.Visible = msoFalse

    ' Botones Navegación
    CrearBoton ws, "GENERAR ESPAÑA", "Start_GenerarXLS", 100, RGB(52, 84, 153)
    CrearBoton ws, "COMPARAR PORTUGAL", "Start_Comparar", 160, RGB(52, 84, 153)
    CrearBoton ws, "CONFIGURACIÓN", "Start_Config", 220, RGB(70, 90, 110)
    
    ' Botón Limpieza
    CrearBoton ws, "ELIMINAR TODAS LAS HOJAS", "EliminarHojas_Click", 310, RGB(200, 0, 0)
End Sub

Private Sub CrearBoton(ws As Worksheet, texto As String, macro As String, topPos As Long, colorRelleno As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 120, topPos, 320, 45)
    With shp
        .Fill.ForeColor.RGB = colorRelleno
        .Line.Visible = msoFalse
        .OnAction = macro
        With .TextFrame2.TextRange
            .Text = texto: .Font.Size = 12: .Font.Bold = True
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
End Sub
