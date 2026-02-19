Attribute VB_Name = "mdl_UI_Menu"
Option Explicit

Public Sub PrepararAspectoApp(ws As Worksheet)
    With ws
        .Cells.Clear
        .Cells.Interior.Color = RGB(245, 247, 250)
        .Cells.RowHeight = 22
        .Cells.ColumnWidth = 4
    End With
    
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = False
    End With
End Sub

Public Sub LimpiarShapes(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
End Sub

Public Sub CrearMenuPrincipal(ws As Worksheet)
    CrearTitulo ws
    CrearBoton ws, "Generar Excel 97", "Start_GenerarXLS", 120
    CrearBoton ws, "Comparar archivos", "Start_Comparar", 190
    CrearBoton ws, "Validar celdas", "Start_Validar", 260
    CrearEstado ws, "Estado: Listo"
End Sub

Private Sub CrearTitulo(ws As Worksheet)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 30, 500, 40)
    With shp.TextFrame2.TextRange
        .Text = "Herramienta de Validación y Exportación"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    End With
    shp.Line.Visible = msoFalse
End Sub

Private Sub CrearBoton(ws As Worksheet, texto As String, macro As String, topPos As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 120, topPos, 320, 50)
    
    With shp
        .Fill.ForeColor.RGB = RGB(52, 84, 153)
        .Line.Visible = msoFalse
        .OnAction = macro
        .TextFrame2.TextRange.Text = texto
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub CrearEstado(ws As Worksheet, texto As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 120, 340, 320, 30)
    With shp.TextFrame2.TextRange
        .Text = texto
        .Font.Size = 10
        .Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    End With
    shp.Line.Visible = msoFalse
End Sub
