Attribute VB_Name = "MENU_Starts"
 
Option Explicit

Public Sub Start_GenerarXLS(): Call MENU_Navegador.AbrirPestañaProceso("VCA_Espana", ThisWorkbook.Sheets("MENU")): End Sub
Public Sub Start_Comparar(): Call MENU_Navegador.AbrirPestañaProceso("VCA_Portugal", ThisWorkbook.Sheets("MENU")): End Sub
Public Sub Start_Config(): Call MENU_Navegador.AbrirPestañaProceso("VCA_Config", ThisWorkbook.Sheets("MENU")): End Sub

Public Sub EliminarHojas_Click()
    Dim ws As Object ' Cambiamos a Object para detectar CUALQUIER tipo de hoja
    Dim contador As Integer
    Dim listaOculta As String
    
    contador = 0
    listaOculta = ""
    
    ' Recorremos TODAS las hojas (Worksheets, Charts, etc.)
    For Each ws In ThisWorkbook.Sheets
        ' Limpiamos el nombre de espacios y comparamos
        If Trim(UCase(ws.Name)) <> "MENU" Then
            contador = contador + 1
            listaOculta = listaOculta & "- " & ws.Name & " (Tipo: " & TypeName(ws) & ")" & vbCrLf
        End If
    Next ws
    
    If contador = 0 Then
        MsgBox "No hay hojas para borrar.", vbInformation, "Limpieza"
        Exit Sub
    End If
    
    ' El mensaje te dirá exactamente QUÉ está contando
    Dim msg As String
    msg = "Excel quiere borrar " & contador & " hojas:" & vbCrLf & vbCrLf & _
          listaOculta & vbCrLf & "¿Proceder con la eliminación?"
          
    If MsgBox(msg, vbQuestion + vbYesNo, "Confirmar Borrado") = vbYes Then
        Application.DisplayAlerts = False
        For Each ws In ThisWorkbook.Sheets
            If Trim(UCase(ws.Name)) <> "MENU" Then
                ws.Visible = -1 ' xlSheetVisible
                ws.Delete
            End If
        Next ws
        Application.DisplayAlerts = True
        
        Call MENU_Core.MostrarMenuInicial
        MsgBox "Limpieza realizada.", vbInformation
    End If
End Sub

' Eventos
Public Sub Paso1_Importar_ESP(): MsgBox "Importación España": End Sub
Public Sub Paso2_Calcular_ESP(): MsgBox "Cálculo España": End Sub
Public Sub Paso1_Importar_POR(): MsgBox "Importación Portugal": End Sub
Public Sub Paso2_Calcular_POR(): MsgBox "Cálculo Portugal": End Sub
Public Sub Paso_Config_A(): MsgBox "Ajuste A": End Sub
Public Sub Paso_Config_B(): MsgBox "Ajuste B": End Sub

