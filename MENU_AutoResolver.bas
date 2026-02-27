Attribute VB_Name = "MENU_AutoResolver"
Option Explicit

' -----------------------------------------------------------------------
' MENU_AutoResolver
' Construye en tiempo de ejecucion el diccionario de macros escaneando
' el propio VBProject. No hay nada hardcodeado que mantener a mano.
'
' COMO FUNCIONA:
'   1. Escanea todos los modulos del proyecto buscando Subs publicos
'   2. Para cada Sub detecta su "firma logica" comparando con una lista
'      de patrones invariantes (palabras clave del negocio)
'   3. Mapea patron -> nombre real actual del Sub
'   4. CrearBoton llama a Macro("ESP_PASO1") y obtiene el nombre real
'
' TRAS OFUSCAR:
'   Los nombres de Sub cambian pero las FIRMAS LOGICAS no, porque estan
'   definidas aqui como strings y el ofuscador no las toca.
'   El escaneo encuentra el nuevo nombre automaticamente.
' -----------------------------------------------------------------------

Private mDict       As Object   ' clave_logica -> nombre_sub_real
Private mInicializado As Boolean

' -----------------------------------------------------------------------
' TABLA DE FIRMAS
' Define que Sub buscamos usando fragmentos del nombre original.
' Estas claves NUNCA cambian — son el contrato entre Logic y Starts.
' Añade aqui una entrada por cada boton que uses en el proyecto.
' -----------------------------------------------------------------------
Private Function TablaDeFirmas() As Object
    Dim t As Object
    Set t = CreateObject("Scripting.Dictionary")
    t.CompareMode = 1   ' case-insensitive

    ' Formato: t("CLAVE_LOGICA") = "fragmento_del_nombre_original"
    ' El fragmento debe ser unico dentro del proyecto.
    t("ESP_PASO1")    = "Importar_ESP"
    t("ESP_PASO2")    = "Calcular_ESP"
    t("POR_PASO1")    = "Importar_POR"
    t("POR_PASO2")    = "Calcular_POR"
    t("CONFIG_A")     = "Config_A"
    t("CONFIG_B")     = "Config_B"
    t("MENU_XLS")     = "GenerarXLS"
    t("MENU_COMP")    = "Comparar"
    t("MENU_CFG")     = "Start_Config"
    t("MENU_LIMPIAR") = "EliminarHojas"

    Set TablaDeFirmas = t
End Function

' -----------------------------------------------------------------------
' InicializarResolver
' Escanea el VBProject y resuelve cada firma al Sub real actual.
' Se llama automaticamente la primera vez que se pide una macro.
' -----------------------------------------------------------------------
Public Sub InicializarResolver()
    Set mDict = CreateObject("Scripting.Dictionary")
    mDict.CompareMode = 1

    Dim firmas As Object
    Set firmas = TablaDeFirmas()

    ' Recoger todos los nombres de Subs publicos del proyecto
    Dim todosLosSubs() As String
    todosLosSubs = EscanearSubsPublicos()

    ' Para cada firma, buscar el Sub cuyo nombre contiene el fragmento
    Dim clave As Variant
    Dim nombreSub As String
    Dim i As Long

    For Each clave In firmas.Keys
        Dim fragmento As String
        fragmento = CStr(firmas(clave))
        nombreSub = ""

        For i = LBound(todosLosSubs) To UBound(todosLosSubs)
            If InStr(1, todosLosSubs(i), fragmento, vbTextCompare) > 0 Then
                nombreSub = todosLosSubs(i)
                Exit For
            End If
        Next i

        ' Guardar el resultado (vacio si no se encontro)
        mDict(CStr(clave)) = nombreSub
    Next clave

    mInicializado = True
End Sub

' -----------------------------------------------------------------------
' EscanearSubsPublicos
' Devuelve array con los nombres de todos los Subs/Functions publicos
' del proyecto actual, escaneando el codigo fuente linea a linea.
' -----------------------------------------------------------------------
Private Function EscanearSubsPublicos() As String()
    Dim resultado() As String
    Dim n As Long
    ReDim resultado(0)
    n = 0

    On Error Resume Next

    Dim vbc As Object
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Then
            Dim i As Long
            For i = 1 To vbc.CodeModule.CountOfLines
                Dim linea As String
                linea = Trim(vbc.CodeModule.Lines(i, 1))
                Dim uLinea As String
                uLinea = UCase(linea)

                ' Detectar Public Sub o Public Function
                If Left(uLinea, 11) = "PUBLIC SUB " Or _
                   Left(uLinea, 16) = "PUBLIC FUNCTION " Then

                    Dim nombre As String
                    nombre = ExtraerNombre(linea)

                    If Len(nombre) > 0 Then
                        If n > 0 Then ReDim Preserve resultado(0 To n)
                        resultado(n) = nombre
                        n = n + 1
                    End If
                End If
            Next i
        End If
    Next vbc

    On Error GoTo 0

    If n = 0 Then
        ReDim resultado(0)
        resultado(0) = ""
    End If

    EscanearSubsPublicos = resultado
End Function

' -----------------------------------------------------------------------
' ExtraerNombre
' Extrae el nombre del Sub/Function de una linea de declaracion.
' -----------------------------------------------------------------------
Private Function ExtraerNombre(ByVal linea As String) As String
    Dim partes() As String
    Dim p As String
    Dim i As Long

    partes = Split(linea, " ")

    For i = 0 To UBound(partes)
        p = LCase(Trim(partes(i)))
        Select Case p
            Case "public", "private", "friend", "sub", "function", ""
                ' saltar
            Case Else
                Dim resultado As String
                resultado = Trim(partes(i))
                ' limpiar parentesis
                If InStr(resultado, "(") > 0 Then
                    resultado = Left(resultado, InStr(resultado, "(") - 1)
                End If
                ExtraerNombre = Trim(resultado)
                Exit Function
        End Select
    Next i
End Function

' -----------------------------------------------------------------------
' Macro
' Punto de entrada publico. Devuelve el nombre real del Sub
' para una clave logica. Si no encuentra nada devuelve la clave
' para no romper el OnAction silenciosamente.
' -----------------------------------------------------------------------
Public Function Macro(ByVal claveLogica As String) As String
    If Not mInicializado Then InicializarResolver

    If mDict.Exists(claveLogica) Then
        Dim nombre As String
        nombre = CStr(mDict(claveLogica))
        If Len(nombre) > 0 Then
            Macro = nombre
            Exit Function
        End If
    End If

    ' Fallback: no se encontro, devolver clave para debug visible
    Macro = "MACRO_NO_ENCONTRADA_" & claveLogica
End Function

' -----------------------------------------------------------------------
' Reset
' Fuerza re-escaneo en la proxima llamada.
' Util si se llama antes de que el VBProject este completamente cargado.
' -----------------------------------------------------------------------
Public Sub Reset()
    mInicializado = False
    Set mDict = Nothing
End Sub

' -----------------------------------------------------------------------
' DiagnosticoResolver
' Muestra en un MsgBox el resultado del escaneo actual.
' Llama a esto para verificar que todo se resolvio correctamente.
' -----------------------------------------------------------------------
Public Sub DiagnosticoResolver()
    If Not mInicializado Then InicializarResolver

    Dim informe As String
    informe = "=== MENU_AutoResolver ===" & vbCrLf & vbCrLf

    Dim clave As Variant
    For Each clave In mDict.Keys
        Dim val As String
        val = CStr(mDict(clave))
        If Len(val) = 0 Then val = "*** NO ENCONTRADO ***"
        informe = informe & CStr(clave) & " -> " & val & vbCrLf
    Next clave

    MsgBox informe, vbInformation, "Diagnostico AutoResolver"
End Sub
