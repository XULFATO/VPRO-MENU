Attribute VB_Name = "mdl_AppState"
Option Explicit

Private Type TAppState
    DisplayFormulaBar As Boolean
    DisplayStatusBar As Boolean
    DisplayWorkbookTabs As Boolean
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
End Type

Private gState As TAppState

Public Sub GuardarEstadoExcel()
    With gState
        .DisplayFormulaBar = Application.DisplayFormulaBar
        .DisplayStatusBar = Application.DisplayStatusBar
        .DisplayWorkbookTabs = ActiveWindow.DisplayWorkbookTabs
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .Calculation = Application.Calculation
    End With
End Sub

Public Sub AplicarModoApp()
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub RestaurarEstadoExcel()
    With gState
        Application.DisplayFormulaBar = .DisplayFormulaBar
        Application.DisplayStatusBar = .DisplayStatusBar
        ActiveWindow.DisplayWorkbookTabs = .DisplayWorkbookTabs
        Application.ScreenUpdating = .ScreenUpdating
        Application.EnableEvents = .EnableEvents
        Application.Calculation = .Calculation
    End With
End Sub
