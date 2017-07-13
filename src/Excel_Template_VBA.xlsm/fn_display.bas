Attribute VB_Name = "fn_display"
Sub modo_apresentacao_in_out()
    If Application.DisplayFullScreen = True Then
            Application.DisplayFullScreen = False
            ActiveWindow.DisplayHeadings = False
            Application.DisplayFormulaBar = True
            ActiveWindow.DisplayWorkbookTabs = False
    Else
            ' modo_apresentacao Macro
            Application.DisplayFullScreen = True
            ActiveWindow.DisplayHeadings = False
            Application.DisplayFormulaBar = False
            ActiveWindow.DisplayWorkbookTabs = False
    End If
    'Coloca o excel Maximizado
    Application.WindowState = xlMaximized
End Sub

Sub Congela()
    'Retira visualizacao
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Sub Descongela()
    'Retorna visualizacao
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub



