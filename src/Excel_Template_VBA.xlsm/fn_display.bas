Attribute VB_Name = "fn_display"
Sub modo_apresenta��o()
    If Application.DisplayFullScreen = True Then
            Application.DisplayFullScreen = False
            ActiveWindow.DisplayHeadings = False
            Application.DisplayFormulaBar = True
            ActiveWindow.DisplayWorkbookTabs = False
    Else
            ' modo_apresenta��o Macro
            Application.DisplayFullScreen = True
            ActiveWindow.DisplayHeadings = False
            Application.DisplayFormulaBar = False
            ActiveWindow.DisplayWorkbookTabs = False
    End If
    'Coloca o excel Maximizado
    Application.WindowState = xlMaximized
End Sub




