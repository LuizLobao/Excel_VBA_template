Attribute VB_Name = "md_basics"
Sub Congela()
    'Retira visualiza��o
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Sub Descongela()
    'Retorna visualiza��o
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
