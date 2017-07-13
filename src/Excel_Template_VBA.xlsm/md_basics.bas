Attribute VB_Name = "md_basics"
Sub Congela()
    'Retira visualização
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Sub Descongela()
    'Retorna visualização
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
