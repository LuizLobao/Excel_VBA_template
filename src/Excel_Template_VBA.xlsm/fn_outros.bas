Attribute VB_Name = "fn_outros"

Function AdminAnaliseDeTemposPadrao(Atividade As String, InicioFim As String)

DataHora = "" 'zero o variavel para garantir que não está trazendo algum valor da memoria
DataHora = Format(Now, "YYYY-MM-DD HH:MM:SS")

'localiza a coluna ATIVIDADE na linha 3
ColAtividade = fnLocalizaColunaCabecalho("AdminTempos", 3, "Atividade")

'localiza linha da Atividade para registro das informações
LinAtividade = fnLocalizaLinhaCabecalho("AdminTempos", CInt(ColAtividade), Atividade)

'registro a hora de inicio ou fim com base na variavel recebida
If InicioFim = "inicio" Then
    AdminTempos.Cells(LinAtividade, ColAtividade + 1) = DataHora
ElseIf InicioFim = "fim" Then
    AdminTempos.Cells(LinAtividade, ColAtividade + 2) = DataHora
End If


End Function

Function AdicionaComentario(aba As String, nomeRange As String, Comentario As String)
Dim MyComments  As Comment

    With Worksheets(aba).Range(nomeRange)
        .ClearComments
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:=Comentario
    End With
    
For Each MyComments In Worksheets(aba).Comments
With MyComments
      .Shape.Width = 200
      .Shape.Height = 250
      '.Shape.AutoShapeType = msoShapeRoundedRectangle
        .Shape.TextFrame.Characters.Font.Name = "Simplon BP Regular"
        .Shape.TextFrame.Characters.Font.Size = 8
        .Shape.TextFrame.Characters.Font.ColorIndex = 2
        .Shape.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Shape.Line.BackColor.RGB = RGB(255, 255, 255)
        .Shape.Fill.Visible = msoTrue
        .Shape.Fill.ForeColor.RGB = RGB(58, 82, 184)
        .Shape.Fill.OneColorGradient msoGradientDiagonalUp, 1, 0.23
    End With
  Next ' comment

    
End Function

Function AtualizadoEm(aba As String, nomeRange As String, Comentario As String)
    
    Worksheets(aba).Range(nomeRange) = Comentario

End Function
Sub AlteraVariavelFiltrosParaUm()
'Objetivo é registrar que algum filtro foi alterado e a consulta ainda não foi executada

        RelAtivo = MostraRelatorioAtivo
        
        If RelAtivo = "Relatorio1" Then
            FiltrosRel1 = 1
        ElseIf RelAtivo = "Relatorio2" Then
            FiltrosRel2 = 1
        ElseIf RelAtivo = "Relatorio3" Then
            FiltrosRel3 = 1
        ElseIf RelAtivo = "Relatorio4" Then
            FiltrosRel4 = 1
        End If
        
        With Worksheets(RelAtivo).Shapes("btn_ExecutarConsulta")
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
        
End Sub
Sub AlteraVariavelFiltrosParaZero()

        RelAtivo = MostraRelatorioAtivo
        
        If RelAtivo = "Relatorio1" Then
            FiltrosRel1 = 0
        ElseIf RelAtivo = "Relatorio2" Then
            FiltrosRel2 = 0
        ElseIf RelAtivo = "Relatorio3" Then
            FiltrosRel3 = 0
        ElseIf RelAtivo = "Relatorio4" Then
            FiltrosRel4 = 0
        End If
        
        With Worksheets(RelAtivo).Shapes("btn_ExecutarConsulta")
            .Fill.ForeColor.RGB = RGB(38, 38, 38)
        End With
        
End Sub
