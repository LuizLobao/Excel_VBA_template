Attribute VB_Name = "md_menu"
Sub Ajuda_Menu_Flutuante()

    RelAtivo = MostraRelatorioAtivo
    
    If RelAtivo = "Capa" Then
        If Worksheets(RelAtivo).Shapes("Ajuda_Capa_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Comunicados").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_DIY").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Export").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Glossario").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_links").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Seta").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_MenuSusp").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Comunicados").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_DIY").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Export").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Glossario").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_links").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_Seta").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_Capa_MenuSusp").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio1" Then
        If Worksheets(RelAtivo).Shapes("Ajuda_R1_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R1_OrcMeta").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Graf").Visible = False

        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R1_OrcMeta").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R1_Graf").Visible = True
        End If
    End If

End Sub
