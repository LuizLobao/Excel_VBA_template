Attribute VB_Name = "md_menu"
Sub Ajuda_Menu_Flutuante()

    RelAtivo = MostraRelatorioAtivo
    
    If RelAtivo = "Capa" Then
       acao = "Exibir/Ocultar Ajuda - Capa"
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
    acao = "Exibir/Ocultar Ajuda - Acompanhamento Mensal"
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
    ElseIf RelAtivo = "Relatorio2" Then
    acao = "Exibir/Ocultar Ajuda - Acompanhamento Diario"
    If Worksheets(RelAtivo).Shapes("Ajuda_R2_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Meses").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R2_DiaAcum").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Semanal").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Meses").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R2_DiaAcum").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R2_Semanal").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio3" Then
    acao = "Exibir/Ocultar Ajuda - Acompanhamento Gross e ARPU"
    If Worksheets(RelAtivo).Shapes("Ajuda_R3_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Meses").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R3_Meses").Visible = True

        End If
    ElseIf RelAtivo = "Relatorio4" Then
    acao = "Exibir/Ocultar Ajuda - Painel de Indicadores"
    If Worksheets(RelAtivo).Shapes("Ajuda_R4_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Meses").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R4_pct").Visible = False
            
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R4_Meses").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R4_pct").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio5" Then
    acao = "Exibir/Ocultar Ajuda - Painel de Físicos B2B"
    If Worksheets(RelAtivo).Shapes("Ajuda_R5_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R5_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_ProjEmp").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_DA").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_ProjRetCorp").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_DACorp").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_ProjGrossCorp").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R5_YTD").Visible = False
            
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R5_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_ProjEmp").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_DA").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_ProjRetCorp").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_DACorp").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_ProjGrossCorp").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R5_YTD").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio6" Then
    acao = "Exibir/Ocultar Ajuda - Churn Precoce - Safra Ativacao"
        If Worksheets(RelAtivo).Shapes("Ajuda_R6_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Menu").Visible = False

        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R6_Menu").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio7" Then
    acao = "Exibir/Ocultar Ajuda - Detalhamento por Produto"
        If Worksheets(RelAtivo).Shapes("Ajuda_R7_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R7_GvsP").Visible = False

        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R7_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R7_GvsP").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio8" Then
    acao = "Exibir/Ocultar Ajuda - Gerencial vs Contabil - Mensal"
        If Worksheets(RelAtivo).Shapes("Ajuda_R8_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Visão").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R8_Visão").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio9" Then
    acao = "Exibir/Ocultar Ajuda - Relatório de Despesas B2B"
        If Worksheets(RelAtivo).Shapes("Ajuda_R9_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Visão").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R9_Visão").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio10" Then
    acao = "Exibir/Ocultar Ajuda - Relatório de Faturamento Liquido"
        If Worksheets(RelAtivo).Shapes("Ajuda_R10_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Visao").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R10_Visao").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio11" Then
    acao = "Exibir/Ocultar Ajuda - Relatório de Curva de Churn"
        If Worksheets(RelAtivo).Shapes("Ajuda_R11_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Menu").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Visao").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Calor1").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Calor2").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Calor3").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Menu").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Visao").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Calor1").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Calor2").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R11_Calor3").Visible = True
        End If
    ElseIf RelAtivo = "Relatorio13" Then
    acao = "Exibir/Ocultar Ajuda - Relatório de Acompanhamento das Retiradas"
        If Worksheets(RelAtivo).Shapes("Ajuda_R13_Logo").Visible = True Then
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Logo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Titulo").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Exec").Visible = False
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Menu").Visible = False
        Else
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Logo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Titulo").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Exec").Visible = True
            Worksheets(RelAtivo).Shapes("Ajuda_R13_Menu").Visible = True
            
        End If
    End If

RegistraLog (acao)
End Sub
