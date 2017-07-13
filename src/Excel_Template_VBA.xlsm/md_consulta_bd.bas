Attribute VB_Name = "md_consulta_bd"
Sub ExecutaConsultas()
   
    Monta_Where_para_Consultas
    
    'Identifica o Relatório Ativo e localiza onde está localizado no AdminMenuSelecionados
    RelAtivo = MostraRelatorioAtivo

    If RelAtivo = "Relatorio1" Then
        consulta_relatorio_1
    End If
    
Sair:
End Sub

Sub Monta_Where_para_Consultas()
Dim Col As Integer
Dim RelAtivo As String
Dim abaMenuSel As String
Dim LinhaCabecalho As Integer

abaMenuSel = "AdminMenuSelecionados"
LinhaCabecalho = 3

'Regitra Hora de Inicio/Fim da Execução
Call AdminAnaliseDeTemposPadrao("#MontaWhere", "inicio")

'Identifica o Relatório Ativo e localiza onde está localizado no abaMenuSel
RelAtivo = MostraRelatorioAtivo
Col = fnLocalizaColunaCabecalho(abaMenuSel, 3, "Nome_Relatorio")
Linrelat = fnLocalizaLinhaCabecalho(abaMenuSel, Col, RelAtivo)
    
    
'Localiza as colunas dos campos para montar o WHERE
colGU = fnLocalizaColunaCabecalho(abaMenuSel, LinhaCabecalho, "GRUPO_UNIDADE")
colInd = fnLocalizaColunaCabecalho(abaMenuSel, LinhaCabecalho, "INDICADOR")
colUF = fnLocalizaColunaCabecalho(abaMenuSel, LinhaCabecalho, "FILIAL")
colProd = fnLocalizaColunaCabecalho(abaMenuSel, LinhaCabecalho, "PRODUTOS")
colSub2 = fnLocalizaColunaCabecalho(abaMenuSel, LinhaCabecalho, "SUB2")
colTipo = fnLocalizaColunaCabecalho(abaMenuSel, LinhaCabecalho, "TIPO")
colFlagCanc = fnLocalizaColunaCabecalho(abaMenuSel, LinhaCabecalho, "FLAG_CANCELAMENTO")

'Identifica o valor selecionado para cada campo
GrupoUnidade = Worksheets(abaMenuSel).Cells(Linrelat + 1, colGU)
indicador = Worksheets(abaMenuSel).Cells(Linrelat + 1, colInd)
Filial = Worksheets(abaMenuSel).Cells(Linrelat + 1, colUF)
Produto = Worksheets(abaMenuSel).Cells(Linrelat + 1, colProd)
Sub2 = Worksheets(abaMenuSel).Cells(Linrelat + 1, colSub2)
Tipo = Worksheets(abaMenuSel).Cells(Linrelat + 1, colTipo)
FlagCanc = Worksheets(abaMenuSel).Cells(Linrelat + 1, colFlagCanc)

'-------------------------------------------------------------------
If GrupoUnidade = "TOTAL" Then
    WhereGrupoUnidade = ""
Else
    If GrupoUnidade <> "Total B2B" Then
        WhereGrupoUnidade = "grupo_unidade = '" & GrupoUnidade & "'"
    Else
        WhereGrupoUnidade = " grupo_unidade in ('EMPRESARIAL','CORPORATIVO','ATACADO') "
    End If
End If
'-------------------------------------------------------------------
If indicador = "TOTAL" Then
    whereIndicador = ""
Else
    whereIndicador = " and Indicador like '%" & indicador & "%'"
End If
'-------------------------------------------------------------------
If Filial = "TOTAL" Then
    WhereFilial = ""
Else
    WhereFilial = " and Filial in (''" & Filial & ")"
End If
'-------------------------------------------------------------------
If Produto = "TOTAL" Then
    WhereProduto = ""
Else
    WhereProduto = " and Produto in (" & Produto & ")"
End If
'-------------------------------------------------------------------
If Sub2 = "TOTAL" Then
    WhereSub2 = ""
Else
    WhereSub2 = " and Subgrupo_2 in (''" & Sub2 & ")"
End If
'-------------------------------------------------------------------
If Tipo = "TOTAL" Then
    WhereTipo = ""
Else
    WhereTipo = " and Tipo in (''" & Tipo & ")"
End If
'-------------------------------------------------------------------
If FlagCanc = "TOTAL" Then
    WhereFlagCanc = ""
Else
    WhereFlagCanc = " and FLAG_CANCELAMENTO = '" & FlagCanc & "'"
End If
'-------------------------------------------------------------------

'Escreve o comando where na celular correspondente na aba AdminBase (Where1)
cmdWhere = WhereGrupoUnidade & whereIndicador & WhereFilial & _
           WhereProduto & WhereSub2 & _
           WhereMedias & WhereTipo & WhereFlagCanc


'Com base no relatorio Ativo, escreve o cmdWhere
col2 = fnLocalizaColunaCabecalho("AdminBase", 2, "#" & RelAtivo)
Worksheets("AdminBase").Cells(4, col2) = cmdWhere



'Regitra Hora de Inicio/Fim da Execução
Call AdminAnaliseDeTemposPadrao("#MontaWhere", "fim")

End Sub
'##########################################################################################################################
'
'   CONSULTA PARA O RELATÓRIO 1
'
'##########################################################################################################################
Sub consulta_relatorio_1()

Dim Col As Integer
Dim ColRel1 As Integer
Dim qryCompleta As String
Dim RelAtivo As String
Dim tamanho As Integer
Dim abaBases As String
    
    
    'Regitra Hora de Inicio/Fim da Execução
    Call AdminAnaliseDeTemposPadrao("#ConsultaRelatorio1", "inicio")
    
    abaBases = "AdminBase"

    'Retira qualquer filtro na tabela
    On Error Resume Next
    Worksheets(abaBases).ShowAllData
    Err.Clear
    
    'Bloco para localizar qual Grupo Unidade está selecionado para o RelAtivo
    RelAtivo = MostraRelatorioAtivo
    Col = fnLocalizaColunaCabecalho("AdminMenuSelecionados", 3, "Nome_Relatorio")
    Linrelat = fnLocalizaLinhaCabecalho("AdminMenuSelecionados", Col, RelAtivo)
    colGU = fnLocalizaColunaCabecalho("AdminMenuSelecionados", 3, "GRUPO_UNIDADE")
    GU = Worksheets("AdminMenuSelecionados").Cells(Linrelat + 1, colGU)
    
    'Localiza o Where1 e Where2 para o relatório
    ColRel1 = fnLocalizaColunaCabecalho(abaBases, 2, "#Relatorio1")
    Where1 = Worksheets(abaBases).Cells(4, ColRel1)
    Where2 = Worksheets(abaBases).Cells(5, ColRel1)
    

    'Monta Query para EMPRESARIAL
    qrySelectEMP = "Select indicador + '-' + mes_ref as chave, mes_ref, sum(valor) as valor from " & tabelaEmp & " where "
    qryWhereEmp = Where1 + " " + Where2
    qryGroupEmp = " group by indicador, mes_ref"
    qryOrderEmp = " order by indicador, mes_ref"
    qryCompleta = qrySelectEMP & qryWhereEmp & qryGroupEmp & qryOrderEmp
    
    'usar apenas para validar a query que foi gerada
    'gBase.Range("C2") = qryCompleta


    'Chama função para rodar a query desejada
    Call RodarConsultaPadrao(abaBases, ColRel1, 3, qryCompleta, 7)

    'Regitra Hora de Inicio/Fim da Execução
    Call AdminAnaliseDeTemposPadrao("#ConsultaRelatorio1", "fim")

End Sub
