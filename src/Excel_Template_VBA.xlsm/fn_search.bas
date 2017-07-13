Attribute VB_Name = "fn_search"
'**************************************************************************************************
'Converte o valor numérico de uma coluna para sua letra
'**************************************************************************************************
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'**************************************************************************************************
'Com base em um texto a ser procurado, retorna em qual COLUNA o texto se encontra
'ABA = nome da woarksheet onde deve ser feita a buca
'LinhaCabecalho = em qual linha deve ser feita a busca
'TextoProcurado = string que será buscada
'**************************************************************************************************
Function fnLocalizaColunaCabecalho(aba As String, LinhaCabecalho As Integer, TextoProcurado As String) As Integer
    
    'Ultima Coluna com cabeçalho
    ColunaFinal = Worksheets(aba).Cells(LinhaCabecalho, Columns.Count).End(xlToLeft).Column
        
    coluna = 1
    Do While coluna <= ColunaFinal
        If Worksheets(aba).Cells(LinhaCabecalho, coluna) = TextoProcurado Then
            ColCabecalho = coluna
        End If
    coluna = coluna + 1
    Loop
    
    fnLocalizaColunaCabecalho = ColCabecalho
End Function

'**************************************************************************************************
'Com base em um texto a ser procurado, retorna em qual LINHA o texto se encontra
'ABA = nome da woarksheet onde deve ser feita a buca
'ColunaCabecalho = em qual coluna deve ser feita a busca
'TextoProcurado = string que será buscada
'**************************************************************************************************
Function fnLocalizaLinhaCabecalho(aba As String, ColunaCabecalho As Integer, TextoProcurado As String) As Integer
    
    'Ultima Coluna com cabeçalho
    LinhaFinal = Worksheets(aba).Cells(Rows.Count, ColunaCabecalho).End(xlUp).Row
    
    Linha = 1
    Do While Linha <= LinhaFinal
        If Worksheets(aba).Cells(Linha, ColunaCabecalho) = TextoProcurado Then
            LinCabecalho = Linha
            fnLocalizaLinhaCabecalho = LinCabecalho
            Exit Function
        End If
    Linha = Linha + 1
    Loop
    
    fnLocalizaLinhaCabecalho = LinCabecalho
End Function

'**************************************************************************************************
'Retorna o nome do relatório/aba ativa
'No meu caso, se não estiver em uma das opções pre-definidas, retorna "teste"
'**************************************************************************************************
Function MostraRelatorioAtivo() As String
    'identificar o relatório ativo
    RelatorioAtivo = ActiveSheet.Name
    
    If RelatorioAtivo <> "Capa" And _
       RelatorioAtivo <> "Relatorio1" And _
       RelatorioAtivo <> "Relatorio2" And _
       RelatorioAtivo <> "Relatorio3" And _
       RelatorioAtivo <> "Relatorio4" Then
       RelatorioAtivo = "teste"
    End If
    MostraRelatorioAtivo = RelatorioAtivo
End Function


'**************************************************************************************************
'
'**************************************************************************************************
Function AchaTextoMenuSelecionado(RelatorioAtivo As String, IndicadorProcurado As String) As String

    'Define onde serão feitas as buscas
    AbaBusca = "AdminMenuSelecionados"
    LinhaTitulos = 3
    
    'Localizar a coluna do INDICADOR PROCURADO na aba AbaBusca
    colInicial = fnLocalizaColunaCabecalho(AbaBusca, LinhaTitulos, "Nome_Relatorio") 'defini que a busca deve começar na coluna com esse texto
    colGU = fnLocalizaColunaCabecalho(AbaBusca, LinhaTitulos, IndicadorProcurado)
    
    'Localizar a coluna do INDICADOR PROCURADO na aba AdminMenuOpções
    colGU2 = fnLocalizaColunaCabecalho("AdminMenuOpções", 5, IndicadorProcurado)
    
    'Localizar a linha do relatório na aba AdminMenuSelecionados
    LinhaFinal = Worksheets(AbaBusca).Cells(Rows.Count, colInicial).End(xlUp).Row 'Ultima linha
    lin = 1
    Do While lin <= LinhaFinal
        If Worksheets(AbaBusca).Cells(lin, colInicial) = RelatorioAtivo Then
            linharelatorio = lin
        End If
    lin = lin + 1
    Loop
    
    ValorProcurado = Worksheets(AbaBusca).Cells(linharelatorio, colGU)
    L = 1
    Do While L <= 50
        If Worksheets("AdminMenuOpções").Cells(L, 2) = ValorProcurado Then
            ValorAchado = Worksheets("AdminMenuOpções").Cells(L, colGU2)
        End If
        L = L + 1
    Loop
    
    Worksheets(AbaBusca).Cells(linharelatorio + 1, colGU) = ValorAchado
    AchaTextoMenuSelecionado = ValorAchado

End Function



