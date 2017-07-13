Attribute VB_Name = "fn_db"
'*******************************************************************************************************
'Função para executar uma consulta SQL e retornar o resultado em um Worksheet (W) especifico
'Usar a opção TAMANHO para informar a quantidade de colunas que estará retornando - será usado para deletar antes de colar
'QUERY será a string com a consulta a ser feita
'LinhaEscrever define em qual linha da worksheet devera escrever os resultados
'*******************************************************************************************************
Function RodarConsultaPadrao(W As String, coluna As Integer, tamanho As Integer, query As String, LinhaEscrever As Integer)
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim objConexao As New ClasseConexao

    'Converte o numero da coluna para letra utilizando a função Col_Letter
    ColA = Col_Letter(CLng(coluna))
    ColB = Col_Letter(CLng(coluna + tamanho - 1))
    
    If LinhaEscrever = Null Then
        LinhaEscrever = 5
    End If


    'Retira visualização
    Call Congela
    
    'Realiza conexão e consulta ao banco de dados
    objConexao.setValores "sqlprd153,1444", "SAEM6001", "N3w@Pl4nej4_DB", "DB_PLANEMP"
    objConexao.abrirConexao cnn
    
    'Realiza consulta ao banco de dados
    objConexao.execQuery query, cnn, rst
    'Apaga os dados anteriores antes de escrever os novos
    Worksheets(W).Range(ColA & LinhaEscrever & ":" & ColB & 50000).ClearContents
    
    'Cola os valores da tabela
    Worksheets(W).Cells(LinhaEscrever, coluna).CopyFromRecordset rst
    'Encerra conexão apenas quando usar select
    cnn.Close
    
    
    'Retorna visualização
    Call Descongela
    
If Err.Number <> 0 Then
    MsgBox Err.Number & vbCrLf & Err.Description
    Err.Number = 0
End If

End Function
