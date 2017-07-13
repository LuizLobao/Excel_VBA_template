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
Public Function RegistraLog(acao As String) As Boolean
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim objConexao As New ClasseConexao

On Error GoTo tratar_erro

UsuId = Environ("USERNAME")
Dominio = Environ("USERDOMAIN")
NomeMaquina = Environ("COMPUTERNAME")
VersaoOffice = versao_office
DataHora = Format(Now, "YYYY-MM-DD HH:MM:SS")
id_REL = id_RELATORIO
nm_rel = nome_RELATORIO
num_vers = VERSAOPLANILHA
   
valores = "'" & UsuId & "', '" & Dominio & "', '" & NomeMaquina & "', '" & VersaoOffice & "', '" & DataHora & "', '" & id_REL & "', '" & nm_rel & "', '" & num_vers & "', '" & acao & "'"


'Realiza conexão e consulta ao banco de dados
objConexao.setValores "sqlprd153,1444", "SAEM6001", "N3w@Pl4nej4_DB", "DB_PLANEMP"
objConexao.abrirConexao cnn

'Realiza consulta ao banco de dados - AJUSTAR NUMERO DO ID_RELATÓRIO ****************
objConexao.execQuery "INSERT INTO " & tabelaLog & " (USUID, Dominio, NomeMaquina, VersaoOffice, DataHora, id_RELATORIO, nome_relatorio,num_versao, acao ) VALUES (" & valores & ")", cnn, rst

Exit Function
tratar_erro:
    MsgBox Err.Description
Sair:
End Function

Public Function versao_office() As String
   
    'Criar um objeto que recebera uma Sheet
    Dim xlApp As Object
    Dim VersionExcel As Integer
   
    Set xlApp = ActiveSheet
    VersionExcel = Val(xlApp.Application.Version)
   
    Select Case VersionExcel
        Case Is > 16
            versao_office = "Maior que 2016"
        Case 16
            versao_office = "2016"
        Case 15
            versao_office = "2013"
        Case 14
            versao_office = "2010"
        Case 12
            versao_office = "2007"
        Case 11
            versao_office = "2003"
        Case 10
            versao_office = "2002"
        Case 9
            versao_office = "2000"
        Case 8
            versao_office = "1997"
        Case Else
            versao_office = "Indeterminada"
    End Select
    versao_office = versao_office & " (" & VersionExcel & ")"
   
    Set xlApp = Nothing

End Function


