VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe de conex�o com banco de dados
'Desenvolvido por Thiago Penha
'Data: 11/06/2013

Private host As String
Private login As String
Private password As String
Private bancoDados As String

Private Sub Class_Initialize()
    'Inicializa��o / Start
End Sub

Private Sub Class_Terminate()
    'Finaliza��o / Finish
End Sub

'Passagem dos parametros para conex�o
Public Sub setValores(ByVal tempHost As String, ByVal tempLogin As String, ByVal tempPassword As String, ByVal tempBancoDados As String)

'Parametro de conex�o com o banco de dados
    host = tempHost
    login = tempLogin
    password = tempPassword
    bancoDados = tempBancoDados

End Sub

'Fun��o para abertura da conex�o com o BD
Public Function abrirConexao(ByRef conexao As ADODB.Connection) As ADODB.Connection
Set conexao = New ADODB.Connection

    'Informa��es do banco de dados
    hostBancoDados = host
    nomeBanco = bancoDados
    userBanco = login
    passwordBanco = password

    On Error Resume Next
    conexao.ConnectionString = "Provider=SQLOLEDB;Server=" & hostBancoDados & ";Database=" & nomeBanco & ";User Id=" & userBanco & ";Password=" & passwordBanco & ";"

    conexao.CursorLocation = adUseClient
    conexao.Open
    conexao.CommandTimeout = 0

    'Se n�o conseguir conex�o, informa erro
    If (conexao.State = 0) Then
        MsgBox "N�o foi poss�vel realizar conex�o com o banco de dados" & vbCrLf & conexao.Errors(0), vbCritical, "Erro banco de dados"
        End
    End If

    OpenConect = conexao

End Function

'Fun��o de encerramento de conex�o
Public Function fecharConexao(ByRef conexao As ADODB.Connection) As ADODB.Connection

    'Se conex�o estiver aberta, fecha
    If (conexao.State <> 0) Then
        conexao.Close
    End If

End Function

'Fun��o para testar conex�o
Public Function testarConexao(ByRef conexao As ADODB.Connection) As Boolean

    'Testa conex�o (se sucedida, retorna true)
    If (conexao.State = 1) Then
        testarConexao = True
    Else
   
        'N�o conseguiu conex�o, retorna falso
        testarConexao = False
    End If
End Function

'Fun��o para execu��o de comandos Sql
Public Sub execQuery(ByVal query As String, ByRef cnn As ADODB.Connection, ByRef rst As Recordset)

    'Executa comando
    rst.Open query, cnn, adOpenStatic, adLockOptimistic

End Sub

'Fun��o para abertura da conex�o com o BD
Public Function testConexao(ByRef conexao As ADODB.Connection) As ADODB.Connection
Set conexao = New ADODB.Connection

    'Informa��es do banco de dados
    hostBancoDados = host
    nomeBanco = bancoDados
    userBanco = login
    passwordBanco = password

    On Error Resume Next
    conexao.ConnectionString = "Provider=SQLOLEDB;Server=" & hostBancoDados & ";Database=" & nomeBanco & ";User Id=" & userBanco & ";Password=" & passwordBanco & ";"

    conexao.CursorLocation = adUseClient
    conexao.Open
    conexao.CommandTimeout = 0

    'Se n�o conseguir conex�o, informa erro
    If (conexao.State = 0) Then
        TextFile_Create
        MsgBox "N�o foi poss�vel realizar conex�o com o banco de dados" & vbCrLf & conexao.Errors(0), vbCritical, "Erro banco de dados"
        Else
        ImportFromText
    End If

    OpenConect = conexao

End Function

