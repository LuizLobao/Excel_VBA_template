VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe de conexão com banco de dados
'Desenvolvido por Thiago Penha
'Data: 11/06/2013

Private host As String
Private login As String
Private password As String
Private bancoDados As String

Private Sub Class_Initialize()
    'Inicialização / Start
End Sub

Private Sub Class_Terminate()
    'Finalização / Finish
End Sub

'Passagem dos parametros para conexão
Public Sub setValores(ByVal tempHost As String, ByVal tempLogin As String, ByVal tempPassword As String, ByVal tempBancoDados As String)

'Parametro de conexão com o banco de dados
    host = tempHost
    login = tempLogin
    password = tempPassword
    bancoDados = tempBancoDados

End Sub

'Função para abertura da conexão com o BD
Public Function abrirConexao(ByRef conexao As ADODB.Connection) As ADODB.Connection
Set conexao = New ADODB.Connection

    'Informações do banco de dados
    hostBancoDados = host
    nomeBanco = bancoDados
    userBanco = login
    passwordBanco = password

    On Error Resume Next
    conexao.ConnectionString = "Provider=SQLOLEDB;Server=" & hostBancoDados & ";Database=" & nomeBanco & ";User Id=" & userBanco & ";Password=" & passwordBanco & ";"

    conexao.CursorLocation = adUseClient
    conexao.Open
    conexao.CommandTimeout = 0

    'Se não conseguir conexão, informa erro
    If (conexao.State = 0) Then
        MsgBox "Não foi possível realizar conexão com o banco de dados" & vbCrLf & conexao.Errors(0), vbCritical, "Erro banco de dados"
        End
    End If

    OpenConect = conexao

End Function

'Função de encerramento de conexão
Public Function fecharConexao(ByRef conexao As ADODB.Connection) As ADODB.Connection

    'Se conexão estiver aberta, fecha
    If (conexao.State <> 0) Then
        conexao.Close
    End If

End Function

'Função para testar conexão
Public Function testarConexao(ByRef conexao As ADODB.Connection) As Boolean

    'Testa conexão (se sucedida, retorna true)
    If (conexao.State = 1) Then
        testarConexao = True
    Else
   
        'Não conseguiu conexão, retorna falso
        testarConexao = False
    End If
End Function

'Função para execução de comandos Sql
Public Sub execQuery(ByVal query As String, ByRef cnn As ADODB.Connection, ByRef rst As Recordset)

    'Executa comando
    rst.Open query, cnn, adOpenStatic, adLockOptimistic

End Sub

'Função para abertura da conexão com o BD
Public Function testConexao(ByRef conexao As ADODB.Connection) As ADODB.Connection
Set conexao = New ADODB.Connection

    'Informações do banco de dados
    hostBancoDados = host
    nomeBanco = bancoDados
    userBanco = login
    passwordBanco = password

    On Error Resume Next
    conexao.ConnectionString = "Provider=SQLOLEDB;Server=" & hostBancoDados & ";Database=" & nomeBanco & ";User Id=" & userBanco & ";Password=" & passwordBanco & ";"

    conexao.CursorLocation = adUseClient
    conexao.Open
    conexao.CommandTimeout = 0

    'Se não conseguir conexão, informa erro
    If (conexao.State = 0) Then
        TextFile_Create
        MsgBox "Não foi possível realizar conexão com o banco de dados" & vbCrLf & conexao.Errors(0), vbCritical, "Erro banco de dados"
        Else
        ImportFromText
    End If

    OpenConect = conexao

End Function

