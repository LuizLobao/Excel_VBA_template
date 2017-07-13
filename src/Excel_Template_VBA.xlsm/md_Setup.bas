Attribute VB_Name = "md_Setup"
'Before Start / Antes de Iniciar
'Go to TOOLS / REFERENCES / Activate 'Microsoft ActiveX Data Objects 2.8 Library
'Vá para FERRAMENTAS/ REFERENCIAS / Ative 'Microsoft ActiveX Data Objects 2.8 Library


'*****************************************************************************
'Nomes de tabelas que serão utilizadas neste relatório
'*****************************************************************************
Public Const tabelaEmp = "[dbo].[VW_Rel17_DASHBOARD_B2B]"



'*****************************************************************************
'LOG DE ACESSO AO RELATORIO
'É preciso criar uma tabela no banco de dados para armazenar os LOGs

                    'CREATE TABLE [controle].[PERM_LOG_Rel17_dashboard_b2b](
                    '    [UsuId] [varchar](20) NOT NULL,
                    '    [Dominio] [varchar](20) NOT NULL,
                    '    [NomeMaquina] [varchar](120) NULL,
                    '    [VersaoOffice] [varchar](150) NULL,
                    '    [DataHora] [datetime] NULL,
                    '    [id_RELATORIO] [varchar](20) NULL,
                    '    [nome_relatorio] [varchar](120) NULL,
                    '    [num_versao] [varchar](20) NULL,
                    '    [ACAO] [varchar](250) NULL
                    ')
Public Const tabelaLog = "controle.PERM_LOG_Rel17_dashboard_b2b"
'*****************************************************************************


'*****************************************************************************
'CONTROLE DE VERSAO
'É preciso ter uma tabela no banco de dados onde iremos registrar qual a versao
'atual de cada relatorio. Podemos usar uma unica tabela para controlar todos
'os relatorios criados pela sua equipe

                    'CREATE TABLE [controle].[PERM_CONTROLE_VERSAO](
                    '    [id_relatorio] [int] NOT NULL,
                    '    [nome_relatorio] [varchar](120) NULL,
                    '    [num_versao] [varchar](20) NOT NULL,
                    '    [data_versao] [datetime] NULL,
                    '    [responsavel_versao] [varchar](150) NULL,
                    '    [mensagem_versao] [varchar](350) NULL,
                    '    [exibir_mensagem] [int] NULL,
                    '    [data_limite_exibicao] [varchar](20) NULL,
                    '    [endereco_nova_versao] [varchar](500) NULL,
                    '    [mudancas_versao] [varchar](255) NULL
                    ')
Public CONTROLEVERSAO As Boolean
Public Const VERSAOPLANILHA = "3.1.1"
Public Const id_RELATORIO = "17"
Public Const nome_RELATORIO = "dashboard_b2b"
'*****************************************************************************
