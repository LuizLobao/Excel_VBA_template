Attribute VB_Name = "md_Setup"
'Before Start / Antes de Iniciar
'Go to TOOLS / REFERENCES / Activate 'Microsoft ActiveX Data Objects 2.8 Library
'Vá para FERRAMENTAS/ REFERENCIAS / Ative 'Microsoft ActiveX Data Objects 2.8 Library

'*****************************************************************************
'Constantes com controle de versão da planilha
'*****************************************************************************
Public CONTROLEVERSAO As Boolean
Public Const VERSAOPLANILHA = "3.1.1"
Public Const id_RELATORIO = "17"
Public Const nome_RELATORIO = "dashboard_b2b"


'*****************************************************************************
'Nomes de tabelas que serão utilizadas neste relatório
'*****************************************************************************
Public Const tabelaLog = "controle.PERM_LOG_Rel17_dashboard_b2b"
Public Const tabelaEmp = "[dbo].[VW_Rel17_DASHBOARD_B2B]"
