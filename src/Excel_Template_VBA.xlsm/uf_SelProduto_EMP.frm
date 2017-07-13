VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_SelProduto_EMP 
   Caption         =   "Seleção de Produtos - Empresarial"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12645
   OleObjectBlob   =   "uf_SelProduto_EMP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_SelProduto_EMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_fechar_Click()
    uf_SelProduto_EMP.Hide
End Sub

Private Sub btn_limpar_Click()
    
    cb_oi_fixo.Value = False
    cb_oi_velox.Value = False
    cb_pos_puro.Value = False
    cb_oct.Value = False
    cb_controle.Value = False
    cb_oi_tv_dth.Value = False
    cb_oi_tv_cabo.Value = False
    cb_oi_tv_ant_pre.Value = False
    cb_oi_tv_ant_pos.Value = False
    cb_cpe.Value = False
    cb_data.Value = False
    cb_frame.Value = False
    cb_ip_connect.Value = False
    cb_vpn_vip.Value = False
    cb_digitronco_circuito.Value = False
    cb_digitronco_canais.Value = False
    cb_voicenet.Value = False
    cb_blackberry.Value = False
    cb_pacote_de_dados.Value = False
    cb_somente_dados.Value = False
    cb_velox_movel.Value = False
    
    'TI
    cb_iaas.Value = False
    
    'CLOUD
    cb_AntiSpam.Value = False
    cb_dominios.Value = False
    cb_email.Value = False
    cb_FiltroConteudo.Value = False
    cb_office365.Value = False
    cb_OiGestaoMobilidade.Value = False
    cb_sharepoint.Value = False
    
    'ICT
    cb_AmbienteAmbulatorio.Value = False
    cb_FiscalizacaoEletronica.Value = False
    cb_GestaoConteudosDigitais.Value = False
    cb_GestaoForcaCampo.Value = False
    cb_GestaoDeFrota.Value = False
    cb_GestaoInfoCanais.Value = False
    cb_MonitoramentoEquipes.Value = False
    cb_telepresenca.Value = False
    
    CheckBox1.Value = False
    
End Sub

Private Sub btn_OK_Click()

'############ FIXO ######################
    If cb_oi_fixo.Value = True Then
        fixo = ", 'OI FIXO'"
        sfixo = "OI FIXO,"
    Else
        fixo = ""
        sfixo = ""
    End If
'############ VELOX ######################
    If cb_oi_velox.Value = True Then
        velox = ", 'OI VELOX'"
        svelox = "OI VELOX,"
    Else
        velox = ""
        svelox = ""
    End If
'############ POS PURO ######################
    If cb_pos_puro.Value = True Then
        pos_puro = ", 'PÓS PURO'"
        spos_puro = "PÓS PURO,"
    Else
        pos_puro = ""
        spos_puro = ""
    End If
'############ OCT ######################
    If cb_oct.Value = True Then
        m_oct = ", 'OCT'"
        sm_oct = "OCT,"
    Else
        m_oct = ""
        sm_oct = ""
    End If
'############ CONTROLE ######################
    If cb_controle.Value = True Then
        m_controle = ", 'CONTROLE'"
        sm_controle = "CONTROLE,"
    Else
        m_controle = ""
        sm_controle = ""
    End If
'############ OI TV DTH ######################
    If cb_oi_tv_dth.Value = True Then
        oi_tv_dth = ", 'OI TV DTH'"
        soi_tv_dth = "OI TV DTH,"
    Else
        oi_tv_dth = ""
        soi_tv_dth = ""
    End If
'############ OI TV CABO ######################
    If cb_oi_tv_cabo.Value = True Then
        oi_tv_cabo = ", 'OI TV CABO'"
        soi_tv_cabo = "OI TV CABO,"
    Else
        oi_tv_cabo = ""
        soi_tv_cabo = ""
    End If
'############ OI TV ANTENEIROS PRE ######################
    If cb_oi_tv_ant_pre.Value = True Then
        oi_tv_ant_pre = ", 'OI TV ANT PRE'"
        soi_tv_ant_pre = "OI TV ANT PRE,"
    Else
        oi_tv_ant_pre = ""
        soi_tv_ant_pre = ""
    End If

'############ OI TV ANTENEIROS POS ######################
    If cb_oi_tv_ant_pos.Value = True Then
        oi_tv_ant_pos = ", 'OI TV ANT POS'"
        soi_tv_ant_pos = "OI TV ANT POS,"
    Else
        oi_tv_ant_pos = ""
        soi_tv_ant_pos = ""
    End If
'############ CPE ######################
    If cb_cpe.Value = True Then
        cpe = ", 'CPE'"
        scpe = "CPE,"
    Else
        cpe = ""
        scpe = ""
    End If
'############ DATA ######################
    If cb_data.Value = True Then
        ddata = ", 'DATA'"
        sddata = "DATA,"
    Else
        ddata = ""
        sddata = ""
    End If
'############ FRAME ######################
    If cb_frame.Value = True Then
        fframe = ", 'FRAME'"
        sfframe = "FRAME,"
    Else
        fframe = ""
        sfframe = ""
    End If
'############ IP CONNECT ######################
    If cb_ip_connect.Value = True Then
        ip_connect = ", 'IP CONNECT'"
        sip_connect = "IP CONNECT,"
    Else
        ip_connect = ""
        sip_connect = ""
    End If
'############ VPN VIP ######################
    If cb_vpn_vip.Value = True Then
        vpn_vip = ", 'VPN VIP'"
        svpn_vip = "VPN VIP,"
    Else
        vpn_vip = ""
        svpn_vip = ""
    End If
'############ DIGITRONCO CIRCUITO ######################
    If cb_digitronco_circuito.Value = True Then
        digitronco_circuito = ", 'DIGITRONCO (CIRCUITOS)'"
        sdigitronco_circuito = "DIGITRONCO (CIRCUITOS),"
    Else
        digitronco_circuito = ""
        sdigitronco_circuito = ""
    End If
'############ DIGITRONCO CANAIS ######################
    If cb_digitronco_canais.Value = True Then
        digitronco_canais = ", 'DIGITRONCO (CANAIS)'"
        sdigitronco_canais = "DIGITRONCO (CANAIS),"
    Else
        digitronco_canais = ""
        sdigitronco_canais = ""
    End If
'############ VOICENET ######################
    If cb_voicenet.Value = True Then
        voicenet = ", 'VOICENET'"
        svoicenet = "VOICENET,"
    Else
        voicenet = ""
        svoicenet = ""
    End If
'############ BLACKBERRY ######################
    If cb_blackberry.Value = True Then
        blackberry = ", 'BLACKBERRY'"
        sblackberry = "BLACKBERRY,"
    Else
        blackberry = ""
        sblackberry = ""
    End If
'############ PACOTE DE DADOS ######################
    If cb_pacote_de_dados.Value = True Then
        pacote_de_dados = ", 'PACOTE DE DADOS'"
        spacote_de_dados = "PACOTE DE DADOS,"
    Else
        pacote_de_dados = ""
        spacote_de_dados = ""
    End If
'############ PACOTE DE DADOS ######################
    If cb_somente_dados.Value = True Then
        somente_dados = ", 'SOMENTE DADOS'"
        ssomente_dados = "SOMENTE DADOS,"
    Else
        somente_dados = ""
        ssomente_dados = ""
    End If
'############ VELOX MOVEL ######################
    If cb_velox_movel.Value = True Then
        velox_movel = ", 'VELOX MÓVEL'"
        svelox_movel = "VELOX MÓVEL,"
    Else
        velox_movel = ""
        svelox_movel = ""
    End If
'###############################################
    If cb_iaas.Value = True Then
        iaas = ", 'IaaS'"
        siaas = "IaaS,"
    Else
        iaas = ""
        siaas = ""
    End If
 '###############################################
    If cb_AntiSpam.Value = True Then
        AntiSpam = ", 'Anti-Spam'"
        sAntiSpam = "Anti-Spam,"
    Else
        AntiSpam = ""
        sAntiSpam = ""
    End If
    
 '###############################################
    If cb_dominios.Value = True Then
        dominios = ", 'Dominios'"
        sdominios = "Dominios,"
    Else
        dominios = ""
        sdominios = ""
    End If
  '###############################################
    If cb_email.Value = True Then
        email = ", 'Email'"
        semail = "Email,"
    Else
        email = ""
        semail = ""
    End If
'###############################################
    If cb_FiltroConteudo.Value = True Then
      FiltroConteudo = ", 'Filtro de Conteúdo'"
      sFiltroConteudo = "Filtro de Conteúdo,"
    Else
      FiltroConteudo = ""
      sFiltroConteudo = ""
    End If
 '###############################################
    If cb_office365.Value = True Then
      office365 = ", 'Office 365'"
      soffice365 = "Office 365,"
    Else
      office365 = ""
      soffice365 = ""
    End If
    
'###############################################
    If cb_OiGestaoMobilidade.Value = True Then
      OiGestaoMobilidade = ", 'Oi Gestão Mobilidade'"
      sOiGestaoMobilidade = "Oi Gestão Mobilidade,"
    Else
      OiGestaoMobilidade = ""
      sOiGestaoMobilidade = ""
    End If
'###############################################
    If cb_sharepoint.Value = True Then
      sharepoint = ", 'Sharepoint'"
      ssharepoint = "Sharepoint,"
    Else
      sharepoint = ""
      ssharepoint = ""
    End If
    
    
    
    
    
'###############################################
    If cb_AmbienteAmbulatorio.Value = True Then
      AmbienteAmbulatorio = ", 'Ambiente Ambulatório'"
      sAmbienteAmbulatorio = "Ambiente Ambulatório,"
    Else
      AmbienteAmbulatorio = ""
      sAmbienteAmbulatorio = ""
    End If
'###############################################
    If cb_FiscalizacaoEletronica.Value = True Then
      FiscalizacaoEletronica = ", 'Fiscalização Eletrônica'"
      sFiscalizacaoEletronica = "Fiscalização Eletrônica,"
    Else
      FiscalizacaoEletronica = ""
      sFiscalizacaoEletronica = ""
    End If
'###############################################
    If cb_GestaoConteudosDigitais.Value = True Then
      GestaoConteudosDigitais = ", 'Gestão de Conteúdos Digitais'"
      sGestaoConteudosDigitais = "Gestão de Conteúdos Digitais,"
    Else
      GestaoConteudosDigitais = ""
      sGestaoConteudosDigitais = ""
    End If
'###############################################
    If cb_GestaoForcaCampo.Value = True Then
      GestaoForcaCampo = ", 'Gestão de Força de Campo'"
      sGestaoForcaCampo = "Gestão de Força de Campo,"
    Else
      GestaoForcaCampo = ""
      sGestaoForcaCampo = ""
    End If
'###############################################
    If cb_GestaoDeFrota.Value = True Then
      GestaoDeFrota = ", 'Gestão de Frotas'"
      sGestaoDeFrota = "Gestão de Frotas,"
    Else
      GestaoDeFrota = ""
      sGestaoDeFrota = ""
    End If
'###############################################
    If cb_GestaoInfoCanais.Value = True Then
      GestaoInfoCanais = ", 'Gestão de Informações de Canais'"
      sGestaoInfoCanais = "Gestão de Informações de Canais,"
    Else
      GestaoInfoCanais = ""
      sGestaoInfoCanais = ""
    End If
'###############################################
    If cb_MonitoramentoEquipes.Value = True Then
      MonitoramentoEquipes = ", 'Monitoramento de Equipes'"
      sMonitoramentoEquipes = "Monitoramento de Equipes,"
    Else
      MonitoramentoEquipes = ""
      sMonitoramentoEquipes = ""
    End If
'###############################################
    If cb_telepresenca.Value = True Then
      telepresenca = ", 'Telepresença'"
      stelepresenca = "Telepresença,"
    Else
      telepresenca = ""
      stelepresenca = ""
    End If

 
    
cmd_where = "and produto in (''" + fixo + velox + pos_puro + m_oct + m_controle + oi_tv_dth + oi_tv_cabo + oi_tv_ant_pre + oi_tv_ant_pos + _
                               cpe + ddata + fframe + ip_connect + vpn_vip + digitronco_circuito + digitronco_canais + _
                               voicenet + blackberry + pacote_de_dados + somente_dados + velox_movel + _
                               AntiSpam + dominios + email + FiltroConteudo + office365 + OiGestaoMobilidade + sharepoint + _
                               iaas + _
                               AmbienteAmbulatorio + FiscalizacaoEletronica + GestaoConteudosDigitais + GestaoForcaCampo + GestaoDeFrota + GestaoInfoCanais + MonitoramentoEquipes + telepresenca + ")"

lista_produto = sfixo + svelox + spos_puro + sm_oct + sm_controle + soi_tv_dth + soi_tv_cabo + soi_tv_ant_pre + soi_tv_ant_pos + _
                scpe + sddata + sfframe + sip_connect + svpn_vip + sdigitronco_circuito + sdigitronco_canais + _
                svoicenet + sblackberry + spacote_de_dados + ssomente_dados + svelox_movel + _
                sAntiSpam + sdominios + semail + sFiltroConteudo + soffice365 + sOiGestaoMobilidade + ssharepoint + _
                siaas + _
                sAmbienteAmbulatorio + sFiscalizacaoEletronica + sGestaoConteudosDigitais + sGestaoForcaCampo + sGestaoDeFrota + sGestaoInfoCanais + sMonitoramentoEquipes + stelepresenca
                
Prod_Selecionados = "Produto(s) Selecionado(s) : " + fixo + velox + pos_puro + m_oct + m_controle + oi_tv_dth + oi_tv_cabo + oi_tv_ant_pre + oi_tv_ant_pos + _
            cpe + ddata + fframe + ip_connect + vpn_vip + digitronco_circuito + digitronco_canais + _
            voicenet + blackberry + pacote_de_dados + somente_dados + velox_movel + _
            AntiSpam + dominios + email + FiltroConteudo + office365 + OiGestaoMobilidade + sharepoint + _
            iaas + _
            AmbienteAmbulatorio + FiscalizacaoEletronica + GestaoConteudosDigitais + GestaoForcaCampo + GestaoDeFrota + GestaoInfoCanais + MonitoramentoEquipes + telepresenca
            


    relatorio_ativo = ActiveSheet.Name
    If relatorio_ativo = "Relatório" Then
        gRelatorio1.Range("E2") = Prod_Selecionados
        gBase.Range("B4") = cmd_where
        gMenu.Range("D10") = lista_produto
    ElseIf relatorio_ativo = "Relatório2" Then
        gRelatorio2.Range("E2") = Prod_Selecionados
        gBase.Range("H4") = cmd_where
        gBase.Range("BG4") = cmd_where 'GROSS VS POSSE
    ElseIf relatorio_ativo = "Relatório3" Then
    
    ElseIf relatorio_ativo = "Relatório4" Then
        gRelatorio4.Range("E2") = Prod_Selecionados
        gBase.Range("S4") = cmd_where
    ElseIf relatorio_ativo = "Relatório7" Then
        gRelatorio7.Range("E2") = Prod_Selecionados
        gBase.Range("BA4") = cmd_where
     ElseIf relatorio_ativo = "Relatório12" Then
        gRelatorio12.Range("E2") = Prod_Selecionados
        gBase.Range("ES2") = cmd_where
    End If
    
'gRelatorio1.Range("E2") = Prod_Selecionados
'gRelatorio2.Range("E2") = Prod_Selecionados
'gRelatorio4.Range("E2") = Prod_Selecionados
'gRelatorio7.Range("E2") = Prod_Selecionados
            
'gBase.Range("B4") = cmd_where
'gBase.Range("H4") = cmd_where
'gBase.Range("S4") = cmd_where
'gBase.Range("BA4") = cmd_where

Executa_Consultas

uf_SelProduto_EMP.Hide

End Sub

Private Sub UserForm_Activate()
relatorio_ativo = ActiveSheet.Name
    If relatorio_ativo = "Relatório12" Then
     cb_oi_fixo.Enabled = True
    cb_oi_velox.Enabled = True
    cb_pos_puro.Enabled = True
    cb_oct.Enabled = True
    cb_controle.Enabled = True
    cb_oi_tv_dth.Enabled = False
    cb_oi_tv_cabo.Enabled = False
    cb_oi_tv_ant_pre.Enabled = False
    cb_oi_tv_ant_pos.Enabled = False
    cb_cpe.Enabled = False
    cb_data.Enabled = False
    cb_frame.Enabled = False
    cb_ip_connect.Enabled = False
    cb_vpn_vip.Enabled = False
    cb_digitronco_circuito.Enabled = False
    cb_digitronco_canais.Enabled = False
    cb_voicenet.Enabled = False
    cb_blackberry.Enabled = False
    cb_pacote_de_dados.Enabled = False
    cb_somente_dados.Enabled = True
    cb_velox_movel.Enabled = True
    CheckBox1.Enabled = False
    'TI
    cb_iaas.Enabled = False
    
    'CLOUD
    cb_AntiSpam.Enabled = False
    cb_dominios.Enabled = False
    cb_email.Enabled = False
    cb_FiltroConteudo.Enabled = False
    cb_office365.Enabled = False
    cb_OiGestaoMobilidade.Enabled = False
    cb_sharepoint.Enabled = False
    
    'ICT
    cb_AmbienteAmbulatorio.Enabled = False
    cb_FiscalizacaoEletronica.Enabled = False
    cb_GestaoConteudosDigitais.Enabled = False
    cb_GestaoForcaCampo.Enabled = False
    cb_GestaoDeFrota.Enabled = False
    cb_GestaoInfoCanais.Enabled = False
    cb_MonitoramentoEquipes.Enabled = False
    cb_telepresenca.Enabled = False
    
    End If
End Sub


