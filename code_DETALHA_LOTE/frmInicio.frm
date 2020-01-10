VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInicio 
   Caption         =   "BUSCAR LOTE"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   OleObjectBlob   =   "frmInicio.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Enum tpConsulta
'    terminal = 1
'    Nrc = 2
'    lote = 3
'End Enum
'
'Private Sub cmdBuscar_Click()
'    Call limparSht
'    Call buscarSelecionados
'    MsgBox "Detalhamento de Terminais Finalizado!", vbInformation, "BUSCA DE LOTE"
'End Sub
'Private Sub limparSht()
'    Dim lngLastRow As Long
'    lngLastRow = lngRetornaUltimaLinha(shtPreenchimento, "A")
'    If lngLastRow > 1 Then shtPreenchimento.Range("A2:G" & lngLastRow).Clear
'End Sub
'Private Sub cmdOk_Click()
'    If bCheckBuscaLote Then
'        If Len(Me.txtLote) > 0 Then
'            Call pesquisaNumeroLote(lote)
'        ElseIf Len(Me.txtNrc) > 0 Then
'            nLote = ""
'            Call pesquisaNumeroLote(Nrc)
'        Else
'            nLote = ""
'            Call pesquisaNumeroLote(terminal)
'        End If
'    End If
'End Sub
'Private Sub buscarSelecionados()
'    Dim i As Integer
'
'    If Not countChecked Then MsgBox "Favor selecionar ao menos um lote!", vbInformation, "BUSCA LOTE": Exit Sub
'
'    Me.lblStatus.Visible = True
'    Me.lblStatus.Caption = "Acessando CSO ..."
'    DoEvents
'    Set Pw = New clsPw3270
'    Call gravaLog("RC", "TELA PW")
'    '==========================================
'
'    'Tratamento Erro operação FVIG
'    '=============================
'    Me.lblStatus.Caption = "Acessando FVIG ..."
'    DoEvents
'    If Not acessoOperacao("RC", "FVIG", "FVIG01", 1) Then MsgBox "Erro no Acceso a transação FVIG!", vbInformation, "BUSCA LOTE": Exit Sub
'    '=============================
'
'    'Entrada na tela de busca de lote
'    '====================================
'    Call IntoText(19, 28, "1", 1, False)
'    Call IntoText(19, 34, "2", 1, True)
'    '====================================
'
'    'Valida a entrada da tela
'    '===========================================================================================
'    If Not InStr(1, RetTextoTela(1, 1, 80), "FVID-01") > 0 Then MsgBox "Erro no Acceso a transação FVIG!", vbInformation, "BUSCA LOTE": Exit Sub
'    '===========================================================================================
'
'    For i = 1 To Me.lstDetalhe.ListItems.Count
'        If Me.lstDetalhe.ListItems(i).Checked = True Then
'            Me.lblStatus.Caption = "Buscando Terminais Conta " & Me.lstDetalhe.ListItems(i).text
'            DoEvents
'            Call listarTerminaisLote(CStr(Left(nLote, 3)), CStr(Mid(nLote, 4, 4)), _
'                CStr(Replace(Me.lstDetalhe.ListItems(i).text, "/", "")), _
'                CStr(Me.lstDetalhe.ListItems(i).ListSubItems(1).text), _
'                shtPreenchimento)
'
'        End If
'    Next i
'    Unload Me
'End Sub
'Private Function countChecked() As Boolean
'
'    Dim iContar As Integer
'
'    For iContar = 1 To Me.lstDetalhe.ListItems.Count
'        If Me.lstDetalhe.ListItems(iContar).Checked = True Then
'            countChecked = True
'            Exit Function
'        End If
'    Next iContar
'    countChecked = False
'End Function
'Private Sub pesquisaNumeroLote(iTpConsulta As tpConsulta)
'
'Dim Pw As clsPw3270
'Dim sNumero As String
'Dim sLocal  As String
'Dim vArrayListWiew As Variant
'
'
'Select Case iTpConsulta
'
'    Case 1
'
'        'Iniciar o Pw Logado
'        '==========================================
'        Me.lblStatus.Visible = True
'        Me.lblStatus.Caption = "Acessando CSO ..."
'        DoEvents
'        Set Pw = New clsPw3270
'        Call gravaLog("RC", "TELA PW")
'        '==========================================
'        sNumero = Right(Me.txtTerminal, Len(Me.txtTerminal) - 5)
'        sLocal = Left(Me.txtTerminal, 5)
'
'        'Tratamento Erro operação O010
'        '=============================
'        Me.lblStatus.Visible = True
'        Me.lblStatus.Caption = "Acessando O010 ..."
'        DoEvents
'        If Not acessoOperacao("RC", "O010", "O010-01", 1) Then MsgBox "Erro no Acceso a transação O010!", vbInformation, "BUSCA LOTE": Exit Sub
'        '=============================
'
'        'buscar o numero do lote
'        '===========================================================================================================================
'        Me.lblStatus.Visible = True
'        Me.lblStatus.Caption = "Buscando Numero de Lote ..."
'        DoEvents
'        Call buscaLoteO010("1", "T", sNumero, sLocal, Right(shtDePara.Range("UserCso"), 7), "ES")
'        '===========================================================================================================================
'    Case 2
'
'        'Iniciar o Pw Logado
'        '==========================================
'        Me.lblStatus.Visible = True
'        Me.lblStatus.Caption = "Acessando CSO ..."
'        DoEvents
'        Set Pw = New clsPw3270
'        Call gravaLog("RC", "TELA PW")
'        '==========================================
'        sNumero = Me.txtNrc.Value
'        'sLocal = Left(Me.txtTerminal, 5)
'
'        'Tratamento Erro operação O010
'        '=============================
'        Me.lblStatus.Visible = True
'        Me.lblStatus.Caption = "Acessando O010 ..."
'        DoEvents
'        If Not acessoOperacao("RC", "O010", "O010-01", 1) Then MsgBox "Erro no Acceso a transação O010!", vbInformation, "BUSCA LOTE": Exit Sub
'        '=============================
'
'        'buscar o numero do lote
'        '===========================================================================================================================
'        Me.lblStatus.Visible = True
'        Me.lblStatus.Caption = "Acessando Lote ..."
'        DoEvents
'        Call buscaLoteO010("1", "N", sNumero, "", Right(shtDePara.Range("UserCso"), 7), "ES")
'
'    Case Is = 3
'        nLote = Me.txtLote.Value
'
'End Select
'
'If Len(nLote) = 0 Then MsgBox "Lote não identificado para o terminal informado!", vbInformation, "BUSCA LOTE": Exit Sub
'Me.txtLote = nLote
'
'Me.lblStatus.Visible = True
'Me.lblStatus.Caption = "Listando Lotes Em Aberto ..."
'DoEvents
'Call listarLote
'Me.lblStatus.Visible = False
''===========================================================================================================================
'For i = 1 To UBound(vArrayLote, 2)
'    vArrayListWiew = Array(CStr(vArrayLote(0, i)), CStr(vArrayLote(1, i)), CStr(vArrayLote(2, i)), CStr(vArrayLote(3, i)))
'    MontaListViewDados Me.lstDetalhe, vArrayListWiew
'Next
'
'End Sub
'Public Function bCheckBuscaLote() As Boolean
'
'    If Len(shtDePara.Range("UserCso")) = 0 Or _
'       Len(shtDePara.Range("PWCSO")) = 0 Then
'        MsgBox "Favor preencher as credencias de acesso!", vbInformation, "BUSCA LOTE"
'        bCheckBuscaLote = False
'        frmAcesso.Show
'        Exit Function
'    End If
'
'    If Len(Me.txtNrc) = 0 _
'        And Len(Me.txtTerminal) = 0 _
'        And Len(Me.txtLote) = 0 Then
'
'        MsgBox "Favor preencher um dos campos de busca!", vbInformation, "BUSCA LOTE"
'        bCheckBuscaLote = False
'        Exit Function
'    End If
'
'    bCheckBuscaLote = True
'End Function
'
'Private Sub UserForm_Initialize()
'    Call iniciarForm
'End Sub
'Function MontaTitulosListView(ByVal oObjListView As Object, _
'                       ByVal aTitulos As Variant, _
'                       ByVal aLarguraTitulo As Variant, _
'                       ByVal aAlinhamentoTitulo As Variant)
''+---------------------------------------------------------------------------------------------
''| ListView é parte do MSCOMCTL.OCX, porque está sendo usado ? no Office
''| vba em geral não é possível usar as ferramentas de Grid como vsflex3.ocx, msflxgrd.ocx etc..
''| então foi necessário usar o ListView da familia do TreeView pois é o mais proximo da Grid.
''+---------------------------------------------------------------------------------------------
'Dim J As Integer
'Dim strAlinhamento As ListColumnAlignmentConstants
'
'oObjListView.FullRowSelect = True  'Seleciona toda linha
'oObjListView.Gridlines = True
'oObjListView.View = lvwReport      'Estilo Grid
'oObjListView.LabelEdit = lvwManual 'Não deixa usuário editar
'
'For J = 0 To UBound(aTitulos)
'    Select Case aAlinhamentoTitulo(J)
'           Case "C": strAlinhamento = lvwColumnCenter
'           Case "E": strAlinhamento = lvwColumnLeft
'           Case "D": strAlinhamento = lvwColumnRight
'    End Select
'    oObjListView.ColumnHeaders.Add , , aTitulos(J), aLarguraTitulo(J), strAlinhamento
'Next J
'
'End Function
'
'Function MontaListViewDados(oObjListView As Object, ByVal _
'                            aDados As Variant)
'                            'aForeColor As Variant)
'
'Dim J As Integer
'Dim nPOS As Long
''+-----------------------------------------------------------------------------------------------
''| ListView é parte do MSCOMCTL.OCX, porque está sendo usado ? no Office
''| vba em geral não é possível usar as ferramentas de Grid como vsflex3.ocx, msflxgrd.ocx etc..
''| então foi necessário usar o ListView da familia do TreeView pois é o mais proximo da Grid.
''+-----------------------------------------------------------------------------------------------
''| Importante saber: Veja que o primeiro elemento é zero(0), pois é assim que funciona o ListView
''| zero(0) é a primeira coluna e 1 em diante são as demais colunas.
''+-----------------------------------------------------------------------------------------------
'
'oObjListView.ListItems.Add , , aDados(0) ', , 'SmallIcon:=aObjIcon(0)
'nPOS = oObjListView.ListItems.Count
'oObjListView.ListItems.Item(nPOS).ForeColor = vbBlue
'
'For J = 1 To UBound(aDados)
'    oObjListView.ListItems(nPOS).ListSubItems.Add , , aDados(J) ', ReportIcon:=aObjIcon(J)
'    oObjListView.ListItems(nPOS).ListSubItems(J).ForeColor = vbBlue
'
'    'Lista.ListItems.Add "index","key","text","icon","SmallIcon"
'
'Next J
'
'End Function
'Private Sub iniciarForm()
'
'    Dim vArrayTitulo    As Variant
'    Dim vArrayTamanho   As Variant
'    Dim vArrayAlinhamento As Variant
'
'    vArrayTitulo = Array("Conta", "Periodo", "Qtd", "Valor")
'    vArrayTamanho = Array("91", "91", "90", "90")
'    vArrayAlinhamento = Array("E", "E", "E", "E")
'    MontaTitulosListView Me.lstDetalhe, vArrayTitulo, vArrayTamanho, vArrayAlinhamento
'End Sub
