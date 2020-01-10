Attribute VB_Name = "mdlCallBack"
'Dim lngLinha        As Long
'Public sUser        As String
'Public sPwd         As String
'Public sEmissao     As String
'Dim rng             As Range
'Public Sub Form_show()
'    frmAcesso.Show vbModeless
'End Sub
'
'Public Sub InsereO010()
'
''**********************************************************************************************
'' 1 - Procedimento que acessa O010
''==============================================================================================
''Rotina principal chamada do form
''Rotina para buscar código de escritório na operação O010
''==============================================================================================
'
'Dim arrayDados  As Variant
'Dim sTerminal   As String
'Dim slote       As String
'Dim clsPw       As clsPw3270
'Dim lngContar   As Long
'Dim lngContador As Long
'
'lngLinha = lngRetornaUltimaLinha(shtInsumos, "E") 'Identifica a ultima linha da planilha
'lngContar = lngLinha - 11 - Application.WorksheetFunction.CountA(shtInsumos.Range("I12:I" & lngLinha))
'
'sUser = Right(frmAcesso.txtUser.Value, 7)
'sPwd = frmAcesso.txtPwd.Value
'sEmissao = frmAcesso.txtSenhaEmissao.Value
'
'lngContador = 1
''Percorre todas as linhas preenchidas
'For Each rng In shtInsumos.Range("A12:A" & lngLinha)
'
'    'Verifica se o registro foi gerada
'    '===========================================================================
'    If Len(rng.Offset(0, 7).Value) = 0 Then
'
'        Application.ScreenUpdating = False
'        Call StatusBar("PROCESSAMENTO " & lngContador & " DE " & lngContar)
'
'        arrayDados = arrayConsulta(rng)
'        If UBound(arrayDados) > 0 Then
'
'            'Estancia da Classe Pw para retorna o CSO já logado
'            '==================================================
'            Set clsPw = New clsPw3270
'            '==================================================
'
'            'Acessa a operação o010 para consultar Escritório
'            '==================================================
'            If acessoOperacao("C0", "O010", "O010-01", 1) Then
'                Call efetivaBusca(arrayDados, rng.Offset(0, 8))
'            End If
'            '==================================================
'        End If
'        Application.ScreenUpdating = True
'        lngContador = lngContador + 1
'        ThisWorkbook.Save
'    End If
'Next rng
'
'End Sub
'Private Function arrayConsulta(rng As Range) As Variant
'
'Dim arrayRetorno As Variant
''arrayRetorno
''0 - O (1)
''1 - T (N NRC, T TERMINAL, L 0800)
''2 - NUMERO (NRC OU TERMINAL)
''3 - LOCAL (CNL)
''4 - NUMERO DE LOTE
''5 - INCLUSAO E EXCLUSAO
'
''Consulta por NRC
''==================================================
'If Len(rng.Value) > 0 Then
'    If Len(rng.Value) = 10 Then
'        ReDim arrayRetorno(0 To 5)
'        arrayRetorno(0) = "1"
'        arrayRetorno(1) = "N"
'        arrayRetorno(2) = rng.Value
'        arrayRetorno(3) = "00000"
'        arrayRetorno(4) = rng.Offset(0, 4).Value
'        arrayRetorno(5) = rng.Offset(0, 5).Value
'    Else
'        rng.Offset(0, 7).Value = "NUMERO NRC INVÁLIDO"
'        ReDim arrayRetorno(0 To 0)
'    End If
'Else
'    'Consulta Por Terminal
'    '=====================
'    If Len(rng.Offset(0, 1).Value) = 5 And _
'       Len(rng.Offset(0, 2).Value) >= 8 Then
'        If Right(rng.Offset(0, 2).Value, 2) = "80" Then
'            ReDim arrayRetorno(0 To 5)
'            arrayRetorno(0) = "1"
'            arrayRetorno(1) = "L"
'            arrayRetorno(2) = "800" & replace0(rng.Offset(0, 2).Value)
'            arrayRetorno(3) = Format(CInt(Left(rng.Offset(0, 1).Value, 3)), "00000")
'            arrayRetorno(4) = rng.Offset(0, 4).Value
'            arrayRetorno(5) = rng.Offset(0, 5).Value
'        ElseIf Right(rng.Offset(0, 2).Value, 2) = "30" Then
'            ReDim arrayRetorno(0 To 5)
'            arrayRetorno(0) = "1"
'            arrayRetorno(1) = "L"
'            arrayRetorno(2) = "300" & replace0(rng.Offset(0, 2).Value)
'            arrayRetorno(3) = Format(CInt(Left(rng.Offset(0, 1).Value, 3)), "00000")
'            arrayRetorno(4) = rng.Offset(0, 4).Value
'            arrayRetorno(5) = rng.Offset(0, 5).Value
'        Else
'            ReDim arrayRetorno(0 To 5)
'            arrayRetorno(0) = "1"
'            arrayRetorno(1) = "T"
'            arrayRetorno(2) = rng.Offset(0, 2).Value & "0"
'            arrayRetorno(3) = rng.Offset(0, 1).Value
'            arrayRetorno(4) = rng.Offset(0, 4).Value
'            arrayRetorno(5) = rng.Offset(0, 5).Value
'        End If
'    Else
'        rng.Offset(0, 7).Value = "NUMERO CNL E/OU TERMINAL INVÁLIDO"
'        ReDim arrayRetorno(0 To 0)
'    End If
'End If
'arrayConsulta = arrayRetorno
'End Function
'Private Function replace0(sTerminal As String) As String
'
'Do While Left(sTerminal, 1) = 0
'    sTerminal = Right(sTerminal, Len(sTerminal) - 1)
'Loop
'replace0 = sTerminal
'End Function
'Private Function ValidaLote(rng As Range) As Boolean
'
'    'Valida NRC
'    '===================================================
'    If Len(rng.Value) <> 10 And _
'       (Len(rng.Offset(0, 1).Value <> 5) Or Len(rng.Offset(0, 2).Value) <> 8) Then
'        rng.Offset(0, 7).Value = "NUMERO LOCAL INVALIDO"
'        ValidaLote = False
'        Exit Function
'    End If
'
'
'    'Valida Cnl
'    If Len(rng.Value) <> 5 Then
'        rng.Offset(0, 5).Value = "NUMERO LOCAL INVALIDO"
'        ValidaLote = False
'        Exit Function
'    End If
'
'    'Valida Terminal
'    If Len(rng.Offset(0, 1).Value) <> 8 Then
'        rng.Offset(0, 5).Value = "NUMERO TERMINAL INVALIDO"
'        ValidaLote = False
'        Exit Function
'    End If
'
'
'    'Valida Lote
'    If Len(rng.Offset(0, 4).Value) <> 8 Then
'        rng.Offset(0, 5).Value = "NUMERO LOTE INVALIDO"
'        ValidaLote = False
'        Exit Function
'    End If
'
'    ValidaLote = True
'End Function
'
'Public Function efetivaBusca(arrayLista As Variant, rngDestino As Range)
'
''**********************************************************************************************
'' 2.1 - Procedimento Chamado Pela Sub Busca Escritório
'' Realiza a Busca da Informação de Escritório Na Operação O010
''==================================================================
''Rotina que acessa o terminal recebido para a busca do escritório
''Rotina chamada através da rotina principal buscaEscritório
''==================================================================
'
'Dim stela   As String
'Dim sTexto  As String
'Dim sOpcao  As String
'Dim i       As Integer
'
'sTexto = "INCLUSÃO ATRAVES DE ROBO - USUARIO: " & frmAcesso.txtUser & " " & Now()
'
''Preenchimento da Tela O010 para realizaçao da busca do ES
'Call IntoText(18, 25, CStr(arrayLista(0)), 1, False)
'Call IntoText(18, 27, CStr(arrayLista(1)), 1, False)
'Call IntoText(18, 29, CStr(arrayLista(2)), 1, False)
'Call IntoText(18, 40, CStr(arrayLista(3)), 1, False)
'Call IntoText(18, 74, sUser, 1, False)
'Call IntoText(19, 74, "WD", 1, False)
'Call IntoText(20, 74, sEmissao, 5, True)
'
''Valida o acesso a tela de consulta
''==================================
'stela = RetTextoTela(1, 1, 80)
'
'If InStr(1, stela, "O030-01") > 0 Then
'    rngDestino.Value = "TERMINAL NAO CADASTRADO"
'    efetivaBusca = False
'ElseIf RetTextoTela(1, 60, 4) = "O122" Then
'    i = ValidaReg
'    Call IntoText(i, 2, "X", 1, False)
'    Call hllapi_wait_for_ready(5)
'    Call hllapi_pfkey(1)
'    Call hllapi_wait_for_ready(5)
'End If
'
'stela = RetTextoTela(1, 1, 80)
'If InStr(1, stela, "O020-01") > 0 Then
'
'    'Caso for Inclusão/Alteração
'    '===========================
'    If arrayLista(5) = "I" Then
'
'        'Paginação
'        Call hllapi_enter
'        Call hllapi_wait_for_ready(5)
'
'        'Verifica se a tela navegou
'        '=======================================
'        stela = RetTextoTela(1, 1, 80)
'        If InStr(1, stela, "O020-02") > 0 Then
'
'            Call IntoText(4, 21, CStr(arrayLista(4)), 1, False) 'Preenche Lote
'            Call IntoText(4, 48, "8", 1, False) 'Preenche Opcao
'            Call IntoText(18, 16, sTexto, 5, True) 'Preenche Obs
'            Call hllapi_enter
'            Call hllapi_enter
'            Call hllapi_wait_for_ready(5)
'            'Confirma Preenchimento
'            '====================================================
'            If RetTextoTela(23, 19, 3) = "PF1" Then
'                Call hllapi_pfkey(1)
'                Call hllapi_wait_for_ready(5)
'                rngDestino.Offset(0, -2) = frmAcesso.txtUser.Value
'                rngDestino.Offset(0, -1).Value = Now()
'                rngDestino.Value = "INCLUIDO COM SUCESSO!"
'            Else
'                rngDestino.Offset(0, -2).Value = frmAcesso.txtUser.Value
'                rngDestino.Offset(0, -1).Value = Now()
'                rngDestino.Value = "ERRO DE INCLUSÃO, VERIFICAR!"
'            End If
'            '====================================================
'        Else
'            rngDestino.Offset(0, -2).Value = frmAcesso.txtUser.Value
'            rngDestino.Offset(0, -1).Value = Now()
'            rngDestino.Value = Trim(RetTextoTela(23, 1, 40))
'        End If
'
'    ElseIf arrayLista(5) = "E" Then
'        If Len(Trim(RetTextoTela(21, 2, 79))) = 0 Then sOpcao = 3 Else: sOpcao = 4
'
'        'Paginação
'        Call hllapi_enter
'        Call hllapi_wait_for_ready(5)
'
'        Call IntoText(4, 21, String(8, " "), 1, False)
'        Call IntoText(4, 48, sOpcao, 1, False)
'
'        Call IntoText(18, 16, sTexto, 5, True) 'Preenche Obs
'        Call hllapi_enter
'
'        'Confirma Preenchimento
'        '====================================================
'        If RetTextoTela(23, 19, 3) = "PF1" Then
'            Call hllapi_pfkey(1)
'            Call hllapi_wait_for_ready(5)
'            rngDestino.Offset(0, -2).Value = frmAcesso.txtUser.Value
'            rngDestino.Offset(0, -1).Value = Now()
'            rngDestino.Value = "INCLUIDO COM SUCESSO!"
'        Else
'            rngDestino.Offset(0, -2).Value = frmAcesso.txtUser.Value
'            rngDestino.Offset(0, -1).Value = Now()
'            rngDestino.Value = "ERRO DE INCLUSÃO, VERIFICAR!"
'        End If
'        '====================================================
'
'    Else
'        rngDestino.Offset(0, -2).Value = frmAcesso.txtUser.Value
'        rngDestino.Offset(0, -1).Value = Now()
'        rngDestino.Value = "ERRO DE INCLUSÃO, VERIFICAR!"
'    End If
'End If
'End Function
'Public Sub efetivaLote(rng As Range, slote As String)
'
'Dim stela As String
'
'    Call IntoText(4, 21, slote, 5, False)
'    Call IntoText(4, 48, "8", 5, True)
'    Call hllapi_enter
'    Call hllapi_wait_for_ready(5)
'    Call hllapi_enter
'    Call hllapi_wait_for_ready(5)
'
'    stela = RetTextoTela(23, 1, 80)
'    Call gravaLog("Busca Efetivada", "Teste")
'    If InStr(1, stela, "O0200091") Then
'        Call hllapi_pfkey(1)
'        Call hllapi_wait_for_ready(5)
'        rng.Offset(0, 7).Value = "OK"
'    Else
'        rng.Offset(0, 7).Value = "PROCESSADO ANTERIORMENTE"
'    End If
'End Sub
'
'Public Function lngRetornaUltimaLinha(sht As Worksheet, sColuna As String) As Long
'    lngRetornaUltimaLinha = sht.Range(sColuna & 65000).End(xlUp).row
'End Function
'Public Sub sleep(iSec As Integer)
'    Application.wait (Now + TimeValue("0:00:" & iSec))
'End Sub
'Public Sub ApagarDados()
'    'Apaga todos os dados da planilha
'    '==============================================================================================
'    If MsgBox("Confirma a exclusão dos dados da planilha?", vbYesNo, "Apagar Dados!") = vbYes Then
'        lngLinha = lngRetornaUltimaLinha(shtInsumos, "A") 'Identifica a ultima linha da planilha
'        shtInsumos.Range("A12:H" & lngLinha).ClearContents
'    End If
'    '===============================================================================================
'End Sub
'Public Function ValidaReg() As Integer
'
'    Dim iLinha As Integer
'
'    For iLinha = 3 To 22
'        If Len(Trim(RetTextoTela(iLinha, 3, 5))) = 0 Then
'            ValidaReg = iLinha - 1
'            Exit Function
'        End If
'    Next iLinha
'End Function
