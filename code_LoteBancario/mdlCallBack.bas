Attribute VB_Name = "mdlCallBack"
'Dim lngLinha    As Long
'Dim sUser       As String
'Dim rng         As Range
'
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
'Dim sCnl        As String
'Dim sTerminal   As String
'Dim slote       As String
'Dim clsPw       As clsPw3270
'
'lngLinha = lngRetornaUltimaLinha(shtInsumos, "A") 'Identifica a ultima linha da planilha
'sUser = Right(frmAcesso.txtUser.Value, 7)
'
''Percorre todas as linhas preenchidas
'For Each rng In shtInsumos.Range("A12:A" & lngLinha)
'
'    'Verifica se o registro foi gerada
'    '===========================================================================
'    If Len(rng.Offset(0, 7).Value) = 0 Then
'
'        'Realiza a validação do preenchimento do registro
'        If ValidaLote(rng) Then
'
'            'Estancia da Classe Pw para retorna o CSO já logado
'            Set clsPw = New clsPw3270
'            Call gravaLog("Busca Efetivada", rng.Offset(0, 1).Value)
'            If bTesteSenha = False Then Exit For
'
'            'Informa na barra de status o andamento do processamento
'            Call StatusBar("INSERINDO LOTE BANCARIO " & rng.row - 11 & " De " & lngLinha - 11)
'
'            Application.ScreenUpdating = False
'
'            'Acessa a operação o010 para consultar Escritório
'            If acessoOperacao("C0", "O010", "O010-01", 1) Then
'                Call gravaLog("Busca Efetivada", rng.Offset(0, 1).Value)
'
'                sCnl = rng.Value 'Informacao da Coluna "A"
'                sTerminal = rng.Offset(0, 1).Value & "0" 'Terminal Coluna "B"
'                slote = rng.Offset(0, 4).Value
'
'                'Chamada da rotina para fazer a busca do Escritório
'                If efetivaBusca(sCnl, sTerminal, slote, sUser, rng.Offset(0, 7)) Then
'                    Call gravaLog("Busca Efetivada", rng.Offset(0, 1).Value)
'                    Call efetivaLote(rng, slote)
'                    Call gravaLog("Busca Efetivada", rng.Offset(0, 1).Value)
'                End If
'            Else
'                Call gravaLog("Busca Efetivada", rng.Offset(0, 1).Value)
'            End If
'        End If
'    End If
'
'    Application.ScreenUpdating = True
'    Set clsPw = Nothing
'
'    rng.Offset(0, 5).Value = frmAcesso.txtUser
'    rng.Offset(0, 6).Value = Now()
'Next rng
'
'End Sub
'Private Function ValidaLote(rng As Range) As Boolean
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
'Public Function efetivaBusca(sCnl As String, sTerminal As String, slote As String, sUser As String, rngDestino As Range) As Boolean
'
''**********************************************************************************************
'' 2.1 - Procedimento Chamado Pela Sub Busca Escritório
'' Realiza a Busca da Informação de Escritório Na Operação O010
''==================================================================
''Rotina que acessa o terminal recebido para a busca do escritório
''Rotina chamada através da rotina principal buscaEscritório
''==================================================================
'
'Dim stela As String
'
''Preenchimento da Tela O010 para realizaçao da busca do ES
'Call IntoText(18, 25, "1", 1, False)
'Call IntoText(18, 27, "T", 1, False)
'Call IntoText(18, 29, sTerminal, 1, False)
'Call IntoText(18, 40, sCnl, 1, False)
'Call IntoText(18, 74, sUser, 1, False)
'Call IntoText(19, 74, "WD", 5, True)
'
''Valida o acesso a tela de consulta
''==================================
'stela = RetTextoTela(1, 1, 80)
'
'If InStr(1, stela, "O030-01") > 0 Then
'    rngDestino.Value = "TERMINAL NAO CADASTRADO"
'    efetivaBusca = False
'ElseIf InStr(1, stela, "O020-01") > 0 Then
'    Call hllapi_enter
'    Call hllapi_wait_for_ready(5)
'
'    stela = RetTextoTela(1, 1, 80)
'    If InStr(1, stela, "O020-02") > 0 Then efetivaBusca = True
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
