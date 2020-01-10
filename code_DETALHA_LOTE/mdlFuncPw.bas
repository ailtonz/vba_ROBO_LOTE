Attribute VB_Name = "mdlFuncPw"
'Declare Function hllapi_init Lib "libhllapi.dll" (ByVal tp As String) As Long
'Declare Function hllapi_deinit Lib "libhllapi.dll" () As Long
'Declare Function hllapi_get_revision Lib "libhllapi.dll" () As Long
'Declare Function hllapi_connect Lib "libhllapi.dll" (ByVal uri As String, ByVal wait As Integer) As Long
'Declare Function hllapi_disconnect Lib "libhllapi.dll" () As Long
'Declare Function hllapi_wait_for_ready Lib "libhllapi.dll" (ByVal timeout As Integer) As Long
'Declare Function hllapi_get_screen_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
'Declare Function hllapi_enter Lib "libhllapi.dll" () As Long
'Declare Function hllapi_get_message_id Lib "libhllapi.dll" () As Long
'Declare Function hllapi_set_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
'Declare Function hllapi_wait Lib "libhllapi.dll" (ByVal timeout As Integer) As Long
'Declare Function hllapi_pfkey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
'Declare Function hllapi_pakey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
'Declare Function hllapi_cmp_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
'Declare Function hllapi_is_connected Lib "libhllapi.dll" () As Long
''Public Const sServidor = "10.20.5.2:23"
'Public Const sServidor = "10.20.1.63:23"
'Public bTesteSenha As Boolean
'Public Sub IntoText(iRow As Integer, iCol As Integer, sTexto As String, iWait As Integer, bEnter As Boolean)
'    Call hllapi_set_text_at(iRow, iCol, sTexto)
'    Call hllapi_set_text_at(iRow, iCol, sTexto)
'    Call hllapi_wait_for_ready(iWait)
'
'    If bEnter = True Then
'        Call hllapi_enter
'        Call hllapi_wait_for_ready(iWait)
'    End If
'End Sub
'Public Function CopiaTela(sProcesso As String) As Variant
'
'    Dim iIndex      As Integer
'    Dim sSepara     As String
'    Dim aList(28)   As Variant
'    Dim vTela       As Variant
'
'    vTela = RetvTela
'
'    For iIndex = 1 To 80
'        sSepara = sSepara & "-"
'    Next iIndex
'
'    aList(1) = sSepara
'    aList(2) = sProcesso
'
'    For iIndex = 1 To 24
'        aList(iIndex + 2) = vTela(iIndex)
'    Next iIndex
'    aList(28) = sSepara & vbCrLf
'
'    CopiaTela = aList
'End Function
'Public Sub LogOut(iRow As Integer, iCol As Integer)
'    Call hllapi_set_text_at(iRow, iCol, "/F")
'    Call hllapi_enter
'    Call hllapi_wait_for_ready(1)
'End Sub
'Public Function RetvTela() As Variant
''======================================================================
''Função Retorna a Posição da tela do CSO conforme parametros informado
''======================================================================
'
'    Dim sTexto As String
'    Dim iIndex As Integer
'    Dim vTela(24) As Variant
'    Dim iValida As Integer
'
'    On Error GoTo ErrHandler
'    sTexto = Space(2000)
'
'    Call hllapi_get_screen_at(1, 1, sTexto)
'    iValida = 1
'    For iIndex = 1 To 24
'        vTela(iIndex) = Mid(sTexto, iValida, 80) & vbCr
'        iValida = iValida + 80
'    Next iIndex
'
'    RetvTela = vTela
'    Exit Function
'
'ErrHandler:
'    RetvTela = Space(0)
'End Function
'Public Function RetTextoTela(iRow As Integer, iColuna As Integer, iTamanho) As String
''======================================================================
''Função Retorna a Posição da tela do CSO conforme parametros informado
''======================================================================
'
'    Dim sTexto As String
'    Dim iIndex As Integer
'
'    On Error GoTo ErrHandler
'    sTexto = Space(iTamanho)
'    Call hllapi_get_screen_at(iRow, iColuna, sTexto)
'
'    RetTextoTela = sTexto
'    Exit Function
'
'ErrHandler:
'    RetTextoTela = Space(0)
'End Function
'Public Function acessoOperacao(sPasso1 As String, sPasso2 As String, _
'                              sValida As String, iValida As Integer) As Boolean
'
''==============================================================================
''Variável sPasso1 para a entrada da 1 tela após o login do cso
''Variável sPasso2 para a entrada da tela da referida operação
''Variável sValida para que seja validado se a tela da operação foi carregada
''Variável iValida recebe o numero da linha aonde o texto deve estar
''==============================================================================
'
'On Error GoTo ErrHandler
'
'    Dim stela   As String
'    Dim vValida As Variant
'    Dim iIndex  As Integer
'
'    'Primeiro preenchimento
'    Call IntoText(23, 29, sPasso1, 5, True)
'    Call hllapi_enter
'    Call hllapi_wait_for_ready(5)
'    Call gravaLog("RC", "PASSO1")
'
'    'Segundo Preenchimento
'    Call IntoText(24, 41, sPasso2, 5, True)
'    Call gravaLog("RC", "PASSO2")
'
'    'Busca a informação da Tela conforme parametro informado da Linha
'    stela = RetTextoTela(iValida, 1, 80)
'
'    vValida = Split(sValida, ";")
'    For iIndex = 0 To UBound(vValida)
'
'        'Valida se a tela está correta
'        If InStr(1, stela, vValida(iIndex)) > 0 Then
'            acessoOperacao = True
'            Exit Function
'        Else
'            acessoOperacao = False
'        End If
'    Next iIndex
'
'    Exit Function
'
'ErrHandler:
'    acessoOperacao = False
'    'Call gravaLog("Erro acessoOperacao!")
'    'MsgBox Err.Number & vbCrLf & Err.Description
'End Function
'
'Public Function PrintTela()
'
'    Dim iIndex      As Integer
'    Dim stela       As String
'    Dim vTela       As Variant
'
'    vTela = RetvTela
'
'    For iIndex = 0 To 24
'        stela = stela & vTela(iIndex) & vbCrLf
'    Next iIndex
'
'    frmTelaPw.lblTexto.Caption = stela
'    DoEvents
'End Function
