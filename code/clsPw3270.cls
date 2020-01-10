VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPw3270"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private Sub Class_Initialize()
'
'Dim iTentativaLogin As Integer
'
'    If iniciaEmulador Then
'        If ConectarServidor Then
'            If TestaSenha Then
'                For iTentativaLogin = 1 To 2
'                    If Not LoginCSO(iTentativaLogin) Then
'                        bTesteSenha = False
'                    End If
'                Next iTentativaLogin
'            End If
'        End If
'    End If
'
'End Sub
'
'Public Function iniciaEmulador() As Boolean
''==========================================
''= Função inicia a emulação do PW3270
''==========================================
'On Error GoTo errConnect
'
'Dim iConectEmulador As Integer
'
'Call hllapi_disconnect
'Call hllapi_deinit
'iniciaEmulador = False
'
''status emulador = 9, não iniciado
''status emulador = 0, iniciado
'iConectEmulador = hllapi_is_connected
'If iConectEmulador = 9 Then
'    'trecho de chamada para iniciacao do emulador
'    iConectEmulador = hllapi_init("")
'    If iConectEmulador = 0 Then iniciaEmulador = True
'    Exit Function
'ElseIf iConectEmulador = 0 Then
'    iniciaEmulador = True
'    Exit Function
'End If
'
'Exit Function
'errConnect:
'    iniciaEmulador = False
'End Function
'Public Function ConectarServidor() As Boolean
''============================================
''Função Realiza a Conexao com o servidor
''============================================
'On Error GoTo errConectar
'
'    Dim iConectServidor As Integer
'    Dim iIndex          As Integer
'
'    Call hllapi_disconnect
'    Conectar = False
'
'    'Enquanto a conexão com o servidor for igual a 0
'    'Realiza a tentativa de Conexão
'    '===============================================
'    iConectServidor = 0
'    For iIndex = 1 To 5
'        If iConectServidor = 0 Then
'            iConectServidor = hllapi_connect(sServidor, 500)
'            Call hllapi_wait_for_ready(5)
'        Else
'            ConectarServidor = True
'            Exit Function
'        End If
'    Next iIndex
'
'    Exit Function
'errConectar:
'    Conectar = False
'End Function
'Public Function TestaSenha() As Boolean
'
'    Dim stela   As String
'    Dim sLogOn  As String
'
'    stela = RetTextoTela(18, 5, 7)
'
'    If stela = "Cliente" Then
'
'        'Seta usuário e senha nas posicoes
'        '========================================================
'        Call IntoText(18, 15, Left(frmAcesso.txtUser, 7), 1, False) 'Usuário
'        Call IntoText(19, 15, frmAcesso.txtPwd.Value, 5, True)
'
'        stela = Trim(RetTextoTela(22, 5, 22))
'        sLogOn = RetTextoTela(5, 6, 7)
'
'        If sLogOn = "Monitor" Then
'            TestaSenha = True
'            bTesteSenha = True
'            Call LogOut(23, 29)
'            Exit Function
'        ElseIf InStr(1, stela, "SENHA OMITIDA") > 0 Then
'            MsgBox "SENHA INCORRETA", vbExclamation, "VERIFICAR SENHA"
'            TestaSenha = False
'            bTesteSenha = False
'            Exit Function
'        ElseIf InStr(1, stela, "USUARIO NAO AUTORIZADO") > 0 Then
'            MsgBox "USUÁRIO NAO AUTORIZADO", vbExclamation, "VERIFICAR SENHA"
'            TestaSenha = False
'            bTesteSenha = False
'            Exit Function
'        End If
'    End If
'End Function
'Public Function LoginCSO(iVezes As Integer) As Boolean
'
''=====================================================================================
''Função Recebe usuário e senha e acessa o emulador Pw3270 criando uma emolução oculta
''=====================================================================================
'
'On Error GoTo ErrHandler
'
'    Dim iConnect As Long
'    Dim stela    As String
'    Dim sOp      As String
'
'
'    'Valida se a Tela Atual é a tela de Login
'    '========================================
'    stela = RetTextoTela(18, 5, 7)
'    If stela = "Cliente" Then
'
'        'Seta usuário e senha nas posicoes
'        '========================================================
'        Call IntoText(18, 15, Left(frmAcesso.txtUser.Value, 7), 1, False) 'Usuário
'        Call IntoText(19, 15, frmAcesso.txtPwd.Value, 5, True)            'Senha
'
'        '=========================================================
'
'        'Valida se o login foi realizado
'        '==========================================================
'
'        'Verifica se é a tela inicial após o login
'        stela = RetTextoTela(5, 6, 7)
'
'        'Verifica se existe processo preso na seção
'        'sOp = Trim(RetTextoTela(7, 1, 80))
'
'        If iVezes = 1 Then
'            If stela = "Monitor" Then 'True para logon sem seção presa
'                Call LogOut(23, 29)
'                LoginCSO = True
'                Exit Function
'            Else
'                ConectarServidor
'                LoginCSO = False
'            End If
'        Else
'            If stela = "Monitor" Then
'                LoginCSO = True
'                Exit Function
'            End If
'        End If
'    End If
'
'    Exit Function
'
'ErrHandler:
'    LoginCSO = False
'End Function
'
''Private Sub Class_Terminate()
''    Call hllapi_disconnect
''    Call hllapi_deinit
''End Sub
