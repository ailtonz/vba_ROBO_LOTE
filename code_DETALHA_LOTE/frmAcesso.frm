VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAcesso 
   Caption         =   "Digite usu�rio e senha CSO"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4500
   OleObjectBlob   =   "frmAcesso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Private Sub cmdOk_Click()
''====================================================
''Chamada do Inicio do Processamento
''====================================================
'
'On Error GoTo ErrHandler
''Valida o preenchimento de usu�rio e senha
''=============================================
'Application.ScreenUpdating = False
'If Validacao Then
'
'    shtDePara.Range("UserCso") = Me.txtUser.Value
'    shtDePara.Range("PWCSO") = Me.txtPwd.Value
'
''    Me.Hide
''
''    'Inicia o processo buscando o numero de CNPJ e numero de Operadora
''    Call Main
'
'End If
'Application.ScreenUpdating = True
'Unload Me
'Call StatusBar("")
'
''MsgBox "Rob� Executado Com Sucesso!", vbInformation, "Segmento �nico"
'
'Exit Sub
'ErrHandler:
'    Application.ScreenUpdating = True
'End Sub
'
'Private Function Validacao() As Boolean
''========================================================
''Fun��o de Valida��o de preenchimento de usu�rio e senha
''========================================================
'
'    'Valida usu�rio
'    If Len(Me.txtUser.Value) <> 8 Then
'        MsgBox "Usu�rio inv�lido, usu�rio deve ter 8 caracteres!", vbCritical, "Erro!"
'        Validacao = False
'        Exit Function
'    End If
'
'    'Valida e Senha
'    If Len(Me.txtPwd.Value) <> 8 Then
'        MsgBox "Senha inv�lida, a senha deve ter 8 caracteres!", vbCritical, "Erro!"
'        Validacao = False
'        Exit Function
'    End If
'
'    Validacao = True
'End Function
'
'Private Sub UserForm_Initialize()
'    Me.txtUser.Value = shtDePara.Range("UserCso").Value
'    Me.txtPwd.Value = shtDePara.Range("PWCSO").Value
'End Sub
'
