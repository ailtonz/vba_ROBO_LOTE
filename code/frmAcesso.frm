VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAcesso 
   Caption         =   "Digite usu�rio e senha CSO"
   ClientHeight    =   2295
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
'
'
'Private Sub cmdOk_Click()
''====================================================
''Chamada do Inicio do Processamento
''====================================================
'
''Valida o preenchimento de usu�rio e senha
''=============================================
'If Validacao Then
'
'    Me.Hide
'
'    'Faz a Inclus�o de dados no formul�rio da Opera��o OP01
'    Call InsereO010
'
'End If
'
'Unload Me
'Call StatusBar("")
'
'MsgBox "Rob� Executado Com Sucesso!", vbInformation, "Segmento �nico"
'End Sub
'
'Private Function Validacao() As Boolean
''========================================================
''Fun��o de Valida��o de preenchimento de usu�rio e senha
''========================================================
'
'    'Valida usu�rio
'    If Len(Me.txtUser) <> 8 Then
'        MsgBox "Usu�rio inv�lido, usu�rio deve ter 8 caracteres!", vbCritical, "Erro!"
'        Validacao = False
'        Exit Function
'    End If
'
'    'Valida e Senha
'    If Len(Me.txtPwd) <> 8 Then
'        MsgBox "Senha inv�lida, a senha deve ter 8 caracteres!", vbCritical, "Erro!"
'        Validacao = False
'        Exit Function
'    End If
'
'     'Valida e Senha
'    If Len(Me.txtSenhaEmissao) <> 5 Then
'        MsgBox "Senha de emiss�o inv�lida, a senha deve ter 5 caracteres!", vbCritical, "Erro!"
'        Validacao = False
'        Exit Function
'    End If
'
'    Validacao = True
'End Function
'
''Private Sub UserForm_Initialize()
'''Rotina para teste
''    Me.txtUser.Value = "e3817920"
''    Me.txtPwd.Value = "gabi3838"
''    Me.txtSenhaEmissao.Value = "*gpxt"
''End Sub
