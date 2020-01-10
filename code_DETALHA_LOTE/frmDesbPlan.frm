VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDesbPlan 
   Caption         =   "Desbloquear Planilha"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmDesbPlan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDesbPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancelar_Click()
    txtSenhaDesb.text = ""
    frmDesbPlan.Hide
End Sub

Private Sub btnLiberar_Click()
    If Not IsEmpty(txtSenhaDesb.text) Or txtSenhaDesb.text <> "" Then
        Call validaSenha(txtSenhaDesb.text, ActiveSheet.Name)
        MsgBox "Planilha desbloqueada, Favor bloquea-la Novamente após o termino da edição", vbOKOnly, "Desbloqueado com Sucesso"
        frmDesbPlan.Hide
    Else
        MsgBox "Favor digitar a Senha", vbOKOnly, "Senha Vazia"
    End If
End Sub

Private Sub UserForm_Activate()
    txtSenhaDesb.text = ""
    lblPlanilha.Caption = ActiveSheet.Name
End Sub



