Attribute VB_Name = "mSeg"
Global liberado As Boolean

Public Function lockPlanilha(senha As String, planilha As String)
    Dim sht As Worksheet
    Dim Range As String
    
    Select Case planilha
    
        Case "DATEL"
            Range = "A:T"
        
        Case "FVIG"
            Range = "A:I"
        
        Case "FCTA"
            Range = "A:K"
    
    End Select
    
    If liberado Then
        liberado = False
    End If
    
    Set sht = ThisWorkbook.Worksheets(planilha)
    
    sht.Range(Range).Locked = False
    
    sht.Protect senha

End Function

Public Function unlockPlanilha(senha As String, planilha As String)
    Dim sht As Worksheet
    
    Set sht = ThisWorkbook.Worksheets(planilha)
    
    sht.Unprotect senha

End Function

Public Function validaSenha(senha As String, planilha As String)
    
    Dim sht As Worksheet
    Dim sorginal As String
    
    sorginal = shtDePara.Cells(1, 10).Value
     
    If senha = soriginal Then
        Call unlockPlanilha(senha, planilha)
        liberado = True
    Else
        MsgBox "Senha de Desbloqueio Invalida", vbOKOnly, "Senha Invalida"
    End If
    
End Function
