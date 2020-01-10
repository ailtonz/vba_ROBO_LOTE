Attribute VB_Name = "mLote"
'Public nLote        As String
'Public vArrayLote   As Variant
'Public Sub buscaLoteO010(s1_O As String _
'                         , s2_T As String _
'                         , s3_N As String _
'                         , s4_L As String _
'                         , s5_U As String _
'                         , s6_E As String)
'
'Dim iLinha As Integer
''=====================================================
''s1 = 1 - Tipo de Consulta (Cadastro)
''s2 = T ou N Tipo de pesquisa (Terminal ou NRC)
''s3 = Numero do Terminal ou NRC
''s4 = L Localidade se o tipo de consulta for Terminal
''s5 = U numero do usuário
''s6 = Escritório (WD ou ES)
''=====================================================
'
''Preenchimento da Tela O010 para realizaçao da busca do Numero Lote
''======================================================================
'Call IntoText(18, 25, s1_O, 1, False)
'Call IntoText(18, 27, s2_T, 1, False)
'Call IntoText(18, 29, s3_N, 1, False)
'Call IntoText(18, 40, s4_L, 1, False)
'Call IntoText(18, 74, s5_U, 1, False)
'Call IntoText(19, 74, s6_E, 5, True)
''=======================================================================
'
''Quando aparecer a lista de terminais, deve ser escolhido a ultima linha
''=======================================================================
'If InStr(1, RetTextoTela(1, 1, 80), "O122") > 0 Then
'    iLinha = retUltLinha
'    Call IntoText(iLinha, 2, "X", 1, False)
'    Call hllapi_wait_for_ready(2)
'    Call hllapi_pfkey(1)
'    Call hllapi_wait_for_ready(2)
'End If
''=======================================================================
'
''Retorna o numero de Lote
''=====================================
'Call hllapi_enter
'Call hllapi_wait_for_ready(2)
'Call gravaLog("RC", "LOTE1")
'
'nLote = Trim(RetTextoTela(4, 21, 10))
''=====================================
'
'End Sub
'Public Function retUltLinha() As Integer
'
'    Dim iLinha As Integer
'
'    For iLinha = 3 To 22
'        If Len(Trim(RetTextoTela(iLinha, 3, 5))) = 0 Then
'            retUltLinha = iLinha - 1
'            Exit Function
'        End If
'    Next iLinha
'End Function
'Public Function listarLote() As Boolean
'
'    Dim iTotalPag As Integer
'    Dim iIndex    As Integer
'    Dim i         As Integer
'    Dim Pw3270    As New clsPw3270
'
'    Set Pw3270 = New clsPw3270
'
'    'Tratamento Erro operação FVIG
'    '=============================
'    If Not acessoOperacao("RC", "FVIG", "FVIG01", 1) Then MsgBox "Erro no Acceso a transação FVIG!", vbInformation, "BUSCA LOTE": Exit Function
'    '=============================
'
'    'Entrada na tela de busca de lote
'    '====================================
'    Call IntoText(19, 28, "1", 1, False)
'    Call IntoText(19, 34, "4", 1, True)
'    '====================================
'
'    'Valida a entrada da tela
'    '===========================================================================================
'    If Not InStr(1, RetTextoTela(1, 1, 80), "FVID01") > 0 Then listarLote = False: Exit Function
'    '===========================================================================================
'
'    'Preechimento das Informações de Lote
'    '==============================================
'    Call IntoText(4, 12, Left(nLote, 3), 1, False)
'    Call IntoText(4, 28, Mid(nLote, 4, 4), 1, False)
'    Call IntoText(4, 68, "XXXX", 1, True)
'    '==============================================
'
'    'Total de paginas a percorrer
'    '==============================================
'    iTotalPag = CInt(RetTextoTela(2, 79, 2))
'    '==============================================
'
'    Call gravaLog("RC", "FVID")
'    ReDim vArrayLote(0 To 4, 0 To 0)
'    For iIndex = 1 To iTotalPag
'        For i = 8 To 21
'            If Len(Trim(RetTextoTela(i, 4, 1))) = 0 And Len(Trim(RetTextoTela(i, 14, 7))) > 0 Then
'                ReDim Preserve vArrayLote(0 To 4, 0 To UBound(vArrayLote, 2) + 1)
'                vArrayLote(0, UBound(vArrayLote, 2)) = RetTextoTela(i, 14, 7)
'                vArrayLote(1, UBound(vArrayLote, 2)) = RetTextoTela(i, 23, 2)
'                vArrayLote(2, UBound(vArrayLote, 2)) = Trim(RetTextoTela(i, 53, 3))
'                vArrayLote(3, UBound(vArrayLote, 2)) = Trim(RetTextoTela(i, 57, 18))
'            End If
'        Next i
'        Call hllapi_enter
'        Call hllapi_wait_for_ready(1)
'    Next iIndex
'End Function
'Public Sub listarTerminaisLote(sBanco As String _
'                               , sAgencia As String _
'                               , sConta As String _
'                               , sPer As String _
'                               , sht As Worksheet)
'
'Dim iPag        As Integer
'Dim iIndex      As Integer
'Dim ultLinha    As Long
'Dim iLinha      As Integer
'
''Iniciar busca dos terminais por Lote
''=======================================
'Call IntoText(4, 10, sBanco, 1, False)
'Call IntoText(4, 26, sAgencia, 1, False)
'Call IntoText(4, 42, sConta, 1, False)
'Call IntoText(4, 61, sPer, 1, True)
''=======================================
'Call gravaLog("RC", "FVID")
'
'iPag = CInt(RetTextoTela(2, 78, 3))
'    For iIndex = 1 To iPag
'        For iLinha = 7 To 21
'            If Len(Trim(RetTextoTela(iLinha, 2, 17))) > 0 Then
'                If Not Replace(RetTextoTela(iLinha, 2, 17), ".", "") = "99999999999999" Then
'                    ultLinha = lngRetornaUltimaLinha(sht, "A") + 1
'                    sht.Cells(ultLinha, 1).NumberFormat = "@": sht.Cells(ultLinha, 1) = sBanco
'                    sht.Cells(ultLinha, 2).NumberFormat = "@": sht.Cells(ultLinha, 2) = sAgencia
'                    sht.Cells(ultLinha, 3).NumberFormat = "@": sht.Cells(ultLinha, 3) = sConta
'                    sht.Cells(ultLinha, 4).NumberFormat = "@": sht.Cells(ultLinha, 4) = sPer
'                    sht.Cells(ultLinha, 5).NumberFormat = "@": sht.Cells(ultLinha, 5) = Replace(RetTextoTela(iLinha, 2, 17), ".", "")
'                    sht.Cells(ultLinha, 6).NumberFormat = "@": sht.Cells(ultLinha, 6) = Trim(RetTextoTela(iLinha, 25, 11))
'                    sht.Cells(ultLinha, 7).NumberFormat = "@": sht.Cells(ultLinha, 7) = IIf(RetTextoTela(iLinha, 20, 1) = "P", "PAGO", "")
'                End If
'            End If
'
'            If Len(Trim(RetTextoTela(iLinha, 47, 17))) > 0 Then
'                If Not Replace(RetTextoTela(iLinha, 47, 17), ".", "") = "99999999999999" Then
'                    ultLinha = lngRetornaUltimaLinha(sht, "A") + 1
'                    sht.Cells(ultLinha, 1).NumberFormat = "@": sht.Cells(ultLinha, 1) = sBanco
'                    sht.Cells(ultLinha, 2).NumberFormat = "@": sht.Cells(ultLinha, 2) = sAgencia
'                    sht.Cells(ultLinha, 3).NumberFormat = "@": sht.Cells(ultLinha, 3) = sConta
'                    sht.Cells(ultLinha, 4).NumberFormat = "@": sht.Cells(ultLinha, 4) = sPer
'                    sht.Cells(ultLinha, 5).NumberFormat = "@": sht.Cells(ultLinha, 5) = Replace(RetTextoTela(iLinha, 47, 17), ".", "")
'                    sht.Cells(ultLinha, 6).NumberFormat = "@": sht.Cells(ultLinha, 6) = Trim(RetTextoTela(iLinha, 69, 11))
'                    sht.Cells(ultLinha, 7).NumberFormat = "@": sht.Cells(ultLinha, 7) = IIf(RetTextoTela(iLinha, 65, 1) = "P", "PAGO", "")
'                End If
'            End If
'        Next iLinha
'        Call hllapi_enter
'        Call hllapi_wait_for_ready(5)
'    Next iIndex
'End Sub
'Public Function lngRetornaUltimaLinha(sht As Worksheet, sColuna As String) As Long
'    lngRetornaUltimaLinha = sht.Range(sColuna & 65000).End(xlUp).row
'End Function
