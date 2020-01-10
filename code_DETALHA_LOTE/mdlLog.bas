Attribute VB_Name = "mdlLog"
'Public Const bLog = False
'Public Sub geraTxtFCTA()
'
'Dim iIndex      As Integer
'Dim iServ       As Integer
'Dim vServ       As Variant
'Dim sFolder     As String
'Dim sFile       As String
'Dim vArq        As Variant
'Dim iCont       As Integer
'Dim iContServ   As Integer
'Dim vLista      As Variant
'Dim sHeader1    As String
'
''Define Caminho do Diretorio a ser Armazenado o arquivo txt
''===============================================================
'sFolder = ThisWorkbook.Path & "\2Via_" & Replace(Date, "/", "")
''===============================================================
''Cria a pasta se não existir
''===============================================================
'If Not FileExists(sFolder, True) Then Call MkDir(sFolder)
''===============================================================
'
''Busca os Dados Gerados e armazena em Array
''==========================================
'vLista = armazenaTerminais
''==========================================
'
''Percorre a lista e Gera um Arquivo para cada linha
''==================================================
'sFile = sFolder & "\FCTA_" & Replace(Date, "/", "") & "_" & Replace(Time, ":", "") & ".txt"
'ReDim vArq(0 To 0)
'iCont = 0
'
''Header 0 - Data e Hora da Geração do Arquivo
''====================================================================
'vArq(iCont) = "0" & Format(Date, "DD/MM/YYYY") & Format(Time, "HH:MM:SS")
'iCont = iCont + 1
''====================================================================
'
'For iIndex = 0 To UBound(vLista) - 1
'
'    '=======================================================================================================================
'    ReDim Preserve vArq(0 To iCont)
'
'    If vLista(iIndex, 11) = "GERAR TXT" Then
'    'Header 1 - Layout Conforme Versão 03
'    '====================================================================
'    ReDim Preserve vArq(0 To iCont)
'    sHeader1 = 1 'Header 1 Caracter
'    sHeader1 = sHeader1 & vLista(iIndex, 17) 'Numero NF 20 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 21) 'NRC 11 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 13) 'Nome do Cliente 250 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 1)  'Local 5 Caracteres
'    sHeader1 = sHeader1 & Left(vLista(iIndex, 2), 4) & "-" & Right(vLista(iIndex, 2), 5) 'Terminal 10 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 22) 'DV 1 Caracter
'    sHeader1 = sHeader1 & vLista(iIndex, 19) 'TA 3 Caracteres
'    sHeader1 = sHeader1 & Left(vLista(iIndex, 3), 2) & "/" & Right(vLista(iIndex, 3), 2) 'Mês/Ano 5 caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 18) 'Emissão 10 Caracteres
'    sHeader1 = sHeader1 & "01/01" 'Pagina 5 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 14) 'Base de Cálculo 15 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 20) 'Aliquota 2 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 15) 'Valor ICMS 15 Caracteres
'    sHeader1 = sHeader1 & Left(vLista(iIndex, 4), 2) & "/" & Mid(vLista(iIndex, 4), 3, 2) & "/" & Right(vLista(iIndex, 4), 4) 'Vencimento 10 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 16) 'Total a Pagar 15 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 23) 'Código de Barras 48 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 24) 'DV 2 1 Caracter
'    sHeader1 = sHeader1 & vLista(iIndex, 25) 'Complemento 9 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 26) 'IPTE 27 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 5) 'Email Para 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 6) 'Email com Copia 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 7) 'Email Assunto 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 8) 'Email Corpo Mensagem 255 Caracteres
'    sHeader1 = sHeader1 & "FCTA" 'Transação Origem 4 Caracteres
'    sHeader1 = sHeader1 & String(2, " ") 'OPCAO CHAMADA NA TRANSACAO (Datel) 2 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 10) 'USUARIO ORIGINADOR DA SOLICITACAO 50 Caracteres
'    sHeader1 = sHeader1 & String(20, " ") 'CNPJ DO CLIENTE (Datel) 20 Caracteres
'    sHeader1 = sHeader1 & String(10, " ") 'NUMERO SISTEMICO (Datel) 10 Caracteres
'    sHeader1 = sHeader1 & String(255, " ") 'ENDERECO (Datel) 255 Caracteres
'    sHeader1 = sHeader1 & String(20, " ") 'TELEFONE DE CONTATO (Datel) 20 Caracteres
'    sHeader1 = sHeader1 & String(255, " ") 'CIDADE (Datel) 255 Caracteres
'    sHeader1 = sHeader1 & String(2, " ") 'UF (datel) 2 Caracteres
'    sHeader1 = sHeader1 & String(10, " ") 'Segmento (Datel) 10 Caracteres
'    sHeader1 = sHeader1 & String(20, " ") 'Incricao Estadual 20  Caracteres
'    sHeader1 = sHeader1 & String(10, " ") 'Codigo Credor
'    sHeader1 = sHeader1 & "000000000000,00" 'Base de Cálculo
'    sHeader1 = sHeader1 & "00" ' Aliquota
'    sHeader1 = sHeader1 & "000000000000,00" 'ICMS
'    sHeader1 = sHeader1 & String(255, " ") 'Descricao
'    sHeader1 = sHeader1 & String(4, " ") 'Codigo OS
'    sHeader1 = sHeader1 & String(3, " ") 'IDENTIFICACAO DO SISTEMA
'    sHeader1 = sHeader1 & String(5, " ") 'Orgao
'    vArq(iCont) = sHeader1
'    iCont = iCont + 1
'    '=============================================================================
'
'    'Definição do Header de Serviços
'    '=============================================================================
'    vServ = VBA.Split(vLista(iIndex, 12), "#") 'Separação em Array dos Serviços
'
'    'Armazenamento dos Serviços (Descriçao,Valor e Sinal)
'    For iServ = 0 To UBound(vServ)
'        ReDim Preserve vArq(0 To iCont)
'        vArq(iCont) = 2 & Replace(vServ(iServ), ";", "")
'        iCont = iCont + 1
'    Next iServ
'
'    'Definição Header 3
'    ReDim Preserve vArq(0 To iCont)
'    vArq(iCont) = 3 & Format(UBound(vServ) + 1, "000000000000.00")
'    iCont = iCont + 1
'
'    End If
'Next iIndex
'
'     'Definição Header 3
'    ReDim Preserve vArq(0 To iCont)
'    vArq(iCont) = 9 & Format(iCont + 1, "000000000000.00")
'
'    'Criação do Arquivo
'    If createTxtLog(sFile) Then
'        Call PrintLog(sFile, vArq)
'    End If
'
'End Sub
'Public Sub geraTxtFVIG()
'
'Dim iIndex      As Integer
'Dim iServ       As Integer
'Dim vServ       As Variant
'Dim sFolder     As String
'Dim sFile       As String
'Dim vArq        As Variant
'Dim iCont       As Integer
'Dim iContServ   As Integer
'Dim vLista      As Variant
'Dim sHeader1    As String
'
''Define Caminho do Diretorio a ser Armazenado o arquivo txt
''===============================================================
'sFolder = ThisWorkbook.Path & "\2Via_" & Replace(Date, "/", "")
''===============================================================
''Cria a pasta se não existir
''===============================================================
'If Not FileExists(sFolder, True) Then Call MkDir(sFolder)
''===============================================================
'
''Busca os Dados Gerados e armazena em Array
''==========================================
'vLista = armazenaTerminaisFVIG
''==========================================
'
''Percorre a lista e Gera um Arquivo para cada linha
''==================================================
'sFile = sFolder & "\FVIG_" & Replace(Date, "/", "") & "_" & Replace(Time, ":", "") & ".txt"
'ReDim vArq(0 To 0)
'iCont = 0
'
''Header 0 - Data e Hora da Geração do Arquivo
''====================================================================
'vArq(iCont) = "0" & Format(Date, "DD/MM/YYYY") & Format(Time, "HH:MM:SS")
'iCont = iCont + 1
''====================================================================
'
'For iIndex = 0 To UBound(vLista) - 1
'
'    '=======================================================================================================================
'    ReDim Preserve vArq(0 To iCont)
'
'    If vLista(iIndex, 10) = "GERAR TXT" Then
'    'Header 1 - Layout Conforme Versão 03
'    '====================================================================
'    ReDim Preserve vArq(0 To iCont)
'    sHeader1 = 1 'Header 1 Caracter
'    sHeader1 = sHeader1 & vLista(iIndex, 18) 'Numero NF 20 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 22) 'NRC 11 Caracteres
'    sHeader1 = sHeader1 & Replace(vLista(iIndex, 14), " ", "") & String(250 - Len(Replace(vLista(iIndex, 14), " ", "")), " ") 'Nome do Cliente 250 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 11)  'Local 5 Caracteres
'    sHeader1 = sHeader1 & Left(vLista(iIndex, 12), 4) & "-" & Right(vLista(iIndex, 12), 5) 'Terminal 10 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 23) 'DV 1 Caracter
'    sHeader1 = sHeader1 & vLista(iIndex, 20) & String(3 - Len(vLista(iIndex, 20)), " ") 'TA 3 Caracteres
'    sHeader1 = sHeader1 & Left(vLista(iIndex, 3), 2) & "/" & Right(vLista(iIndex, 3), 2) 'Mês/Ano 5 caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 19) 'Emissão 10 Caracteres
'    sHeader1 = sHeader1 & "01/01" 'Pagina 5 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 15) 'Base de Cálculo 15 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 21) 'Aliquota 2 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 16) 'Valor ICMS 15 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 29) 'Vencimento 10 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 17) 'Total a Pagar 15 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 24) 'Código de Barras 48 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 25) 'DV 2 1 Caracter
'    sHeader1 = sHeader1 & vLista(iIndex, 26) 'Complemento 9 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 27) 'IPTE 27 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 5) & String(255 - Len(vLista(iIndex, 5)), " ") 'Email Para 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 6) & String(255 - Len(vLista(iIndex, 6)), " ") 'Email com Copia 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 7) & String(255 - Len(vLista(iIndex, 7)), " ") 'Email Assunto 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 8) & String(255 - Len(vLista(iIndex, 8)), " ") 'Email Corpo Mensagem 255 Caracteres
'    sHeader1 = sHeader1 & "FVIG" 'Transação Origem 4 Caracteres
'    sHeader1 = sHeader1 & String(2, " ") 'OPCAO CHAMADA NA TRANSACAO (Datel) 2 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 9) & String(50 - Len(vLista(iIndex, 9)), " ") 'USUARIO ORIGINADOR DA SOLICITACAO 50 Caracteres
'    sHeader1 = sHeader1 & String(20, " ") 'CNPJ DO CLIENTE (Datel) 20 Caracteres
'    sHeader1 = sHeader1 & String(10, " ") 'NUMERO SISTEMICO (Datel) 10 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 28) & String(255 - Len(vLista(iIndex, 28)), " ") 'ENDERECO (Datel) 255 Caracteres
'    sHeader1 = sHeader1 & String(20, " ") 'TELEFONE DE CONTATO (Datel) 20 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 30) & String(255 - Len(vLista(iIndex, 30)), " ") 'CIDADE (Datel) 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 31) 'UF (datel) 2 Caracteres
'    sHeader1 = sHeader1 & String(10, " ") 'Segmento (Datel) 10 Caracteres
'    sHeader1 = sHeader1 & String(20, " ") 'Incricao Estadual 20  Caracteres
'    sHeader1 = sHeader1 & String(10, " ") 'Codigo Credor
'    sHeader1 = sHeader1 & "000000000000,00" 'Base de Cálculo
'    sHeader1 = sHeader1 & "00" ' Aliquota
'    sHeader1 = sHeader1 & "000000000000,00" 'ICMS
'    sHeader1 = sHeader1 & String(255, " ") 'Descricao
'    sHeader1 = sHeader1 & String(4, " ") 'Codigo OS
'    sHeader1 = sHeader1 & String(3, " ") 'IDENTIFICACAO DO SISTEMA
'    sHeader1 = sHeader1 & String(5, " ") 'Orgao
'    vArq(iCont) = sHeader1
'    iCont = iCont + 1
'    '=============================================================================
'
'    'Definição do Header de Serviços
'    '=============================================================================
'    vServ = VBA.Split(vLista(iIndex, 13), "#") 'Separação em Array dos Serviços
'
'    'Armazenamento dos Serviços (Descriçao,Valor e Sinal)
'    For iServ = 0 To UBound(vServ)
'        ReDim Preserve vArq(0 To iCont)
'        vArq(iCont) = 2 & Replace(vServ(iServ), ";", "")
'        iCont = iCont + 1
'    Next iServ
'
'    'Definição Header 3
'    ReDim Preserve vArq(0 To iCont)
'    vArq(iCont) = 3 & Format(UBound(vServ) + 1, "000000000000.00")
'    iCont = iCont + 1
'
'    End If
'Next iIndex
'
'     'Definição Header 3
'    ReDim Preserve vArq(0 To iCont)
'    vArq(iCont) = 9 & Format(iCont + 1, "000000000000.00")
'
'    'Criação do Arquivo
'    If createTxtLog(sFile) Then
'        Call PrintLog(sFile, vArq)
'    End If
'
'End Sub
'Public Sub gravaLog(sLog As String, sTerminal As String)
'    Dim sFile As String
'    Dim vTela As Variant
'
'    If bLog Then
'        sFile = ThisWorkbook.Path & "\Log_" & Replace(Date, "/", "_") & "_" & sTerminal & ".txt"
'        If createTxtLog(sFile) Then
'            vTela = CopiaTela(sLog)
'            Call PrintLog(sFile, vTela)
'        End If
'    End If
'End Sub
'Public Sub PrintLog(sFile As String, arrayTela As Variant)
'
'    Dim iIndex As Integer
'
'    Open sFile For Append As #1
'    For iIndex = 0 To UBound(arrayTela)
'        Print #1, arrayTela(iIndex)
'    Next iIndex
'
'    Close #1
'End Sub
'Sub StatusBar(Msg As String)
' Application.ScreenUpdating = True
'If Len(Msg) > 0 Then
'    Application.DisplayStatusBar = True
'    Application.StatusBar = Msg
'Else
'    Application.StatusBar = "Ready"
'    Application.DisplayStatusBar = True
'End If
'DoEvents
'Application.ScreenUpdating = False
'End Sub
'Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
'
'    'Purpose:   Return True if the file exists, even if it is hidden.
'    'Arguments: strFile: File name to look for. Current directory searched if no path included.
'    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
'    'Note:      Does not look inside subdirectories for the file.
'    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
'    Dim lngAttributes As Long
'
'    'Include read-only files, hidden files, system files.
'    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
'
'    If bFindFolders Then
'        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
'    Else
'        'Strip any trailing slash, so Dir does not look inside the folder.
'        Do While Right$(strFile, 1) = "\"
'            strFile = Left$(strFile, Len(strFile) - 1)
'        Loop
'    End If
'
'    'If Dir() returns something, the file exists.
'    On Error Resume Next
'    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
'End Function
'Public Function createTxtLog(strFile As String) As Boolean
''=========================================================
''Função Cria Arquivo Txt para Log das telas do CSO
''=========================================================
'
'    On Error GoTo ErrHandler
'
'    Dim fso     As Object
'    Dim iIndex  As Integer
'    Dim oFile   As Object
'
'    'Valida se já existe o arquivo na pasta Raiz
'    '===========================================
'    If FileExists(strFile) Then
'        createTxtLog = True
'        Exit Function
'    Else
'
'        'Se nao existir o arquivo, é criado
'        '==================================
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        Set oFile = fso.CreateTextFile(strFile)
'
'        If FileExists(strFile) Then
'            createTxtLog = True
'        Else
'            createTxtLog = False
'        End If
'
'        oFile.Close
'        Set fso = Nothing
'        Set oFile = Nothing
'    End If
'
'    Exit Function
'
'ErrHandler:
'    createTxtLog = False
'End Function
'Public Sub geraTxtDatel()
'
'Dim iIndex      As Integer
'Dim iServ       As Integer
'Dim vServ       As Variant
'Dim sFolder     As String
'Dim sFile       As String
'Dim vArq        As Variant
'Dim iCont       As Integer
'Dim vLista      As Variant
'Dim sHeader1    As String
'Dim sEmailPara  As String
'Dim sEmailCopia As String
'Dim sEmailAssunto As String
'Dim sEmailMensagem As String
'Dim sUsuarioSol     As String
'Dim sEndereco       As String
'Dim sCidade         As String
'Dim sDescricao      As String
'
''Define Caminho do Diretorio a ser Armazenado o arquivo txt
''===============================================================
'sFolder = ThisWorkbook.Path & "\2Via_" & Replace(Date, "/", "")
''===============================================================
''Cria a pasta se não existir
''===============================================================
'If Not FileExists(sFolder, True) Then Call MkDir(sFolder)
''===============================================================
'
''Busca os Dados Gerados e armazena em Array
''==========================================
'vLista = armazenaTerminaisFT01
''==========================================
'
''Percorre a lista e Gera um Arquivo para cada linha
''==================================================
'sFile = sFolder & "\DATEL_" & Replace(Date, "/", "") & "_" & Replace(Time, ":", "") & ".txt"
'ReDim vArq(0 To 0)
'iCont = 0
'
''Header 0 - Data e Hora da Geração do Arquivo
''====================================================================
'vArq(iCont) = "0" & Format(Date, "DD/MM/YYYY") & Format(Time, "HH:MM:SS")
'iCont = iCont + 1
''====================================================================
'
'For iIndex = 0 To UBound(vLista) - 1
'
'    '=======================================================================================================================
'    ReDim Preserve vArq(0 To iCont)
'
'    'Header 1 - Layout Conforme Versão 03
'    '====================================================================
'    sHeader1 = 1 'Header 1 Caracter
'    sHeader1 = sHeader1 & String(20, " ") 'Numero NF 20 Caracteres
'    sHeader1 = sHeader1 & String(11, " ") 'NRC 11 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 2) & String(250 - Len(vLista(iIndex, 2)), " ") 'Nome do Cliente 250 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 8)  'Local 5 Caracteres
'    sHeader1 = sHeader1 & String(10, " ") 'Terminal 10 Caracteres
'    sHeader1 = sHeader1 & " " 'DV 1 Caracter
'    sHeader1 = sHeader1 & String(3, " ") 'TA 3 Caracteres
'    sHeader1 = sHeader1 & String(5, " ") 'Mês/Ano 5 caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 23) 'Emissão 10 Caracteres
'    sHeader1 = sHeader1 & "01/01" 'Pagina 5 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 26) 'Base de Cálculo 15 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 27) 'Aliquota 2 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 28) 'Valor ICMS 15 Caracteres
'    sHeader1 = sHeader1 & Left(vLista(iIndex, 11), 2) & "/" & Mid(vLista(iIndex, 11), 3, 2) & "/" & Right(vLista(iIndex, 11), 4) 'Vencimento 10 Caracteres
'    sHeader1 = sHeader1 & Format(CDbl(vLista(iIndex, 13)), "000000000000.00") 'Total a Pagar 15 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 34) 'Código de Barras 48 Caracteres
'    sHeader1 = sHeader1 & "0" 'DV 2 1 Caracter
'    sHeader1 = sHeader1 & String(9, " ") 'Complemento 9 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 35) 'IPTE 27 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 14) & String(255 - Len(vLista(iIndex, 14)), " ") 'Email Para 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 15) & String(255 - Len(vLista(iIndex, 15)), " ") 'Email com Copia 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 16) & String(255 - Len(vLista(iIndex, 16)), " ") 'Email Assunto 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 17) & String(255 - Len(vLista(iIndex, 17)), " ") 'Email Corpo Mensagem 255 Caracteres
'    sHeader1 = sHeader1 & "FT01" 'Transação Origem 4 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 1) 'OPCAO CHAMADA NA TRANSACAO (Datel) 2 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 18) & String(50 - Len(vLista(iIndex, 18)), " ") 'USUARIO ORIGINADOR DA SOLICITACAO 50 Caracteres
'    sHeader1 = sHeader1 & Format(vLista(iIndex, 3), String(20, "0")) 'CNPJ DO CLIENTE (Datel) 20 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 21) & String(10 - Len(vLista(iIndex, 21)), " ") 'NUMERO SISTEMICO (Datel) 10 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 4) & String(255 - Len(vLista(iIndex, 4)), " ") 'ENDERECO (Datel) 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 10) & vLista(iIndex, 9) & String(20 - Len(vLista(iIndex, 10) & vLista(iIndex, 9)), " ") 'TELEFONE DE CONTATO (Datel) 20 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 5) & String(255 - Len(vLista(iIndex, 5)), " ") 'CIDADE (Datel) 255 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 6) 'UF (datel) 2 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 24) & String(10 - Len(vLista(iIndex, 24)), " ") 'Segmento (Datel) 10 Caracteres
'    sHeader1 = sHeader1 & String(20, " ") 'Incricao Estadual 20  Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 25) & String(10 - Len(vLista(iIndex, 25)), " ") 'Codigo Credor 10 Caracteres
'    sHeader1 = sHeader1 & vLista(iIndex, 29) 'Base de Cálculo
'    sHeader1 = sHeader1 & vLista(iIndex, 30) ' Aliquota
'    sHeader1 = sHeader1 & vLista(iIndex, 31) 'ICMS
'    sHeader1 = sHeader1 & vLista(iIndex, 12) & String(255 - Len(vLista(iIndex, 12)), " ") 'Descricao
'    sHeader1 = sHeader1 & vLista(iIndex, 33) 'Codigo PS
'    sHeader1 = sHeader1 & String(3, "9") 'IDENTIFICACAO DO SISTEMA
'    sHeader1 = sHeader1 & vLista(iIndex, 32) & "  " 'Orgao
'    vArq(iCont) = sHeader1
'    iCont = iCont + 1
'    '=============================================================================
'
'    'Definição do Header de Serviços
'    '=============================================================================
'    ReDim Preserve vArq(0 To iCont)
'    vArq(iCont) = 2 & "ARRECADACAO DE NFFST" & String(255 - Len("ARRECADACAO DE NFFST"), " ") & _
'                  Format(CDbl(vLista(iIndex, 13)), "000000000000.00") & " "
'    iCont = iCont + 1
'
'    'Definição Header 3
'    ReDim Preserve vArq(0 To iCont)
'    vArq(iCont) = 3 & Format(1, "000000000000.00")
'    iCont = iCont + 1
'
'Next iIndex
'
'     'Definição Header 3
'    ReDim Preserve vArq(0 To iCont)
'    vArq(iCont) = 9 & Format(iCont + 1, "000000000000.00")
'
'    'Criação do Arquivo
'    If createTxtLog(sFile) Then
'        Call PrintLog(sFile, vArq)
'    End If
'
'End Sub
