Attribute VB_Name = "mdlLog"
'Public Const bLog = True
'
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
'    For iIndex = 1 To UBound(arrayTela)
'        Write #1, arrayTela(iIndex)
'    Next iIndex
'
'    Close #1
'End Sub
'Sub StatusBar(Msg As String)
'
'If Len(Msg) > 0 Then
'    Application.DisplayStatusBar = True
'    Application.StatusBar = Msg
'Else
'    Application.StatusBar = "Ready"
'    Application.DisplayStatusBar = True
'End If
'DoEvents
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
'    On Error GoTo errHandler
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
'errHandler:
'    createTxtLog = False
'End Function
'Sub IniciaPrintTela()
'
'    Load frmTelaPw
'    frmTelaPw.Show vbModal
'
'    Call PrintTela
'    DoEvents
'    'frmTelaPw.Hide
'
'End Sub
'Sub TerminaPrint()
'    Unload frmTelaPw
'End Sub
