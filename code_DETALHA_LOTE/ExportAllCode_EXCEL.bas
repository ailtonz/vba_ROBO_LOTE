Attribute VB_Name = "ExportAllCode_EXCEL"

Private Sub AddRefGuid()
On Error Resume Next

    'Add VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3
   
    Application.VBE.VBProjects(1).References.AddFromGuid _
        "{0002E157-0000-0000-C000-000000000046}", 2, 0
 
End Sub

Public Sub ExportAllCode_EXCEL() '' Extracao de codigos do projeto
'' https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files

    Dim c As VBComponent
    Dim Sfx As String
    Dim sFileName As String: sFileName = "\" & Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))

    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then
            
            '''' EXCEL
            CreateDir Application.ActiveWorkbook.Path & sFileName & "\code\"
            c.Export FileName:=Application.ActiveWorkbook.Path & sFileName & "\code\" & c.Name & Sfx
                        
        End If
    Next c
        MsgBox "Created source files in " & Application.ActiveWorkbook.Path & sFileName
End Sub


Private Function CreateDir(strPath As String) '' Criar estrutura de diretorios
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Function
