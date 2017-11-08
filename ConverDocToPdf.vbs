
Function isValidDoc(fileName)

    isValidDoc = False

    If (LCase(Right(fileName,4))=".doc" Or LCase(Right(fileName,4))="docx" ) And Left(fileName,1)<>"~" Then
        isValidDoc = True
     End If

End Function

Function convertToPdf(filePath)
    
    On Error Resume Next
    
    MsgBox filePath
    
    Const wdExportFormatPDF = 17
    Set oWord = WScript.CreateObject("Word.Application")
    Set oDoc = oWord.Documents.Open(filePath)
    Set pdfFilePath = Left(filePath, InStrRev(filePath, "."))&"pdf"
    
    MsgBox pdfFilePath
    odoc.ExportAsFixedFormat Left(filePath, InStrRev(filePath, "."))&"pdf",wdExportFormatPDF
    If Err.Number Then
        MsgBox Err.Description
        convertToPdf = Nothing
    End If
    
    odoc.Close
    oword.Quit
    Set oDoc = Nothing
    Set oWord = Nothing

End Function

Sub ConvertDocToPdf()

    On Error Resume Next

    Set fso = WScript.CreateObject("Scripting.Filesystemobject")

    Set oArgs = WScript.Arguments
        For Each s In oArgs
            Set ffile = fso.GetFile(s)
            If (fso.FileExists(s) And isValidDoc(ffile.name)) Then
                convertToPdf(ffile.path)
            End If
        Next
    Set oArgs = Nothing

    Set fds = fso.GetFolder(".")
    Set ffs = fds.Files
    For Each ff In ffs
        If (isValidDoc(ff.name)) Then
            convertToPdf(ff.path)
        End If
    Next

    MsgBox "Word文件已全部轩换为PDF格式!"

End Sub

call ConvertDocToPdf()
