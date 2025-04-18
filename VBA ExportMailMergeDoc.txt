Public Sub ExportMailMergeDoc()
    ExportMailMerge "doc"
End Sub

Public Sub ExportMailMergePDF()
    ExportMailMerge "pdf"
End Sub

Sub ExportMailMerge(Optional Format As String = "doc")
    On Error GoTo HandleError

    Dim doc As Document
    Set doc = ActiveDocument

    ' Validate format
    If LCase(Format) <> "doc" And LCase(Format) <> "pdf" Then
        MsgBox "Invalid export format: " & Format, vbCritical
        Exit Sub
    End If

    ' Check if "EmployeeName" field exists
    Dim fieldExists As Boolean
    Dim mmField As MailMergeField
    fieldExists = False
    For Each mmField In doc.MailMerge.Fields
        If InStr(mmField.Code.Text, "EmployeeName") > 0 Then
            fieldExists = True
            Exit For
        End If
    Next
    If Not fieldExists Then
        MsgBox "EmployeeName merge field not found", vbCritical
        Exit Sub
    End If

    ' Set up export directory
    Dim exportPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    exportPath = fso.GetParentFolderName(GetLocalPath(doc.FullName)) & "\MergeExport"
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
    End If

    ' Prepare mail merge
    With doc.MailMerge
        If .State <> wdMainAndDataSource Then
            MsgBox "Mail merge is not properly set up.", vbCritical
            Exit Sub
        End If
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        .DataSource.ActiveRecord = wdFirstRecord

        Dim i As Long
        Dim total As Long
        total = .DataSource.RecordCount

        For i = 1 To total
            .DataSource.ActiveRecord = i
            Dim rawName As String
            rawName = Trim(.DataSource.DataFields("EmployeeName").Value)

            If rawName = "" Then
                MsgBox "EmployeeName value is blank, this should not happen. Check for blank value in data source", vbCritical
                Exit Sub
            End If

            ' Sanitize filename
            Dim fileName As String
            fileName = rawName
            fileName = Replace(fileName, " ", "")
            fileName = RemoveNonAlpha(fileName)

            If LCase(Format) = "doc" Then
                fileName = fileName & ".docx"
            Else
                fileName = fileName & ".pdf"
            End If

            .DataSource.FirstRecord = i
            .DataSource.lastRecord = i
            .Destination = wdSendToNewDocument
            .Execute Pause:=False

            Dim resultDoc As Document
            Set resultDoc = ActiveDocument

            If LCase(Format) = "doc" Then
                resultDoc.SaveAs2 fileName:=exportPath & "\" & fileName, FileFormat:=wdFormatXMLDocument
            Else
                resultDoc.ExportAsFixedFormat OutputFileName:=exportPath & "\" & fileName, _
                    ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint
            End If

            resultDoc.Close SaveChanges:=False
        Next i
    End With

    MsgBox "Mail merge export successful. " & total & " file(s) created in: " & exportPath, vbInformation
    Exit Sub

HandleError:
    MsgBox "Error: " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Function RemoveNonAlpha(str As String) As String
    Dim i As Long, result As String, ch As String
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If ch Like "[A-Za-z]" Then
            result = result & ch
        End If
    Next i
    RemoveNonAlpha = result
End Function


