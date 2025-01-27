Sub SaveAsPDF()
    Dim docName As String
    Dim pdfPath As String
    
    ' Get the name of the active document (without the extension)
    docName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    
    ' Set the PDF file path (same folder as the Word document)
    pdfPath = ActiveDocument.Path & "\" & docName & ".pdf"
    
    ' Save the document as PDF
    ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfPath, _
        ExportFormat:=wdExportFormatPDF
    
    MsgBox "PDF saved as: " & pdfPath
End Sub