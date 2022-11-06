option Explicit

Public Sub SaveDocumentCopyAsPdf(ByRef doc As Document, ByVal saveCopyFullName As String, Optional ByVal startPage As Integer, Optional ByVal endPage As Integer)
'#######################################################################################
' SaveDocumentCopyAsPdf
'   Exports a copy of the Document to pdf and saves it to the file system. Can optionally
'   select a range of pages to export.
'
' Args
'   doc: The Document to copy
'   saveCopyFullName: The complete folder path and filename (with .pdf extension) of the copy
'   startPage: optional page number to begin the export at
'   endPage: optional page number to end the export at
'#######################################################################################
' check and adjust if start/end page args are missing
Dim firstPage As Integer
Dim lastPage As Integer
Dim pgCount As Integer
pgCount = doc.ComputeStatistics(wdStatisticPages)

firstPage = IIf(IsMissing(startPage), 1, startPage)

If IsMissing(endPage) Then
    lastPage = pgCount
Else
    lastPage = IIf(endPage > pgCount, pgCount, endPage)
End If

' do the export to pdf
doc.ExportAsFixedFormat _
    OutputFileName:=saveCopyFullName, _
    ExportFormat:=wdExportFormatPDF, _
    OpenAfterExport:=False, _
    OptimizeFor:=wdExportOptimizeForPrint, _
    Range:=wdExportFromTo, From:=firstPage, To:=lastPage, _
    Item:=wdExportDocumentContent, _
    IncludeDocProps:=False, _
    KeepIRM:=True, _
    CreateBookmarks:=wdExportCreateNoBookmarks, _
    DocStructureTags:=True, _
    BitmapMissingFonts:=True, _
    UseISO19005_1:=False

End Sub

Public Sub DeletePageFromDocument(ByRef doc As Document, ByVal pageNumber As Integer)
'#######################################################################################
' DeletePageFromDocument
'   Uses an entire page bookmark to remove a page, by number, from the document
'
' Args
'   doc: The Document to remove from
'   pageNumber: page number to remove
'#######################################################################################
Application.ScreenUpdating = False

doc.Activate
Selection.GoTo wdGoToPage, wdGoToAbsolute, pageNumber
doc.Bookmarks("\Page").Range.Delete

Application.ScreenUpdating = True

End Sub

Public Sub SaveDocumentCopyAsDocx(ByRef doc As Document, ByVal saveCopyFullName As String, Optional ByVal startPage As Integer, Optional ByVal endPage As Integer)
'#######################################################################################
' SaveDocumentCopyAsDocx
'   Exports a copy of the Document to .docx and saves it to the file system. Can optionally
'   select a range of pages to export.
'
' Args
'   doc: The Document to copy
'   saveCopyFullName: The complete folder path and filename (with .docx extension) of the copy
'   startPage: optional page number to begin the export at
'   endPage: optional page number to end the export at
'#######################################################################################
Dim newDoc As Document
Set newDoc = Application.Documents.Add(doc.FullName)

' Set the optional page ranges to not affect the doc when omitted
Dim pgCount As Integer
pgCount = ThisDocument.ComputeStatistics(wdStatisticPages)

If IsMissing(startPage) Then
    startPage = 1
End If

If IsMissing(endPage) Then
    endPage = pgCount
End If


' Remove pages, starting from the end of the document
Dim iterPage As Integer

If endPage < pgCount Then
    For iterPage = pgCount To (endPage + 1) Step -1
        DeletePageFromDocument newDoc, iterPage
    Next iterPage
End If

If startPage > 1 Then
    For iterPage = 1 To (startPage - 1) Step 1
        DeletePageFromDocument newDoc, iterPage
    Next iterPage
End If

' save and close the copy
newDoc.SaveAs2 _
    fileName:=saveCopyFullName, _
    FileFormat:=wdFormatXMLDocument, _
    AddToRecentFiles:=False
newDoc.Close savechanges:=False

End Sub
