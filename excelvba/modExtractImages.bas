Attribute VB_Name = "modExtractImages"
'#######################################################################################
' modExtractImages
'
' Extract images from Excel Charts or NamedRanges
'
'#######################################################################################


Public Function CollectShapesByPrefix(ByRef wb As Workbook, ByVal prefix As String) As Collection
'#######################################################################################
' CollectShapesByPrefix
'   Scans all Shapes in the workbooks and gathers any shape whose name begins with
'   prefix into the collection
'
' Args
'   wb: Source Workbook
'   prefix: Naming convention of the Shape objects to collect, for example "Fig_PieChart"
'       and "Fig_LineChart" would both be found with the prefix "Fig_"
'
' Returns
'   Collection of Excel Shape objects. The collection can be empty.
'#######################################################################################
Dim shapes As Collection
Set shapes = New Collection

Dim prefixLen As Integer
prefixLen = Len(prefix)
Dim sht As Worksheet
Dim shp As Shape
For Each sht In wb.Sheets
    For Each shp In sht.shapes
        If Left(shp.Name, prefixLen) = prefix Then
            shapes.Add shp
        End If
    Next shp
Next sht

Set CollectShapesByPrefix = shapes

End Function

Public Sub SaveShapesAsEmfWithPowerPoint(ByRef shapes As Collection, ByVal saveFolder As String, Optional ByVal userPromptMessage As String = vbNullString)
'#######################################################################################
' SaveShapesAsEmfWithPowerPoint
'   Copy/Pastes given Excel Shape objects into a new PowerPoint instance only to export
'   them as .emf images to the file system
'
' Args
'   shapes: Array of Excel Shape objects
'   saveFolder: Existing folder to place the exported images. Macro fails if this folder
'       does not exist
'#######################################################################################

' Guard against the save folder not existing
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(saveFolder) Then
    MsgBox "The folder designated for saving images does not exist. Please ensure it is created before running this macro again." & vbNewLine & vbNewLine & saveFolder
    Exit Sub
End If

' prompt the user to confirm running the macro
Dim msgPrompt As String
Dim resultPrompt As VbMsgBoxResult

If userPromptMessage = vbNullString Then
    msgPrompt = "Settings for extracting EnhancedMetaFile (.emf) images:" & _
        vbNewLine & vbNewLine & _
        shapes.Count & " images will be saved in '" & saveFolder & "'." & _
        vbNewLine & vbNewLine & _
        "The export is done using PowerPoint, which will open and close while this macro is running. " & _
        "It is recommended to close any currently open ppt files before proceeding."
Else
    msgPrompt = userPromptMessage
End If

resultPrompt = MsgBox(msgPrompt, vbOKCancel)

If resultPrompt <> vbOK Then
    Exit Sub
End If


' perform the export
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim ppt As Object
Dim pr As Object
Dim sl As Object

Set ppt = CreateObject("PowerPoint.Application")
Set pr = ppt.Presentations.Add

Dim shp As Shape
Dim slideCount As Integer
For Each shp In shapes
    shp.Copy
    slideCount = pr.Slides.Count
    Set sl = pr.Slides.Add(slideCount + 1, 11)
    sl.shapes.PasteSpecial DataType:=2, Link:=msoFalse 'DataType 2 is PpPasteDataType.ppPasteEnhancedMetafile
    sl.shapes(sl.shapes.Count).Export fso.BuildPath(saveFolder, shp.Name & ".emf"), 5 '5 is PpShapeFormat.ppShapeFormatEMF
Next shp

' exit cleanup
Application.CutCopyMode = False

pr.Close
ppt.Quit

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub