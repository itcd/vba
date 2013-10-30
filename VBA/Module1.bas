Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Dim listSheet As Worksheet
    Set listSheet = ActiveWorkbook.activeSheet
    
    For Each c In ActiveCell.CurrentRegion.Cells
        Dim filename As String
        filename = c.Value
        'MsgBox ("TEXT;" & Application.ActiveWorkbook.Path & "\" & filename)
        
        If filename <> "" Then
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
    
            With activeSheet.QueryTables.Add(Connection:= _
                "TEXT;" & Application.ActiveWorkbook.Path & "\" & filename, Destination:=Range("$A$1"))
                .Name = filename
                .FieldNames = True
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .TextFilePromptOnRefresh = False
                .TextFilePlatform = 936
                .TextFileStartRow = 1
                .TextFileParseType = xlDelimited
                .TextFileTextQualifier = xlTextQualifierDoubleQuote
                .TextFileConsecutiveDelimiter = True
                .TextFileTabDelimiter = True
                .TextFileSemicolonDelimiter = False
                .TextFileCommaDelimiter = False
                .TextFileSpaceDelimiter = True
                .TextFileColumnDataTypes = Array(1)
                .TextFileTrailingMinusNumbers = True
                .Refresh BackgroundQuery:=False
            End With
        
            ActiveWorkbook.activeSheet.Name = filename & Format(Now(), " hhmmss")
        End If
         
    Next
    
    listSheet.Activate

End Sub
