Attribute VB_Name = "Module1"
Sub LoopThroughFiles()
'
' Loop through all *.txt files in current folder
'
' Keyboard Shortcut: Ctrl+q
'
    Dim previousSheet As Worksheet
    Set previousSheet = ActiveWorkbook.ActiveSheet
    
    Dim filename As String
    filename = Dir(Application.ActiveWorkbook.Path & "\*.txt")
    Do While Len(filename) > 0
        'MsgBox (filename)
        
        ' add a worksheet
        ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
        
        ' remove extension from filename
        Dim filename_no_suffix As String
        filename_no_suffix = Left(filename, InStr(filename, ".txt") - 1)
        ActiveWorkbook.ActiveSheet.Name = filename_no_suffix & Format(Now(), " hhmmss")
        
        ' import text file
        With ActiveSheet.QueryTables.Add(Connection:= _
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
        
        filename = Dir()
    Loop
    
    ' return to previous worksheet
    previousSheet.Activate

End Sub

