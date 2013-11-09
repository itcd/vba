Attribute VB_Name = "Module1"
Sub CopyFile()
Attribute CopyFile.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' CopyFiles Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Dim i As Integer
    i = 1
    Dim fromPath As String
    ' get the source path in cell (i, 1). i-th row, first column
    fromPath = ActiveWorkbook.ActiveSheet.Cells(i, 1).Value
    ' get the target path in cell (i, 2) i-th row, second column
    Dim toPath As String
    toPath = ActiveWorkbook.ActiveSheet.Cells(i, 2).Value
    
    ' get workbook name and reference number
    Dim workbookName As String
    workbookName = ActiveWorkbook.ActiveSheet.Cells(i, 3).Value
    Dim tabName As String
    tabName = ActiveWorkbook.ActiveSheet.Cells(i, 4).Value
    
    ' http://msdn.microsoft.com/en-us/library/office/aa221353(v=office.11).aspx
    ' Loop through a range of cells
    Do While Len(fromPath) > 0 And Len(toPath) > 0
        ' add current path (the Excel sheet's location) to the front of the paths
        fromPath = Application.ActiveWorkbook.Path & "\" & fromPath
        toPath = Application.ActiveWorkbook.Path & "\" & toPath
        
        ' get the path and name of the empty workbook
        Dim workbookPathAndName As String
        workbookPathAndName = Application.ActiveWorkbook.Path & "\" & workbookName
        MsgBox fromPath & " >> " & toPath & "    workbook=" & workbookPathAndName
        
        ' Copy or move files
        ' Check if the source files exist
        If Len(Dir(fromPath)) = 0 Then
            MsgBox fromPath & " doesn't exist. Nothing to do."
        Else
            'MsgBox fromPath & " exists."
            
            ' http://www.techonthenet.com/excel/formulas/dir.php
            ' Check if the target foler exists
            If Len(Dir(toPath, vbDirectory)) = 0 Then
                'MsgBox toPath & " doesn't exist. It will be created."
                
                ' http://msdn.microsoft.com/en-us/library/k1d22wfh(v=vs.90).aspx
                ' Create folder
                MkDir (toPath)
            'Else
                'MsgBox toPath & " exists."
            End If
            
            ' http://msdn.microsoft.com/en-us/library/2wcf3ba6(v=vs.84).aspx
            ' Use fso.MoveFile to move files or use fso.CopyFile to copy files.
            Dim fso
            Set fso = CreateObject("Scripting.FileSystemObject")
            
            ' Copy the empty workbook
            fso.CopyFile workbookPathAndName, toPath & "\" & workbookName
  
            ' Copy data files
            fso.CopyFile fromPath, toPath
            'fso.MoveFile fromPath, toPath
            
            ' http://msdn.microsoft.com/en-us/library/office/ff194819.aspx
            ' Import text files into the workbook
            Dim ControlFile As String
            ControlFile = ActiveWorkbook.Name
            Workbooks.Open filename:=toPath & "\" & workbookName
            ActiveSheet.Name = tabName
            
            '================================================================
            ' Import text files
            Dim filename As String
            filename = Dir(toPath & "\*.txt")
            Do While Len(filename) > 0
                ' add a worksheet
                ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
                
                'MsgBox (LCase(filename))
                
                ' remove extension from filename
                Dim filename_no_suffix As String
                filename_no_suffix = Left(filename, InStr(LCase(filename), ".txt") - 1)
                ActiveWorkbook.ActiveSheet.Name = tabName & " " & filename_no_suffix ' & Format(Now(), " hhmmss")
                
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
            '================================================================
            
            ' Save and close the workbook
            ActiveWorkbook.Close SaveChanges:=True
            
            ' Return to this control workbook
            Windows(ControlFile).Activate
            
        End If
        
         
        ' go to the next row
        i = i + 1
        
        ' get the source path and target path
        fromPath = ActiveWorkbook.ActiveSheet.Cells(i, 1).Value
        toPath = ActiveWorkbook.ActiveSheet.Cells(i, 2).Value
        
        ' get workbook name and reference number
        workbookName = ActiveWorkbook.ActiveSheet.Cells(i, 3).Value
        tabName = ActiveWorkbook.ActiveSheet.Cells(i, 4).Value
    Loop

End Sub
