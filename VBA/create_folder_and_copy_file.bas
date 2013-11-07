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
    
    ' http://msdn.microsoft.com/en-us/library/office/aa221353(v=office.11).aspx
    ' Loop through a range of cells
    Do While Len(fromPath) > 0 And Len(toPath) > 0
        ' add current path (the Excel sheet's location) to the front of the paths
        fromPath = Application.ActiveWorkbook.Path & "\" & fromPath
        toPath = Application.ActiveWorkbook.Path & "\" & toPath
        MsgBox fromPath & " >> " & toPath
        
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
            fso.CopyFile fromPath, toPath
        End If
         
        ' go to the next row
        i = i + 1
        ' get the source path
        fromPath = ActiveWorkbook.ActiveSheet.Cells(i, 1).Value
        ' get the target path
        toPath = ActiveWorkbook.ActiveSheet.Cells(i, 2).Value
    Loop
    
End Sub
