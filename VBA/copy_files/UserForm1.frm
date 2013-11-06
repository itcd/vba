VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2496
   ClientLeft      =   36
   ClientTop       =   348
   ClientWidth     =   3768
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim fromPath As String
    fromPath = Application.ActiveWorkbook.Path & "\" & TextBox1.Text
    Dim toPath As String
    toPath = Application.ActiveWorkbook.Path & "\" & TextBox2.Text
    
    ' Check if the source files exist
    If Len(Dir(fromPath)) = 0 Then
        MsgBox fromPath & " doesn't exist."
    Else
        MsgBox fromPath & " exists."
    End If
    
    ' http://www.techonthenet.com/excel/formulas/dir.php
    ' Check if the target foler exists
    If Len(Dir(toPath, vbDirectory)) = 0 Then
        MsgBox toPath & " doesn't exist. It will be created."
        
        ' http://msdn.microsoft.com/en-us/library/k1d22wfh(v=vs.90).aspx
        ' Create folder
        MkDir (toPath)
    Else
        MsgBox toPath & " exists."
    End If
    
    ' http://msdn.microsoft.com/en-us/library/2wcf3ba6(v=vs.84).aspx
    ' Use fso.MoveFile to move files or use fso.CopyFile to copy files.
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile fromPath, toPath
    
End Sub

