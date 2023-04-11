VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StripRepeatsMenu 
   Caption         =   "StripRepeats"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "StripRepeatsMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StripRepeatsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub executeButton_Click()
    Dim currentCell As String
    Dim previousCell As String
    
    Dim startRow As Long
    Dim endRow As Long
    Dim inputColumn As Long
    Dim outputColumn As Long
    
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    inputColumn = Val(inputBox.Text)
    outputColumn = Val(outputBox.Text)
    
    For counter = startRow To endRow
        currentCell = Cells(counter, inputColumn).Value
        If currentCell <> previousCell Then
            Cells(counter, outputColumn).Value = currentCell
            Cells(counter, outputColumn).Font.Bold = True
        End If
        previousCell = currentCell
    Next counter
    Unload Me
End Sub

