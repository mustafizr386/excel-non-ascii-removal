VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompileDataMenu 
   Caption         =   "CompileData"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "CompileDataMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompileDataMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub executeButton_Click()
    Dim currentCell As String
    Dim inputCell As String
    Dim dumpCell As Integer
    
    Dim startRow As Long
    Dim endRow As Long
    Dim strippedInputColumn As Long '3
    Dim outputColumn As Long '4
    Dim dataInputColumn As Long '2
    
    
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    strippedInputColumn = Val(strippedBox.Text)
    dataInputColumn = Val(dataBox.Text)
    outputColumn = Val(outputBox.Text)
    
    
    For counter = startRow To endRow
        currentCell = Cells(counter, outputColumn).Value
        inputCell = Cells(counter, strippedInputColumn).Value
        If inputCell <> "" Then
            dumpCell = counter
            Cells(dumpCell, outputColumn).Value = Chr(34) & Cells(counter, dataInputColumn) & Chr(34) & Chr(10)
        Else
            Cells(dumpCell, outputColumn).Value = Cells(dumpCell, outputColumn) & Chr(34) & Cells(counter, dataInputColumn) & Chr(34) & Chr(10)
        End If
    Next counter
    
    For counter = startRow To endRow
        currentCell = Cells(counter, outputColumn).Value
        If currentCell <> "" And Right(currentCell, 1) = Chr(10) Then
            Cells(counter, outputColumn).Value = Left(currentCell, Len(currentCell) - 1)
        End If
    Next counter
    
    Unload Me
End Sub

