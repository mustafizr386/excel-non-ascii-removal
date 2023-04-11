VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClearExtraNewLinesMenu 
   Caption         =   "ClearExtraNewLines"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "ClearExtraNewLinesMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClearExtraNewLinesMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function removeNewLine(str As String)
    Dim i As Integer
    Dim result As String
    Dim countr As Integer
    result = str
    countr = 0
    Debug.Print (str)
    Debug.Print ("oh done")
    removeNewLine = Right(result, Len(str) - countr)
End Function



Private Sub executeButton_Click()
    Dim currentCell As String
    
    Dim startRow As Long
    Dim endRow As Long
    Dim inputColumn As Long
    
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    inputColumn = Val(inputBox.Text)
    
    For counter = startRow To endRow
        currentCell = Cells(counter, inputColumn).Value
        If currentCell <> "" And Right(currentCell, 1) = Chr(10) Then
            Debug.Print ("yes")
            Cells(counter, inputColumn).Value = Left(currentCell, Len(currentCell) - 1)
        End If
    Next counter
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
