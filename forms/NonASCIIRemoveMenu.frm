VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NonASCIIRemoveMenu 
   Caption         =   "NonASCIIRemove"
   ClientHeight    =   1905
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "NonASCIIRemoveMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NonASCIIRemoveMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub executeButton_Click()
    Dim currentCell As String
    Dim inputColumn As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim currentChar As String
    Dim currentAscii As Integer
    Dim newCell As String
    
    inputColumn = Val(inputBox.Text)
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    For counter = startRow To endRow
        currentCell = Cells(counter, inputColumn).Value
        newCell = ""
        For inside = 1 To Len(currentCell)
            currentChar = Mid(currentCell, inside, 1)
            currentAscii = Asc(currentChar)
            If Not (currentAscii > 126) Then 'Or currentAscii < 33
                newCell = newCell & currentChar
            End If
        Next inside
        If currentCell <> newCell Then
            Cells(counter, inputColumn) = newCell
        End If
    Next counter
    Unload Me
End Sub
