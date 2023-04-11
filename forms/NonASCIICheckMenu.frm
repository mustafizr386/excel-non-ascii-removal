VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NonASCIICheckMenu 
   Caption         =   "NonASCIIRemove"
   ClientHeight    =   2715
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "NonASCIICheckMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NonASCIICheckMenu"
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
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    
    inputColumn = Val(inputBox.Text)
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    r = Val(rBox.Text)
    g = Val(gBox.Text)
    b = Val(bBox.Text)
    
    For counter = startRow To endRow
        currentCell = Cells(counter, inputColumn).Value
        newCell = ""
        For inside = 1 To Len(currentCell)
            currentChar = Mid(currentCell, inside, 1)
            currentAscii = Asc(currentChar)
            If Not (currentAscii > 126) Then 'Or currentAscii < 33
                newCell = newCell & currentChar
            Else
                Cells(counter, inputColumn).Characters(inside, 1).Font.Color = RGB(255, 255, 255)
                Cells(counter, inputColumn).Characters(inside, 1).Font.Size = Cells(counter, inputColumn).Characters(inside, 1).Font.Size + 4
            End If
        Next inside
        If currentCell <> newCell Then
            Cells(counter, inputColumn).Interior.Color = RGB(r, g, b)
        End If
    Next counter
    Unload Me
End Sub


