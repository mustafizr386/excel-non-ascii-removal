VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LengthCheckMenu 
   Caption         =   "LengthCheck"
   ClientHeight    =   2850
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "LengthCheckMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LengthCheckMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub executeButton_Click()
    Dim currentCell As String
    
    Dim startRow As Long
    Dim endRow As Long
    Dim maxLength As Integer
    Dim inputColumn As Long
    Dim r, g, b As Integer
    
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    inputColumn = Val(inputBox.Text)
    maxLength = Val(lengthBox.Text)
    r = Val(rBox.Text)
    g = Val(gBox.Text)
    b = Val(bBox.Text)
    
    For counter = startRow To endRow
        currentCell = Cells(counter, inputColumn)
        If Len(currentCell) > maxLength Then
            Cells(counter, inputColumn).Interior.Color = RGB(r, g, b)
        End If
    Next counter
    Unload Me
End Sub
