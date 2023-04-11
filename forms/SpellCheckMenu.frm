VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SpellCheckMenu 
   Caption         =   "SpellCheck"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "SpellCheckMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SpellCheckMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub executeButton_Click()
    Dim rng As Range, cell As Range, tarCell As Range
    Dim start As Integer, ending As Integer
    
    Dim words() As String
    
    Set rng = ActiveSheet.Range(regionBox.Text)
    
    Dim timer As Integer
    timer = 0
    
    For Each cell In rng.Cells
            If (timer Mod 200) = 0 Then
                Application.Wait (Now + TimeValue("0:00:01"))
            End If
            words() = Split(cell.Text)
            For Each word In words
                If Not Application.CheckSpelling(word:=word) Then
                    start = InStr(1, cell.Text, word)
                    ending = Len(word)
                    Cells(cell.Row, cell.Column).Characters(start, ending).Font.Size = Cells(cell.Row, cell.Column).Characters(start, ending).Font.Size + 4
                    Cells(cell.Row, cell.Column).Characters(start, ending).Font.Color = RGB(255, 255, 0)
                    Cells(cell.Row, cell.Column).Interior.ColorIndex = 3
                End If
            Next word
            timer = timer + 1
    Next cell
    
    Unload Me
End Sub
