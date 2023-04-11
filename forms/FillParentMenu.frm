VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FillParentMenu 
   Caption         =   "FillParent"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "FillParentMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FillParentMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub executeButton_Click()
    Dim rng As Range, cell As Range
    Dim previous As String
    
    Set rng = ActiveSheet.Range(regionBox.Text)
    previous = "ERROR FIRST ENTRY BLANK"
    
    Dim counter As Integer, colCounter As Integer
    counter = 0
    colCounter = 0
    For Each clmn In rng.Columns
        Debug.Print (colCounter)
        For Each cell In clmn.Cells
            If colCounter = 0 Then
                If cell.Text = "" Then
                    cell = previous
                Else
                    cell.Font.Bold = True
                    counter = counter + 1
                End If
                If (counter Mod 2) = 0 Then
                    cell.Interior.ColorIndex = 34
                Else
                    cell.Interior.ColorIndex = 15
                End If
            Else
                cell.Interior.ColorIndex = Cells(cell.Row, cell.Column - colCounter).Interior.ColorIndex
            End If
            previous = cell.Text
        Next cell
        colCounter = colCounter + 1
    Next clmn
    
    Unload Me
End Sub
