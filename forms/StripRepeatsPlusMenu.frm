VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StripRepeatsPlusMenu 
   Caption         =   "StripRepeats+"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "StripRepeatsPlusMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StripRepeatsPlusMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub executeButton_Click()
    Dim rng As Range, cell As Range
    Dim previous As String
    
    
    Set rng = ActiveSheet.Range(regionBox.Text)
    previous = "rlock!23"
    
    For Each clmn In rng.Columns
        For Each cell In clmn.Cells
            Dim clear As Boolean
            clear = False
            If previous = cell Then
                clear = True
            End If
            previous = cell
            If clear Then
                cell.Value = ""
            End If
        Next cell
    Next clmn
    
    Unload Me
End Sub
