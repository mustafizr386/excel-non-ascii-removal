VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HierEnumerateMenu 
   Caption         =   "HierEnumerate"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3285
   OleObjectBlob   =   "HierEnumerateMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HierEnumerateMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub executeButton_Click()
    Dim rng As Range, cell As Range, prevRng As Range
    Dim outputCol As Integer, counter As Integer, prevCol As Integer
    Dim isPrev As Boolean
    Dim genInput As String, lastGen As String
    
    Set rng = ActiveSheet.Range(regionBox.Text)
    outputCol = outputBox.Value
    genInput = prevRegionBox.Value
    
    If genInput <> "na" Then
        isPrev = True
        prevCol = prevRegionBox.Value
    Else
        isPrev = False
        prevCol = 1
    End If
    
    
    counter = 1
    lastGen = ""
    
    
    For Each cell In rng.Cells
        If cell <> "" Then
            If isPrev And Cells(cell.Row, prevCol) <> "" Then
                lastGen = Cells(cell.Row, prevCol).Text
                counter = 1
                Debug.Print ("lastGen " + lastGen)
            End If
            
            If isPrev Then
                Debug.Print (str(counter) + " " + lastGen)
                Cells(cell.Row, outputCol) = lastGen + "." + Right(str(counter), Len(str(counter)) - 1)
                counter = counter + 1
            Else
                Debug.Print (cell)
                Cells(cell.Row, outputCol) = counter
                counter = counter + 1
            End If
        End If
    Next cell
    
    Unload Me
End Sub

