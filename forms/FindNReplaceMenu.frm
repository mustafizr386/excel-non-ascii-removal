VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindNReplaceMenu 
   Caption         =   "FindNReplace"
   ClientHeight    =   2610
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "FindNReplaceMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindNReplaceMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub executeButton_Click()
    Dim fndList As Variant
    Dim rplcList As Variant
    Dim curCell As Variant
    
    Dim inputColumn As Long
    Dim outputColumn As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim targetLength As Long
    
    fndList = Array("Communication", "Prevention", "immunizations", "United States Of America", "United States", "Veterans")
    rplcList = Array("Comm", "Pevnt", "Imune", "USA", "US", "Vets")
    
    inputColumn = Val(inputBox.Text)
    outputColumn = Val(outputBox.Text)
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    targetLength = Val(lengthBox.Text)
    
    For x = startRow To endRow
        curCell = Cells(x, inputColumn).Value
        If Len(curCell) > targetLength Then
            Cells(x, outputColumn).Value = curCell
        End If
    Next x
    
    For word = LBound(fndList) To UBound(fndList)
        For x = startRow To endRow
            curCell = Cells(x, outputColumn).Value
            If InStr(curCell, fndList(word)) <> 0 And Len(curCell) > targetLength Then
                Cells(x, outputColumn) = Replace(curCell, fndList(word), rplcList(word))
            End If
        Next x
    Next word
    Unload Me
End Sub
