Attribute VB_Name = "DataCompiler"
Sub DataCompile_A_StripRepeats()
    StripRepeatsMenu.Show
End Sub

Sub DataCompile_B_CompileData()
    CompileDataMenu.Show
End Sub

Sub DataCompile_C_ClearExtraNewLines()
    ClearExtraNewLinesMenu.Show
End Sub

Sub DataAnalysis_LengthCheck()
    LengthCheckMenu.Show
End Sub

Sub DataAnalysis_NonASCIIRemove()
    NonASCIIRemoveMenu.Show
End Sub

Sub DataAnalysis_NonASCIICheck()
    NonASCIICheckMenu.Show
End Sub

Sub DataAnalysis_NonASCIIExport()
    NonASCIIExportMenu.Show
End Sub

Sub DataAnalysis_LengthExport()
    LengthExportMenu.Show
End Sub

Sub DataAnalysis_RemoveLeadingNumbers()
    RemoveLeadingNumbersMenu.Show
End Sub

Sub DataAnalysis_80CharAbbreviate()
    FindNReplaceMenu.Show
End Sub

Sub DataAnalysis_StripRepeatingPlus()
    StripRepeatsPlusMenu.Show
End Sub


Sub Hierarchy_Enumeration()
    HierEnumerateMenu.Show
End Sub

Sub CatchSpellingErrors()
    SpellCheckMenu.Show
End Sub

Sub FillParent()
    FillParentMenu.Show
End Sub
