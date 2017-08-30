Attribute VB_Name = "NewMacros"
Sub Paste_Table()
Attribute Paste_Table.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    
    On Error GoTo er
    
    Selection.Paste
    With Selection.Tables(1)
        .AutoFitBehavior (wdAutoFitWindow)
        .Range.ParagraphFormat.SpaceAfter = 0
        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.Font.Size = 9
        .Range.Font.Name = "Calibri"
        .Rows.Height = 13
        .Rows.HeightRule = wdRowHeightExactly
    End With
    
    Exit Sub
    
er:
    MsgBox "Selection is not table format."
    
End Sub
