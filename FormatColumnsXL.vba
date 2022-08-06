Function FormatColumns(ArgSourceCol, ArgTargetCol)
    Columns(ArgSourceCol).Select
    Selection.Copy
    Columns(ArgTargetCol).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Function