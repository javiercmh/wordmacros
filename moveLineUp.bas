Sub MoveLineUp()
'
' Move line up Macro
'
'
    Selection.EndKey Unit:=wdLine
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Cut
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
End Sub
