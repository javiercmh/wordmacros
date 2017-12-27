Sub Duplicate()
'
' Duplicate line Macro
'
'
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveUp Unit:=wdParagraph, Count:=1, Extend:=wdExtend
    Selection.Copy
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
End Sub
