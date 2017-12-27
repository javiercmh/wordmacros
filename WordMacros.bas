Attribute VB_Name = "NewMacros"
Sub Duplicate()
Attribute Duplicate.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Duplicate"
'
' Duplicate Macro
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
Sub MoveLineDown()
'
' Move line down Macro
'
'
    Selection.EndKey Unit:=wdLine
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Cut
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
End Sub
