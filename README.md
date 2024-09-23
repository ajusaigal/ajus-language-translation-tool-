# ajus-language-translation-tool- VB SCRIPT
malayalam to english
Sub Multi Find NReplace()

Dim Rng As Range
Dim InputRng As Range, ReplaceRng As Range
xTitleId = "AjusBigData"
Set InputRng = Application. Selection
Set InputRng = Application. InputBox ("Original Range ", xTitleId, InputRng. Address, Type:=8)
Set ReplaceRng = Application. InputBox("Replace Range :", xTitleId, Type:=8)
Application. Screen Updating = False
For Each Rng In Replace Rng.Columns (1).Cells
InputRng.Replace what:=Rng.Value, replacement:=Rng.Offset(0, 1).Value
,Look At:=xlPart, _
Search Order: =xlByRows, Match Case:=False, SearchFormat:=False, Replace Format:=False
Next
Application. Screen Updating = True
End Sub
