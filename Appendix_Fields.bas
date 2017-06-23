Attribute VB_Name = "Appendix_Fields"
Sub AppendixFields_regular()

'counter = 1

'Do While counter <= 30

    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ ABC \c", PreserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    PreserveFormatting:=False
    Selection.TypeText Text:="=MOD("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1,26)+1 \* ALPHABETIC"
    Selection.Collapse direction:=wdCollapseEnd
    Selection.MoveEnd (wdLine)
    Selection.Collapse (wdCollapseEnd)
    
    
'    counter = counter + 1
    
    
    
'Loop

'ActiveDocument.Fields.Update






End Sub

Sub AppendixFields_mod26()

'counter = 1

'Do While counter <= 30

    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ ABC", PreserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    PreserveFormatting:=False
    Selection.TypeText Text:="=INT(("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1)/26)+1 \* ALPHABETIC"
    Selection.Collapse direction:=wdCollapseEnd
    Selection.MoveEnd (wdLine)
    Selection.Collapse (wdCollapseEnd)
 '   Selection.Range.InsertParagraphAfter
 '   Selection.Collapse (wdCollapseEnd)
    
    
    
'    counter = counter + 1
    
    
    
'Loop

'ActiveDocument.Fields.Update



End Sub


Sub tryThis()

counter = 1

Do While counter <= 30
    Call AppendixFields_mod26
    Call AppendixFields_regular
    Selection.Collapse (wdCollapseEnd)
    Selection.InsertParagraphAfter
    
    counter = counter + 1
Loop

'ActiveDocument.Fields.Update



End Sub


Sub nested_field()

    

End Sub



