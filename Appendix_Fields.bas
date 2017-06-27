Attribute VB_Name = "Appendix_Fields"
Function appendixField2()
    
    Dim testField As Range
    testField.Text = "SEQ ABC \c"
    
    
End Function

Function appendixFields1()

    Dim testField As Field
    testField.Data
    testField.preserveFormatting = False
    

    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ ABC \c", preserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    preserveFormatting:=False
    Selection.TypeText Text:="=MOD("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1,26)+1 \* ALPHABETIC"
    Selection.Collapse direction:=wdCollapseEnd
    Selection.MoveEnd (wdLine)
    Selection.Collapse (wdCollapseEnd)


    
End Function



Sub AppendixFields_Full()

    Dim i As Integer
    
    i = 1
    
'    Do While i <= 100
    
'    Selection.TypeText Chr(10)

    'Set A2Z
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ ABC", preserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    preserveFormatting:=False
    Selection.TypeText Text:="=MOD("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1,26)+1"
'    Selection.Fields.ToggleShowCodes
    
    Selection.MoveRight (3)
    'Selection.PreviousField
    'Selection.PreviousField
'    Selection.MoveUp Unit:=wdLine, count:=1
    
'    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.PreviousField
    Selection.PreviousField
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText Text:="SET A2Z"
    
'    Selection.MoveUp Unit:=wdLine, count:=1
    
'    Selection.Delete Unit:=wdCharacter, count:=1
    
    'Selection.PreviousField
    Selection.NextField
    Selection.PreviousField
    
    Selection.Collapse (wdCollapseEnd)
    

    'Set AA2ZZ
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "SEQ ABC \c", preserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    preserveFormatting:=False
    Selection.TypeText Text:="=INT(("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1)/26)"
    
    Selection.MoveRight (3)
    Selection.PreviousField
    Selection.PreviousField
    
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText Text:="SET AA2ZZ"
    
    Selection.PreviousField
    Selection.NextField
    
    Selection.Collapse (wdCollapseEnd)
    
    'Set If statement
    'Selection.TypeText (" ")
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "AA2ZZ \* ALPHABETIC", preserveFormatting:=False
    Selection.Fields.ToggleShowCodes
'    Selection.PreviousField
    Selection.NextField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    preserveFormatting:=False
    Selection.TypeText Text:="IF"
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="= " & Chr(34) & " " & Chr(34) & " " & Chr(34) & Chr(34)
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "AA2ZZ \* ALPHABETIC", preserveFormatting:=False
    
    'Selection.MoveRight (3)
  '  Selection.PreviousField
  Selection.Fields.ToggleShowCodes
    Selection.NextField
    
    
    
    
    Selection.Collapse (wdCollapseEnd)
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "A2Z \* ALPHABETIC", preserveFormatting:=False
    Selection.PreviousField
    'Selection.NextField
    Selection.Collapse (wdCollapseEnd)
    
    Selection.MoveLeft Unit:=wdCharacter, count:=4, Extend:=wdExtend
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText ("QUOTE")
    
    Selection.Expand (wdParagraph)
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeText (Chr(10))
    
    i = i + 1
    
'    Loop

    

'ActiveDocument.Fields.Update






End Sub

Sub AppendixFields_mod26()

'counter = 1

'Do While counter <= 30

    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ ABC", preserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    preserveFormatting:=False
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

''Currently this one isn't working
''I need to figure out how to get it all nested together because
' it's entirely too much work!
Sub field_range()
    Dim rng As Range
    
    Set rng = ActiveDocument.Paragraphs.Last

    rng.Fields.Add Range:=rng, Type:=wdFieldEmpty, Text:= _
        "SEQ ABC", preserveFormatting:=False
'    rng.PreviousField
    rng.Fields.Add Range:=rng, Type:=wdFieldEmpty, _
    preserveFormatting:=False
    rng.inser Text:="=INT(("
    rng.NextField
    rng.Collapse direction:=wdCollapseEnd
    rng.Text Text:="-1)/26)+1 \* ALPHABETIC"
    rng.Collapse direction:=wdCollapseEnd
    rng.MoveEnd (wdLine)
    rng.Collapse (wdCollapseEnd)



End Sub


Sub tryThis()



'counter = 1
'
'Do While counter <= 30
'    Call AppendixFields_mod26
'    Call AppendixFields_regular
'    Selection.Collapse (wdCollapseEnd)
'    Selection.InsertParagraphAfter
'
'    counter = counter + 1
'Loop

'ActiveDocument.Fields.Update



End Sub


Sub nestedFields()


'counter = 1

'Do While counter <= 30

    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ ABC", preserveFormatting:=False
    Selection.TypeText Text:="SET A2Z"
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText Text:="=MOD("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1,26)+1 \* ALPHABETIC"
    Selection.Collapse direction:=wdCollapseEnd
    
    'Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText Text:="QUOTE "
    Selection.NextField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
'    Selection.Collapse Direction:=wdCollapseStart
    
    
    
    Selection.NextField
    Selection.NextField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False

'    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
'    "SEQ ABC", preserveFormatting:=False
'    Selection.PreviousField
'    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
'    preserveFormatting:=False
'    Selection.TypeText Text:="=INT(("
'    Selection.NextField
'    Selection.Collapse direction:=wdCollapseEnd
'    Selection.TypeText Text:="-1)/26)+1 \* ALPHABETIC"
'    Selection.Collapse direction:=wdCollapseEnd
'    Selection.MoveEnd (wdLine)
'    Selection.Collapse (wdCollapseEnd)
' '   Selection.Range.InsertParagraphAfter
' '   Selection.Collapse (wdCollapseEnd)
    
    
    
'    counter = counter + 1
    
    
    
'Loop

'ActiveDocument.Fields.Update

    

End Sub

Sub selectField()

'Selection.NextField





End Sub

Sub newtrial()

Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
Selection.TypeText "IF"
Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="AA2ZZ \* ALPHABETIC", preserveFormatting:=False

Selection.Fields.ToggleShowCodes
Selection.Fields.ToggleShowCodes
Selection.NextField
Selection.NextField
Selection.Collapse (wdCollapseEnd)

'Selection.Collapse (wdCollapseEnd)
Selection.TypeText Text:="= " & Chr(34) & " " & Chr(34) & " " & Chr(34) & Chr(34)
Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="AA2ZZ \* ALPHABETIC", preserveFormatting:=False


End Sub
