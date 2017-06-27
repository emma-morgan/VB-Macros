Attribute VB_Name = "Appendix_Fields"
Sub AppendixFields_Full()

    Application.ScreenUpdating = False

    Dim i As Integer
    
    i = 1
    
    Do While i <= 100
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ ABC", preserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    preserveFormatting:=False
    Selection.TypeText Text:="=MOD("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1,26)+1"
    
    Selection.MoveRight (3)
    Selection.PreviousField
    Selection.PreviousField
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText Text:="SET A2Z"
    
    Selection.NextField
    Selection.PreviousField
    
    Selection.Collapse (wdCollapseEnd)
    
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
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText ("IF=" & Chr(34) & " " & Chr(34) & " " & Chr(34) & Chr(34))
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "AA2ZZ \* ALPHABETIC", preserveFormatting:=False
'    Selection.Fields.ToggleShowCodes
    Selection.PreviousField
    Selection.NextField
    Selection.Collapse (wdCollapseStart)
    Selection.MoveRight Unit:=wdCharacter, count:=4
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    preserveFormatting:=False
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "AA2ZZ \* ALPHABETIC", preserveFormatting:=False
    
    
    Selection.PreviousField
    Selection.Collapse (wdCollapseEnd)
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "A2Z \* ALPHABETIC", preserveFormatting:=False
    Selection.PreviousField
    Selection.NextField
    Selection.Collapse (wdCollapseEnd)
    
    Selection.MoveLeft Unit:=wdCharacter, count:=4, Extend:=wdExtend
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        preserveFormatting:=False
    Selection.TypeText ("QUOTE")
    
    Selection.Expand (wdParagraph)
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeText (Chr(10))
    
    i = i + 1
    
    Loop

    Application.ScreenUpdating = True

ActiveDocument.Fields.Update

End Sub


