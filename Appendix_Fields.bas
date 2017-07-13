Attribute VB_Name = "Appendix_Fields"
Sub AppendixFields_Full()

    Application.ScreenUpdating = False
    
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = True

    Dim i As Integer
    
'    i = 1
'
'    Do While i <= 50
'
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
    "SEQ Append1", PreserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    PreserveFormatting:=False
    Selection.TypeText Text:="=MOD("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1,26)+1"
    
    Selection.MoveRight (3)
    Selection.PreviousField
    Selection.PreviousField
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="SET A2Z"
    
    Selection.NextField
    Selection.PreviousField
    
    Selection.Collapse (wdCollapseEnd)
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "SEQ Append2", PreserveFormatting:=False
    Selection.PreviousField
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    PreserveFormatting:=False
    Selection.TypeText Text:="=INT(("
    Selection.NextField
    Selection.Collapse direction:=wdCollapseEnd
    Selection.TypeText Text:="-1)/26)"
    
    Selection.MoveRight (3)
    Selection.PreviousField
    Selection.PreviousField

    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="SET AA2ZZ"
    
    Selection.PreviousField
    Selection.NextField
    
    Selection.Collapse (wdCollapseEnd)
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText ("IF=" & Chr(34) & Chr(34) & " " & Chr(34) & Chr(34))
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "AA2ZZ \* ALPHABETIC", PreserveFormatting:=False
'    Selection.Fields.ToggleShowCodes
    Selection.PreviousField
    Selection.NextField
    Selection.Collapse (wdCollapseStart)
    Selection.MoveRight Unit:=wdCharacter, count:=4

    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "AA2ZZ \* ALPHABETIC", PreserveFormatting:=False
    
    
    Selection.PreviousField
    Selection.NextField
    Selection.Collapse (wdCollapseEnd)
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "A2Z \* ALPHABETIC", PreserveFormatting:=False
    Selection.PreviousField
    Selection.NextField
    Selection.Collapse (wdCollapseEnd)
    
    Selection.MoveLeft Unit:=wdCharacter, count:=4, Extend:=wdExtend
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText ("QUOTE")
'
'    Selection.Expand (wdParagraph)
'    Selection.Collapse (wdCollapseEnd)
'    Selection.TypeText (Chr(10))
'
'    i = i + 1
'
'   ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
'
'    Loop

'    Application.ScreenUpdating = True



ActiveDocument.Fields.Update

End Sub


Sub clear_appendix_numbers()

    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    
    With Selection.find
       .Text = "Appendix ^$"
        .Replacement.Text = "Appendix"
        .Forward = True
        .Wrap = wdFindContinue
'        .format = False
        .MatchCase = False
        .MatchWholeWord = False
'        .MatchKashida = False
'        .MatchDiacritics = False
'        .MatchAlefHamza = False
'        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.HomeKey Unit:=wdStory
    
    Selection.find.Execute Replace:=wdReplaceAll
    
End Sub

Sub AddFieldsAppendices()
'
' Must place cursor at first instance of "Appendix", between the x and the period.
' After this runs, need to manually fix 'Appendix AA'
'
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    
    With Selection.find
       .Text = "Appendix"
        .Forward = True
        .Wrap = wdFindContinue
'        .format = False
        .MatchCase = False
        .MatchWholeWord = False
'        .MatchKashida = False
'        .MatchDiacritics = False
'        .MatchAlefHamza = False
'        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.HomeKey Unit:=wdStory
    
    Selection.find.Execute
    Selection.Collapse (wdCollapseEnd)
'    Selection.TypeText (" ")
    'With Selection.Find.Replacement.Font
    '    .Size = 10
    '    .Bold = True
    '    .Italic = False
    '    .Color = wdColorAutomatic
    'End With
    Call AppendixFields_Full
    
    
    
    With Selection.find
        .Text = "Appendix "
        .Replacement.Text = "Appendix ^c"
        .Forward = True
        .Wrap = wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute Replace:=wdReplaceAll
    
    Selection.WholeStory
    Selection.Fields.Update
    
End Sub


Sub AddFieldsAppendices_EM()

'    Dim appendField As Range

    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    
    With Selection.find
       .Text = "Appendix"
        .Forward = True
        .Wrap = wdFindStop
'        .format = False
        .MatchCase = True
        .MatchWholeWord = False
'        .MatchKashida = False
'        .MatchDiacritics = False
'        .MatchAlefHamza = False
'        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    'Move the beginning of the doucment
    Selection.HomeKey Unit:=wdStory
    
    Selection.find.Execute
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeText (" ")
    'With Selection.Find.Replacement.Font
    '    .Size = 10
    '    .Bold = True
    '    .Italic = False
    '    .Color = wdColorAutomatic
    'End With
    Call AppendixFields_Full
    
    Selection.Collapse (wdCollapseStart)
    
    Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Selection.Copy
    
    Selection.Collapse (wdCollapseEnd)
    
'    Selection.Copy
    
    Selection.find.Execute
    
    Do While Selection.find.Found = True
        Selection.Collapse (wdCollapseEnd)
        Selection.TypeText (" ")
        Selection.Paste
        Selection.Collapse (wdCollapseEnd)
        Selection.find.Execute
    Loop
        
    
    ActiveDocument.Fields.Update
    
End Sub

Sub DeleteTextEntrySentence()

Dim tbl As Table

    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "text entry "
        .Forward = False
        .Wrap = wdFindStop
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
    End With
    
For Each tbl In ActiveDocument.Tables


    
    
    tbl.Select
    
    Selection.find.Execute
    
    If Selection.find.Found = True Then
        Selection.Expand (wdSentence)
        Selection.Delete
    End If
    
    Selection.Collapse
    
    Next
    
End Sub

Sub DeleteRowWithSpecifiedText()
    Dim sText As String

    sText = "Refer to the Display Logic panel for this question's logic."
    Selection.find.ClearFormatting
    With Selection.find
        .Text = sText
        .Wrap = wdFindContinue
    End With
    Do While Selection.find.Execute
        If Selection.Information(wdWithInTable) Then
            Selection.Rows.Delete
        End If
    Loop
End Sub

Sub RedoAppendixNumbering()
'Find and replace all the current appendix numbering

    Selection.HomeKey Unit:=wdStory
    
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
       .Text = "Appendix^$"
        .Replacement.Text = "Appendix"
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute Replace:=wdReplaceAll
End Sub



Sub insert_Appendix_cross_ref()
'
' Appendix_cross_ref Macro
'
    Selection.HomeKey Unit:=wdStory
    
    Dim appendCounter As Integer
    appendCounter = 1
    Dim appendRef As String
    
    Selection.find.ClearFormatting
    With Selection.find
        .Text = "See Appendix "
        .Forward = True
        .Wrap = wdFindStop
    End With
    
    Selection.find.Execute
    
    Do While Selection.find.Found = True
        appendRef = "Append" & Trim(Str(appendCounter))
        Debug.Print appendRef

        Selection.Collapse (wdCollapseEnd)

        Selection.InsertCrossReference ReferenceType:="Bookmark", ReferenceKind:= _
        wdContentText, ReferenceItem:=appendRef, InsertAsHyperlink:=True, _
        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
        
        appendCounter = appendCounter + 1
        
        Selection.find.Execute
    Loop
        
End Sub

Sub Appendix_cross_ref()
'
' Appendix_cross_ref Macro
'
'
    Selection.MoveLeft Unit:=wdCharacter, count:=10, Extend:=wdExtend
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Append1"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Append2"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    Selection.InsertCrossReference ReferenceType:="Bookmark", ReferenceKind:= _
        wdContentText, ReferenceItem:="Append1", InsertAsHyperlink:=True, _
        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
End Sub


Sub Appendix_bookmarks_tag()

Dim tbl As Table
Dim exportTag As String
exportTag = ""

For Each tbl In ActiveDocument.Sections(2).Range.Tables
    Selection.find.ClearFormatting
    Selection.find.Text = "Export Tag: "
    tbl.Select
    Selection.find.Execute
    If Selection.find.Found = True Then
        Selection.Collapse (wdCollapseEnd)
        Selection.MoveRight Unit:=wdWord, count:=1, Extend:=True
        exportTag = Selection.Range.Text
        Debug.Print exportTag
    End If
    Selection.find.ClearFormatting
    Selection.find.Text = "Appendix "
    tbl.Select
    Selection.find.Execute
    If Selection.find.Found = True And Not exportTag = "" Then
        Selection.Collapse (wdCollapseStart)
        Selection.MoveRight Unit:=wdWord, count:=2, Extend:=True
        ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:=exportTag
    End If
    exportTag = ""
    
Next

End Sub

Sub add_appendix_ref_to_body()

Dim hasAppendix As Boolean
Dim exportTag As String

For Each tbl In ActiveDocument.Sections(1).Range.Tables
    hasAppendix = False
    exportTag = ""
    Selection.find.ClearFormatting
    Selection.find.Text = "See Appendix"
    tbl.Select
    Selection.find.Execute
    If Selection.find.Found = True Then
        hasAppendix = True
    End If
    
    If hasAppendix = True And tbl.Columns.count = 1 Then
        Selection.find.Text = "Export Tag: "
        tbl.Select
        Selection.find.Execute
        If Selection.find.Found = True Then
            Selection.Collapse (wdCollapseEnd)
            Selection.MoveRight Unit:=wdWord, count:=1, Extend:=True
            exportTag = Selection.Range.Text
            Debug.Print exportTag
        End If
        
    ElseIf hasAppendix = True And tbl.Columns.count > 1 Then
        Selection.Previous(wdTable, 1).Select
        Selection.find.Text = "Export Tag: "
        Selection.find.Execute
        If Selection.find.Found = True Then
        Selection.Collapse
            Selection.MoveRight Unit:=wdWord, count:=1, Extend:=True
            exportTag = Selection.Range.Text
            Debug.Print exportTag
        End If
        
    End If
    If Not exportTag = "" Then
        tbl.Select
        Selection.find.Text = "Appendix"
        Selection.find.Execute
        Selection.InsertCrossReference ReferenceType:="Bookmark", ReferenceKind:= _
            wdContentText, ReferenceItem:=exportTag, InsertAsHyperlink:=True, _
            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
    End If
    
       
    Next
    
 '       On Error Resume Next
        
    
    
End Sub
