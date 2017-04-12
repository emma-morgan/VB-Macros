Attribute VB_Name = "OldQualtricsToolsMacros"
Sub Add_Extra_Borders()
'Wasn't sure where this was called as it looked like it wasn't being called by any other macros
    
    With ActiveDocument
            Dim nTables As Long
            nTables = .Tables.count
        
        For i = 1 To nTables
            nrow = .Tables(i).Rows.count
            'Determine page of first row in table
            .Tables(i).Rows(1).Select
            FirstRowPage = Selection.Information(wdActiveEndPageNumber)
            
            'Determine page of last row in table
            .Tables(i).Rows.Last.Select
            LastRowPage = Selection.Information(wdActiveEndPageNumber)
            
            'If table spans more than one page
            If FirstRowPage <> LastRowPage Then
                For j = 1 To nrow
                   If j <> nrow Then
                        .Tables(i).Rows(j).Select
                        CurRowPage = Selection.Information(wdActiveEndPageNumber)
                        .Tables(i).Rows(j + 1).Select
                        NextRowPage = Selection.Information(wdActiveEndPageNumber)
                        
                        If CurRowPage <> NextRowPage Then
                            .Tables(i).Rows(j).Select
                            Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                        End If
                   End If
                Next
            End If
        Next
    End With
End Sub

Sub preview_page_breaks()

'Wasn't sure if this was similar or the same as the macro I was working on last week? the comments
'suggest that they are pretty similar



'This macro will iterate through and ensure that questions and tables are on the same page
'The macro can be run multiple times
'If running a second time, the macro will first iterate through and check for page breaks
'Existing breaks will be removed, and additional breaks added to keep everything running
'   smoothly and in order of what should be happenning.


    With ActiveDocument
    Dim nTables As Long
    nTables = .Tables.count
        
    For i = 1 To nTables
    
        nrow = .Tables(i).Rows.count
        ncol = .Tables(i).Columns.count
                
        'Question tables will have 1 row, others will have multiple
        
        Debug.Print "Table " + Str(i) + "(" + Str(nrow) + "x"; Str(ncol) + ")"
        
        If ncol > 1 Then
            .Tables(i).Rows(1).Select
        
            answerRow1 = Selection.Information(wdActiveEndPageNumber)
        
            .Tables(i).Rows(nrow).Select
            answerRowN = Selection.Information(wdActiveEndPageNumber)
            
            .Tables(i - 1).Rows(1).Select
            questionRow1 = Selection.Information(wdActiveEndPageNumber)
            
            qRows = .Tables(i - 1).Rows.count
            .Tables(i - 1).Rows(qRows).Select
            questionRowN = Selection.Information(wdActiveEndPageNumber)
            
            If .Tables(i - 1).Columns.count <> 1 Then
                Debug.Print "Previous is not question; table " + Str(i - 1)
            
            Else
                If questionRow1 <> answerRowN Then
                    Debug.Print "Table " + Str(i)
                    Debug.Print "Question: " + Str(questionRowN) + "-" + Str(questionRowN)
                    Debug.Print "Table: " + Str(answerRowN) + "-" + Str(answerRowN)
                    
                   .Tables(i - 1).Rows(1).Select
                    Selection.InsertBreak (wdPageBreak)
                End If
            
                
            
            End If
        
       End If
    Next
    
    End With

End Sub


Sub clear_page_breaks()
'Same as the above, unsure if this is different from the macros I was working on last week as it looks like this
'is paired with the macro above, and I wrote a pair of macros very similar

nsection = ActiveDocument.Sections.count
Debug.Print (nsection)
For i = 1 To nsection

    ActiveDocument.Sections(i).Range.Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^m"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchPrefix = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll
Next

End Sub




Sub appendix_table_formatting_CBB()
'CBB? Didn't see it called anywhere else and you mentioned that CBB macros may be otudated


'Created by CB;

    With ActiveDocument
    
    
    Set entireDoc = .Range
    With entireDoc
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Color = wdColorAutomatic
        .ParagraphFormat.Alignment = wdAlignLeft
    End With
    
    Dim nTables As Long
    nTables = .Tables.count
    Debug.Print nTables
    
    For i = 1 To nTables
        nrow = .Tables(i).Rows.count
        Debug.Print nrow
        'set widths for each table
        .Tables(i).PreferredWidthType = wdPreferredWidthPercent
        .Tables(i).PreferredWidth = 100
        
        'format first 6 rows
        For j = 1 To nrow
            .Tables(i).Rows(j).Select
            If j < 4 Then
                With Selection
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                    .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                End With
            ElseIf j = 4 Then
                Selection.Font.Italic = True
                Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            ElseIf j = 6 Then
                With Selection
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphLeft
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                End With
            
            ElseIf j > 6 Then
                forShading = j Mod 2
                With Selection
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    .Shading.BackgroundPatternColor = -738132173
                End With
                If forShading = 0 Then
                    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
                End If
                If j = nrow Then
                    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                End If
            End If
            
        Next
    
    
    Next
 
    
    End With


End Sub



Sub BorderAtBreak()
'Didn't see this macro being called anywhere and the comments suggest that the macro may be unfinished?

'This tells me the first row after the page break
'I need to expand this to find a row after EACH page break
'Which will provide the formatting I need

Dim nTables As Long
    nTables = ActiveDocument.Tables.count
For t = 1 To nTables

    Dim r As Range
    Dim tblStartPage As Long
    Dim oTable As Table
    Dim oTableRange As Range
    Dim oRow As row
    Dim msg As String
    Dim bottomCell

    Dim i As Long
    i = 0

    Dim pbCellsArray() As Integer




' get the page os the start of the table
    Set oTable = ActiveDocument.Tables(t)
    Set oTableRange = oTable.Range
       oTableRange.Collapse 1
       tblStartPage = _
          oTableRange.Information(wdActiveEndPageNumber)
    ' loop through each row checking if it is the same page
    For Each oRow In oTable.Rows
       Set r = oRow.Range
       r.Collapse 1
       If r.Information(wdActiveEndPageNumber) <> _
          tblStartPage Then
          b = oRow.Index - 1
          Set bottomCell = ActiveDocument.Range(Start:=oTable.Cell(b, 1).Range.Start, _
                End:=oTable.Cell(b + 1, 1).Range.End)
    
            With bottomCell.Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleSingle
                .Color = wdColorAutomatic
            End With
          
          i = i + 1
          
          tblStartPage = tblStartPage + 1
       
       End If
    Next
    


Next

End Sub
Sub fix_page_breaks_CBB_orig()
'There was another fix_page_breaks macro that didn't have CBB_orig attached to it so I assumed that was a more updated
'version of this macro.

'Macro written by CBB to adjust page breaks in appendix tables
' Will need to adapt code to work with preview tables as well

    With ActiveDocument
        Dim nTables As Long
        nTables = .Tables.count
    
    For i = 1 To nTables
        nrow = .Tables(i).Rows.count
        
        'Determine page of first row in table
        .Tables(i).Rows(1).Select
        FirstRowPage = Selection.Information(wdActiveEndPageNumber)
        
        'Determine which page the first row of comments starts on
        .Tables(i).Rows(6).Select
        FirstCommentPage = Selection.Information(wdActiveEndPageNumber)
        
        'Determine page of last row in table
        .Tables(i).Rows.Last.Select
        LastRowPage = Selection.Information(wdActiveEndPageNumber)
        
        'If header is split between two pages
        If FirstRowPage <> FirstCommentPage Then
            'figure out which row has page break
            For j = 1 To 5
                .Tables(i).Rows(j).Select
                CurRowPage = Selection.Information(wdActiveEndPageNumber)
                Debug.Print CurRowPage

                If CurRowPage = FirstCommentPage Then
                    Exit For
                End If
            Next
            For k = 1 To (j - 1)
                .Tables(i).Rows(1).Select
                 Selection.MoveUp unit:=wdLine, count:=1
                 Selection.TypeParagraph
            Next
        End If

        'If table spans more than one page
        If FirstRowPage <> LastRowPage Then
            'The first 6 cells will be repeated on every page
            Set rptHeadCells = .Range(Start:=.Tables(i).Cell(1, 1).Range.Start, _
                End:=.Tables(i).Cell(5, 1).Range.End)
            'Select the entire table
            Set myTable = .Range(Start:=.Tables(i).Cell(1, 1).Range.Start, _
                End:=.Tables(i).Cell(nrow, 1).Range.End)
            'Make the first 6 rows into a header that will repeat across pages
                rptHeadCells.Rows.HeadingFormat = True
                myTable.Rows.AllowBreakAcrossPages = False
         End If
    Next
    
    End With
End Sub

Sub DefaultParagraphSpacing()
'Didn't see this being called anywhere so unsure if it is beign used


With ActiveDocument
    .Paragraphs.SpaceAfterAuto = False
    .Paragraphs.SpaceAfter = 0
    .Paragraphs.SpaceBeforeAuto = False
    .Paragraphs.SpaceBefore = 0
End With
    
End Sub

Sub removeTotalRowCoded()
'Didn't see this beign called anyhwere so unsure if it is beign used

With ActiveDocument
    
    Dim loopCount As Integer
    loopCount = 1
    
    Dim nTables As Integer
    Dim nrow As Integer
        
    nTables = ActiveDocument.Tables.count
    
    For i = 1 To nTables
        
        nrow = ActiveDocument.Tables(i).Rows.count
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = .Styles("Heading 5")
    With Selection.Find
     .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.HomeKey unit:=wdStory
    Selection.Find.Execute
    
    Do While Selection.Find.Found = True And loopCount < 1000
    
        Debug.Print iCount
        Selection.Expand wdParagraph
        Selection.Delete
        Selection.EndOf
        Selection.HomeKey unit:=wdStory
        Selection.Find.Execute
    Loop
    
    
    
    End With


End Sub


Sub left_right_padding_change()
'Didn't see this being called anywhere so unsure if it was being used

    Dim i As Integer
    Dim ncol As Integer
    Dim nTables As Integer
    
    With ActiveDocument
    
    nTables = .Tables.count
        
    For i = 1 To nTables
        ncol = .Tables(i).Columns.count
    
        If ncol = 1 Then
        
            With .Tables(i)
                .Spacing = InchesToPoints(0)
                .TopPadding = InchesToPoints(0)
                .BottomPadding = InchesToPoints(0)
                .LeftPadding = InchesToPoints(0)
                .RightPadding = InchesToPoints(0)
                
            End With
        End If
        
    Next
    
    End With

End Sub

Sub remove_extra_carraigeReturn()
'This was replaced by an different version I wrote a while back

    With ActiveDocument
    
    Dim para As Paragraph
    Dim i As Integer
    Dim rng As Range
    Dim nextPar As Range
    Dim j As Integer
    Dim loopCount As Integer
   
    nParagraphs = .Paragraphs.count
    
    i = 1
    
    Do While i < .Paragraphs.count
    
        Debug.Print ("Paragraph: " & i)
    
        Set rng = .Paragraphs(i).Range
        rng.Select
        Debug.Print (rng)
        
        If Selection.Text = Chr(13) Or Selection.Text = Chr(160) & Chr(13) Then
        
            Debug.Print ("True: " & i)
            If i = .Paragraphs.count Then End
            j = i + 1
            Set nextPar = .Paragraphs(j).Range
            nextPar_Text = nextPar.Text
            loopCount = 1
            
            Do While (nextPar_Text = Chr(13) Or nextPar_Text = Chr(160) & Chr(13)) And loopCount < 10 And j <= .Paragraphs.count
            
                nextPar.Select
                Selection.Delete
                Selection.EndOf
                nextPar_Text = .Paragraphs(j).Range.Text
                loopCount = loopCount + 1
            Loop
                            
            Selection.EndOf
                    
        End If
        
        i = i + 1
    
    Loop
   
   End With
    
End Sub


Sub UpdateDocuments()
'This update documents macro didn't work properly so there is a more updated version I believe

Application.ScreenUpdating = False
Dim strFolder As String, strFile As String, wdDoc As Document
strFolder = GetFolder
If strFolder = "" Then Exit Sub
strFile = Dir(strFolder & "\*.docx", vbNormal)
While strFile <> ""
  Set wdDoc = Documents.Open(FileName:=strFolder & "\" & strFile, AddToRecentFiles:=False, Visible:=False)
  With wdDoc
    'Call your other macro or insert its code here
    Call define_table_styles
    Call format_appendix
    .Close SaveChanges:=True
  End With
  strFile = Dir()
Wend
Set wdDoc = Nothing
Application.ScreenUpdating = True
End Sub
 
Function GetFolder() As String
'Same as the function above, they are paired together

Dim oFolder As Object
GetFolder = ""
Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "Choose a folder", 0)
If (Not oFolder Is Nothing) Then GetFolder = oFolder.Items.Item.path
Set oFolder = Nothing
End Function

Sub Fix_tables()
'I believe that this macro was implemented into prior macros so it did it simultaneously

With ActiveDocument

Dim nTables As Long
nTables = .Tables.count
'ActiveDocument.Styles.Add Name:="ResponseTables2", Type:=wdStyleTypeParagraph
    With ActiveDocument.Styles("ResponseTables").ParagraphFormat
        .LeftIndent = InchesToPoints(0.08)
        .RightIndent = InchesToPoints(0.08)
    End With

For i = 1 To nTables
    Dim nRows As Long
    Dim nCols As Long
    
    nRows = .Tables(i).Rows.count
    nCols = .Tables(i).Columns.count
    
    
    For j = 1 To nRows
        For k = 1 To nCols
            .Tables(i).Cell(j, k).TopPadding = 0
            .Tables(i).Cell(j, k).BottomPadding = 0
            .Tables(i).Cell(j, k).LeftPadding = 0
            .Tables(i).Cell(j, k).RightPadding = 0

        Next
    Next
    
    If nCols > 1 Then
        .Tables(i).Rows.Select
        Selection.Style = ActiveDocument.Styles("ResponseTables")
    End If
    '.Tables(i).Range.ParagraphStyle = "Heading 1"
    '.Tables(i).Range.Paragraphs(i).Style = ResponseTables
    
Next

End With
End Sub

Sub AdjustHeading()
'This was for trying to get the headers on new pages if necessary but wasn't working properly but
'haven't really looked deeply into the code

    Dim CursorVert As Single

    Dim Pgheight As Single
    Dim styleName As String

    styleName = "Heading 3"
    If ActiveDocument.Styles(styleName).ParagraphFormat.PageBreakBefore Then
        MsgBox styleName & " has 'Page break before' set. Run aborted"
        Exit Sub
    End If
    Selection.HomeKey unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(styleName)
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
    End With
    Do While Selection.Find.Execute
        If Asc(Selection.Range.Characters(1)) = 12 Then
            Selection.MoveStart unit:=wdCharacter, count:=1
        End If
        With ActiveDocument.Sections(Selection.Information(wdActiveEndSectionNumber)).PageSetup
            CursorVert = Selection.Information(wdVerticalPositionRelativeToPage) - .TopMargin
            Pgheight = .PageHeight - .TopMargin - .BottomMargin
        End With
        If CursorVert > Selection.Style.ParagraphFormat.SpaceBefore Then
            If CursorVert / Pgheight > 0.66 And Len(Selection.Range) > 1 Then
                Selection.End = Selection.Start
                Selection.TypeText Chr(12)
            End If
        End If
        Selection.Start = Selection.End
    Loop
End Sub

Sub SelectHeadingandContent()
'Same as above
'This was for trying to get the headers on new pages if necessary but wasn't working properly but
'haven't really looked deeply into the code

Dim headStyle As Style

' Checks that you have selected a heading. If you have selected multiple paragraphs, checks only the first one. If you have selected a heading, makes sure the whole paragraph is selected and records the style. If not, exits the subroutine.

If ActiveDocument.Styles(Selection.Paragraphs(1).Style).ParagraphFormat.OutlineLevel < wdOutlineLevelBodyText Then
    Set headStyle = Selection.Paragraphs(Selection.Paragraphs.count).Style
    Selection.Expand wdParagraph
Else: Exit Sub
End If


' Loops through the paragraphs following your selection, and incorporates them into the selection as long as they have a higher outline level than the selected heading (which corresponds to a lower position in the document hierarchy). Exits the loop if there are no more paragraphs in the document.

Do While ActiveDocument.Styles(Selection.Paragraphs(Selection.Paragraphs.count).Next.Style).ParagraphFormat.OutlineLevel > headStyle.ParagraphFormat.OutlineLevel
    Selection.MoveEnd wdParagraph
    If Selection.Paragraphs(Selection.Paragraphs.count).Next Is Nothing Then Exit Do
Loop

' Turns screen updating back on.


End Sub

Sub loopHeadings()
'Same as above
'This was for trying to get the headers on new pages if necessary but wasn't working properly but
'haven't really looked deeply into the code

    Dim Headings As Variant
    Headings = _
        ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
    Debug.Print LBound(Headings)
    Debug.Print UBound(Headings)
    ActiveDocument.ActiveWindow.View.CollapseAllHeadings
    
    

End Sub

Sub ReadPara()
'Same as above
'This was for trying to get the headers on new pages if necessary but wasn't working properly but
'haven't really looked deeply into the code

    Dim DocPara As Paragraph

    For Each DocPara In ActiveDocument.Paragraphs

     If Left(DocPara.Range.Style, Len("Heading 3")) = "Heading 3" Then

       Debug.Print DocPara.Range.Text

     End If

    Next


End Sub

Public Sub CreateOutline()
'Same as above
'This was for trying to get the headers on new pages if necessary but wasn't working properly but
'haven't really looked deeply into the code'

    Dim docOutline As Word.Document
    Dim docSource As Word.Document
    Dim rng As Word.Range

    Dim astrHeadings As Variant
    Dim strText As String
    Dim intLevel As Integer
    Dim intItem As Integer

    Set docSource = ActiveDocument
    Set docOutline = Documents.Add

    ' Content returns only the
    ' main body of the document, not
    ' the headers and footer.
    Set rng = docOutline.Content
    astrHeadings = _
     docSource.GetCrossReferenceItems(wdRefTypeHeading)

    For intItem = LBound(astrHeadings) To UBound(astrHeadings)
        ' Get the text and the level.
        strText = Trim$(astrHeadings(intItem))
        intLevel = GetLevel(CStr(astrHeadings(intItem)))

        ' Add the text to the document.
        rng.InsertAfter strText & vbNewLine

        ' Set the style of the selected range and
        ' then collapse the range for the next entry.
        rng.Style = "Heading " & intLevel
        rng.Collapse wdCollapseEnd
    Next intItem
End Sub

Private Function GetLevel(strItem As String) As Integer
'Same as above
'This was for trying to get the headers on new pages if necessary but wasn't working properly but
'haven't really looked deeply into the code

    ' Return the heading level of a header from the
    ' array returned by Word.

    ' The number of leading spaces indicates the
    ' outline level (2 spaces per level: H1 has
    ' 0 spaces, H2 has 2 spaces, H3 has 4 spaces.

    Dim strTemp As String
    Dim strOriginal As String
    Dim intDiff As Integer

    ' Get rid of all trailing spaces.
    strOriginal = RTrim$(strItem)

    ' Trim leading spaces, and then compare with
    ' the original.
    strTemp = LTrim$(strOriginal)

    ' Subtract to find the number of
    ' leading spaces in the original string.
    intDiff = Len(strOriginal) - Len(strTemp)
    GetLevel = (intDiff / 2) + 1
End Function


