Attribute VB_Name = "NewMacros3"
Sub define_table_styles()

    'After defining table styles, you MUST  edit table style
        'to uncheck "allow spacing between cells" box!

    Call Define_Matrix_Style
    Call define_appendix_table_style
    Call define_basic_table_style


End Sub



Sub format_survey_preview()

    'This macro should be used BEFORE any manual updates to the survey preview
    
    Dim i As Integer
    Dim ncol As Integer
    Dim nrow As Integer
    Dim nTables As Integer
    
    
    'This calls the formatting macros in order

    'Change global font and spacing, format title header
    
    Call Preview_Style_Change
    
    Call replace_newline
    Call RemoveEmptyParagraphs
    
    
    Call number_of_respondents
    Call Insert_OIRE
    Call Insert_logo
    Call Insert_footer
    
    With ActiveDocument

    nTables = .Tables.Count
        

        
    For i = 1 To nTables
        ncol = .Tables(i).Columns.Count
        nrow = .Tables(i).Rows.Count
        Debug.Print ncol

        .Tables(i).AllowPageBreaks = False

        'We have one macro that will iterate through each table and perform
        'the appropriate formatting functions
        Call format_preview_tables(i, nrow, ncol)
        Call Replace_zeros(i)
        Call Replace_NaN(i)
        Call format_See_Appendix(i)

    Next
    
    End With
    
End Sub

Sub finish_clean_preview()

' This macro should be run AFTER the human components are finished
' This will number questions and delete question export tags from each table
' These macros can also easily be run separately, as long as the numbering is run first
' These apply ONLY to question info rows, so we can take advantage of this

    Dim i As Integer
    Dim ncol As Integer
    Dim nrow As Integer
    Dim nTables As Integer
    
    nTables = ActiveDocument.Tables.Count
        
    Call number_questions
    Call remove_denominatorRow
    Call Remove_Export_Tag

End Sub

Sub format_appendix()
'
' Macro that will call all the steps required to format appendix tables
'   for coded and raw text appendices

    With ActiveDocument
    
    Call Preview_Style_Change
       
    Dim nTables As Long
    nTables = .Tables.Count
'   MsgBox ("Number of Tables: " & nTables)
    Debug.Print nTables
    
    Dim i As Integer
'    i = 1
    For i = 1 To nTables
        
        Dim celltxt As String
        celltxt = .Tables(i).Cell(4, 1).Range.Text
'        Debug.Print celltxt
        If InStr(1, celltxt, "Coded Comments") Then
            isCodedComment = True
        Else
            isCodedComment = False
        End If
        
 '       Debug.Print isCodedComment
        
    
        .Tables(i).Select
        Selection.ClearParagraphAllFormatting
        Selection.EndOf
        
        nrow = .Tables(i).Rows.Count
        ncol = .Tables(i).Columns.Count
'        Debug.Print nrow
        
        'Remove text from second column of coded comment table header
        Call duplicateHeaderText(i)
            
        If (nrow >= 6) Then
            
         'set widths for each table
         .Tables(i).PreferredWidthType = wdPreferredWidthPercent
         .Tables(i).PreferredWidth = 100
         
         'Sort tables alphabetically for plain text, by N then alphabetically for coded
         Call alphabetize_table(i)
        
        .Tables(i).Style = "Appendix_style_table"
'        .Tables(i).AllowAutoFit = False
                'Fixed a problem with suddenly changing column width, but caused
                'other issues with N column size
        
        'Align text vertically to be centered
            'Ideally this would be a part of the table style, but I couldn't find it....
        .Tables(i).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        
        .Tables(i).Rows.HeightRule = wdRowHeightAuto
                
        If ncol = 1 Then
            .Tables(i).ApplyStyleLastRow = False
            .Tables(i).ApplyStyleLastColumn = False
        ElseIf ncol = 2 And isCodedComment = True Then
            'Verify that it's a coded comment table
            .Tables(i).ApplyStyleLastRow = True
            .Tables(i).ApplyStyleLastColumn = True
            .Tables(i).Columns(2).Select
            Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
            Selection.Columns.PreferredWidth = InchesToPoints(0.55)
            Selection.EndOf
        Else
            .Tables(i).ApplyStyleLastRow = False
            .Tables(i).ApplyStyleLastColumn = False
        
        End If
                 
         For j = 1 To 6
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
                 With Selection
                     .Font.Italic = True
                     .ParagraphFormat.Alignment = wdAlignParagraphCenter
                     .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                     .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                     .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                     .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                 End With
             ElseIf j = 5 Then
                 Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                 Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
                 Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                 Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
             ElseIf j = 6 Then
                 With Selection
                     .Font.Bold = True
                     .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                     .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                     .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                     .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                 End With
                 
                 If ncol = 2 Then
                    .Tables(i).Cell(j, 2).Select
                    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End If
             
             End If
             
         Next
         
        Call Appendix_Merge_Header(i)
        
        Set rptHeadCells = .Range(Start:=.Tables(i).Cell(1, 1).Range.Start, _
             End:=.Tables(i).Cell(3, ncol).Range.End)

                 'Make the first 6 rows into a header that will repeat across pages
         rptHeadCells.Rows.HeadingFormat = True

         
         'Need to add back side border to "responses" line
         'Also repeat bottom border so that it will exist if the table breaks
            'across multiple pages
         .Tables(i).Rows(3).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
         .Tables(i).Rows(3).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
         .Tables(i).Rows(3).Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
         .Tables(i).Rows(3).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    
        End If
        
            
    Next
     
    End With
    
    Call Insert_footer
    
    'Make sure the stupid footer is the correct width...
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1)
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        
    End With
    
    

End Sub


Sub Preview_Style_Change()

'First step in formatting preview
'Change global font and spacing for the document
    
    'Change paragraph spacing to have no space before or after
    'With HTML export, we need a few additional steps
    'Lauren discovered these in the senior survey; individual macros written
    ' and sent 11/17/16; incorporated 12/1/16
    
    'Specify Header 5 (block headers) to be Italic, Bold, size 14 font
    
    On Error Resume Next
    
    With ActiveDocument
    
        With .PageSetup
            .TopMargin = InchesToPoints(0.5)
            .BottomMargin = InchesToPoints(0.5)
            .LeftMargin = InchesToPoints(0.5)
            .RightMargin = InchesToPoints(0.5)
            
            .HeaderDistance = InchesToPoints(0.5)
            .FooterDistance = InchesToPoints(0.2)
            
        End With
        
        .Paragraphs.SpaceAfterAuto = False
        .Paragraphs.SpaceBeforeAuto = False
        .Paragraphs.SpaceBefore = 0
        .Paragraphs.SpaceAfter = 0
        .Paragraphs.Format.Alignment = wdAlignParagraphLeft
        
                
        'Change style of title (Heading 4), Block names (Header 5), and regular text (Compact)
                
        With .Styles("Heading 4")
            With .Font
                .Name = "Arial"
                .Size = 16
                .Color = wdColorAutomatic
            End With
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
                
        With .Styles("Heading 5").Font
            .Name = "Arial"
            .Size = 14
            .Color = wdColorAutomatic
            .Italic = True
            .Bold = True
            .Underline = False
        End With
        
        With .Styles("Compact").Font
            .Name = "Arial"
            .Size = 10
            .Color = wdColorAutomatic
        End With
        
        With .Styles("Normal")
            With .Font
                .Name = "Arial"
                .Size = 10
                .Color = wdColorAutomatic
            End With
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.SpaceBefore = 0
        End With
        
        With .Sections(1).Footers(wdHeaderFooterPrimary).Range
            .Paragraphs.SpaceBefore = 0
            .Paragraphs.SpaceAfter = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpacingSingle
        End With
        
    'Find "Number of Respondents", select line, and change font to 10
    '.Wrap = wdFindContinue will find this regardless of where the cursor is in the doc
       
    End With
    
End Sub

Sub number_of_respondents()

    With ActiveDocument
    
        With Selection.Find
            .Text = "Number of Respondents: "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
        End With
        
        Selection.Find.Execute
        
        Selection.Expand wdLine
        Selection.Font.Size = 10
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With

End Sub

Sub Insert_OIRE()
'
' Moves to the upper right hand corner and inserts, then formats, text
' This is inserted as style Heading 4 to match Survey name;
    ' this is then adjusted when we change the format of Heading 4 in Preview_Style_Change
' Created by Adam Kaminski, summer 2016
' Edits by ECM
    
    
    With ActiveDocument
        'Move to the top right of the page
        Selection.HomeKey Unit:=wdStory
        Selection.TypeParagraph
        Selection.HomeKey Unit:=wdStory
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        'Insert text
        oireName = "Office of Institutional" + Chr(10) + "Research & Evaluation" + Chr(10)
        Selection.TypeText Text:=oireName
        'Break into two lines
        'Selection.MoveLeft Unit:=wdCharacter, Count:=21
        'Selection.TypeParagraph
        'Selection.MoveRight Unit:=wdCharacter, Count:=21
        'Selection.TypeParagraph
    End With

End Sub

Sub Insert_logo()
'
' Inserts the Tufts logo in the upper left hand corner
' Created by Adam Kaminski, summer 2016
' Edits by ECM

    With ActiveDocument
        'Navigate to the top of the page
        Selection.HomeKey Unit:=wdStory
        'Pick an image via its path and insert it
        Selection.InlineShapes.AddPicture FileName:= _
        "Q:\Student Work\Emma's Student Work\Report Generation\Report Macros_Adam\tufts_logo_black.png" _
        , LinkToFile:=False, SaveWithDocument:=True
        'Select the image
        ActiveDocument.InlineShapes(1).Select
        'format the image (lock aspect ratio and adjust height)
        With Selection.InlineShapes(1)
            .LockAspectRatio = msoTrue
            .Height = 35
        End With
        'Move it to the upper left hand corner (0, 0)
        Set nShp = Selection.InlineShapes(1).ConvertToShape
        With .Shapes(1)
            .Top = 0
            .Left = 0
        End With

    End With
    
End Sub

Sub Insert_footer()
'
' Inserts a footer
'As written, assumes there is only one section; if this changes, we need to uncomment these lines

'    Dim i As Long
   ' For i = 1 To ActiveDocument.Sections.Count
'    For Each Section In ActiveDocument.Sections
'        Dim myfooter As Word.Range

    'Clear the footer if anything exists
    
    ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Select
    Selection.Delete
    
    'In the event that we are JUST using this function, we need to change the style and format
    
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range
            .Paragraphs.SpaceBefore = 0
            .Paragraphs.SpaceAfter = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpacingSingle
            .Font.Name = "Arial"
            .Font.Size = 9
    End With

    Dim footerTable As Table
    With ActiveDocument
        Set insert_footerTable = .Tables.Add(.Sections(1).Footers(wdHeaderFooterPrimary).Range, 2, 3)
                
        Dim oireFooter As String
        Dim analystFooter As String
        Dim internalUse As String
        
        oireFooter = "Office of Institutional Research & Evaluation" + _
            Chr(10) + "NAME OF SURVEY, YEAR, AND SPECIAL POPULATION (IF APPLICABLE)"
        
        analystFooter = "Prepared by: ANALYST NAME" + Chr(10) + _
            "INSERT DATE"
            
        internalUse = "**This report is intended for internal use only**"
            
        Set footerTable = .Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1)
                        
        With footerTable
        
'            .Rows.leftindent = InchesToPoints(0)
'
'            .Columns.PreferredWidthType = wdPreferredWidthPercent
'
'            .Columns(2).PreferredWidth = 13
            
 '           .Columns(1).PreferredWidth = 40

 '           .Columns(1).SetWidth ColumnWidth:=InchesToPoints(2.9), RulerStyle:=wdAdjustNone
 '           .Columns(2).SetWidth ColumnWidth:=InchesToPoints(0.7), RulerStyle:=wdAdjustProportional
 '           .Columns(3).SetWidth ColumnWidth:=InchesToPoints(2.9), RulerStyle:=wdAdjustNone
                
'            .PreferredWidthType = wdPreferredWidthPercent
 '           .PreferredWidth = 100
                
            .TopPadding = InchesToPoints(0.08)
            .BottomPadding = InchesToPoints(0)
            .LeftPadding = InchesToPoints(0)
            .RightPadding = InchesToPoints(0)
            
        
            With .Cell(1, 1).Range
                .Text = oireFooter
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
            End With
            
            .Cell(1, 2).Range.Select
            Selection.Collapse
            With Selection
                .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "PAGE ", PreserveFormatting:=True
                .TypeText Text:=" of "
                .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "NUMPAGES ", PreserveFormatting:=True
            End With
            
            .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            With .Cell(1, 3).Range
                .Text = analystFooter
                .ParagraphFormat.Alignment = wdAlignParagraphRight
            End With
            
            .Cell(2, 1).Range.Text = internalUse
            
            'Remove borders from the footer table
            
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
                
        End With
        
        'Merge cells of second row and format text to be centered and italicized
        
        Dim mrgrng As Range

        Set mrgrng = footerTable.Cell(2, 1).Range
        mrgrng.End = footerTable.Cell(2, 3).Range.End
        mrgrng.Cells.Merge
        
        footerTable.Rows(2).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.Italic = True
    
'        footerTable.Range.ParagraphFormat.LeftIndent = 0
'        footerTable.Range.ParagraphFormat.RightIndent = 0
        
'        footerTable.AutoFitBehavior (wdAutoFitWindow)
    'Return to print layout view

'    ActiveDocument.Sections(1).Footers.tabl ables(i).PreferredWidthType = wdPreferredWidthPercent
'         .Tables(i).PreferredWidth = 100

    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If

    
    End With
    
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1)
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        
        .Cell(1, 2).PreferredWidthType = wdPreferredWidthPercent
        .Cell(1, 2).PreferredWidth = 12
        
        .Cell(1, 1).PreferredWidthType = wdPreferredWidthPercent
        .Cell(1, 1).PreferredWidth = 44
        
        .Cell(1, 3).PreferredWidthType = wdPreferredWidthPercent
        .Cell(1, 3).PreferredWidth = 44
        
        .Rows.LeftIndent = InchesToPoints(0)
    End With

    
End Sub


Sub format_preview_tables(i As Integer, nrow As Integer, ncol As Integer)

    If ncol = 1 Then
        Call format_question_info(i, nrow)
    ElseIf ncol = 3 Then
        Call format_mc_singleQ(i, nrow, ncol)
    ElseIf ncol > 3 Then
        Call format_matrix_table(i, nrow, ncol)
    
    End If

End Sub

Sub define_basic_table_style()

    On Error Resume Next
    ActiveDocument.Styles("basic_table_style").Delete
    
    ActiveDocument.Styles.Add Name:="basic_table_style", Type:=wdStyleTypeTable
    
    With ActiveDocument.Styles("basic_table_style")
        With .Table

            .AllowPageBreaks = False
            .AllowBreakAcrossPage = False
            
        End With
        
    End With
        
    
End Sub

Sub format_question_info(i As Integer, nrow As Integer)

'Format question text and information

    With ActiveDocument
        .Tables(i).Style = "basic_table_style"
        
        'format the question info, identified by single column
            ' Set table width to full page
        .Tables(i).PreferredWidthType = wdPreferredWidthPercent
        .Tables(i).PreferredWidth = 100
        
        With .Tables(i)
            .Spacing = InchesToPoints(0)
            .TopPadding = InchesToPoints(0)
            .BottomPadding = InchesToPoints(0)
            .LeftPadding = InchesToPoints(0)
            .RightPadding = InchesToPoints(0)
        End With
            
        'Bold question text
        .Tables(i).Rows(2).Select
        With Selection
            .Font.Bold = True
        End With
    
        'Make display logic red to highlight
        If nrow >= 3 Then
            Dim r As Long
            For r = 3 To nrow
                .Tables(i).Rows(r).Select
                With Selection.Font
                    .Bold = True
                    .Color = wdColorDarkRed
                End With
            Next
        End If
        
    ' Stop table from breaking across page

End With
    
End Sub

Sub format_mc_singleQ(i As Integer, nrow As Integer, ncol As Integer)
'
'Sub format_mc_singleQ_testing()
'
'Dim i As Integer
'Dim nrow As Integer
'Dim ncol As Integer
'
'i = 16
'ncol = 3
'nrow = 3

' format_mc_checkall_singleQ Macro
'
    With ActiveDocument
    
        .Tables(i).Style = "basic_table_style"
        
        'Adjust cell padding for multiple choice
        With .Tables(i)
            .LeftPadding = InchesToPoints(0.08)
            .RightPadding = InchesToPoints(0.08)
            .TopPadding = InchesToPoints(0.01)
            .BottomPadding = InchesToPoints(0.01)
            .Spacing = InchesToPoints(0)

        End With
    
        .Tables(i).Select
        
        'Remove inside borders
        Selection.Borders.InsideLineStyle = wdLineStyleNone
        
        'Select N column
        'Adjust font and right align
        .Tables(i).Columns(1).Select
        With Selection
            With .Font
                .Bold = True
                .Italic = True
                .Color = wdColorGray40
            End With
            
            With .ParagraphFormat
                .Alignment = wdAlignParagraphRight
            End With
        End With
        
        'Select % column
        'Bold and right align
        .Tables(i).Columns(2).Select
        With Selection
            .Font.Bold = True
            .ParagraphFormat.Alignment = wdAlignParagraphRight
        End With
        
        'Delete first row from this type of question
        .Tables(i).Rows(1).Select
        Selection.Rows.Delete
    
    End With


End Sub


Sub Define_Matrix_Style()

    'If the style exists from a previous run, delete and redefine
    
    On Error Resume Next
    ActiveDocument.Styles("Matrix_table_style").Delete
    
    ActiveDocument.Styles.Add Name:="Matrix_table_style", Type:=wdStyleTypeTable
    
    With ActiveDocument.Styles("Matrix_table_style")
        With .Table
            .RowStripe = 1
            .ColumnStripe = 0
            .AllowPageBreaks = False
            .AllowBreakAcrossPage = False
            
            With .Condition(wdEvenRowBanding)
                With .Shading
                    .Texture = wdTextureNone
                    .ForegroundPatternColor = wdColorAutomatic
                    .BackgroundPatternColor = -738132173
                End With
            
                With .Borders(wdBorderVertical)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = wdColorAutomatic
                End With
    
                With .Borders(wdBorderLeft)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = wdColorAutomatic
                End With
            
                With .Borders(wdBorderRight)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = wdColorAutomatic
                End With
            
            End With
          
            With .Borders(wdBorderVertical)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
    
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
            With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
    
            With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
        End With
        
    End With

End Sub

Sub format_matrix_table(i As Integer, nrow As Integer, ncol As Integer)
   
    With ActiveDocument



        With .Tables(i)
            .Style = "Matrix_table_style"
            .LeftPadding = InchesToPoints(0)
            .RightPadding = InchesToPoints(0)
            .TopPadding = InchesToPoints(0.01)
            .BottomPadding = InchesToPoints(0.01)
            .Spacing = InchesToPoints(0)
            
        End With
                    
        With .Tables(i).Cell(1, 1)
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleNone
            End With
            With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleNone
            End With
        End With
        

        .Tables(i).PreferredWidthType = wdPreferredWidthPercent
        .Tables(i).PreferredWidth = 100

        
        .Tables(i).Columns(1).Select
        With Selection.Cells
            .SetWidth _
            ColumnWidth:=InchesToPoints(3.5), _
            RulerStyle:=wdAdjustNone
            '.PreferredWidthType = wdPreferredWidthPoints
            '.PreferredWidth = InchesToPoints(3.5)
            '.PreferredWidth = InchesToPercent(3.5)
        End With
        
        '.Tables(i).Width
                
        'Format N columns

        Dim nColumns As Long
        nColumns = .Tables(i).Columns.Count

        For j = 1 To nColumns
    
            .Tables(i).Columns(j).Select
            
            Selection.Find.ClearFormatting
            
            With Selection.Find
                .Text = "N"
                .MatchWholeWord = True
            End With
            Selection.Find.Execute
            
            If Selection.Find.Found = True Then
                .Tables(i).Columns(j).Select
                With Selection.Cells
                    '.PreferredWidthType = wdPreferredWidthPoints
                    '.PreferredWidth = InchesToPoints(0.47)
                    .SetWidth _
                    ColumnWidth:=InchesToPoints(0.47), _
                    RulerStyle:=wdAdjustNone
                End With
                                 
                With Selection.Font
                     .Bold = True
                     .Italic = True
                     .Color = wdColorGray40
                 End With
                 
                 With Selection.ParagraphFormat
                     .Alignment = wdAlignParagraphCenter
                 End With
                 
                 Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                 'Selection.Cells.SetWidth = 3.5
                 
                 
             End If
        Next

        
        '.Tables(i).PreferredWidthType = wdPreferredWidthPercent
        '.Tables(i).PreferredWidth = 100
        
                
        'Format percentage columns
          
       Dim PerColumns As Long
       PerColumns = .Tables(i).Columns.Count
          
       For k = 1 To PerColumns
    
        .Tables(i).Columns(k).Select
        
        Selection.Paragraphs.LeftIndent = InchesToPoints(0.08)
        Selection.Paragraphs.RightIndent = InchesToPoints(0.08)
        
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = "%"
            .MatchWholeWord = False
        End With
        
        Selection.Find.Execute
        
        
        If Selection.Find.Found = True Then
            .Tables(i).Columns(k).Select
            With Selection.Cells
                '.PreferredWidthType = wdPreferredWidthPercent
                '.PreferredWidth = InchesToPoints(
                '.SetWidth _
                'ColumnWidth:=InchesToPoints(PerColWidth), _
                'RulerStyle:=wdAdjustNone
                .PreferredWidth = None
                '.AutoFit
                '.PreferredWidthType = wdPreferredWidthAuto
                '.PreferredWidth = 0
                
            End With
                
            With Selection.Font
                .Bold = True
                .Italic = False
                .Color = wdColorAutomatic
            End With
             
            With Selection.ParagraphFormat
                .Alignment = wdAlignParagraphCenter
            End With
            
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        End If
  
        Next
        
        
        '.Tables(i).PreferredWidthType = wdPreferredWidthPercent
        '.Tables(i).PreferredWidth = 100
        
        
       'Center align test horizontal and vertical
        
       
        
        'Format header
        .Tables(i).Rows(1).Select
        
        With Selection.Font
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
        End With
        
        With Selection.ParagraphFormat
            .Alignment = wdAlignParagraphCenter
        End With
        
        With Selection.Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With

    End With
    

End Sub



Sub Replace_zeros(i As Integer)
'
' Searches for "0.0%" and replaces it with "--"
' Created by Adam Kaminsky
' Edited by EM to make sure the program didn't stop part of the way through

    Application.DisplayAlerts = False
    
'     Dim i As Integer
'     Dim nTables As Integer
'     nTables = ActiveDocument.Tables.Count
'
'    For i = 1 To nTables
    
    ActiveDocument.Tables(i).Range.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "0.0%"
        .Replacement.Text = "--"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchPrefix = True
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll

'    Next

    
End Sub

Sub Replace_NaN(i As Integer)
'
' Searches for "NaN%" resulting from denominator 0 and replaces it with "--"
' Adapted from "Replace 0" code
' Created by Adam Kaminsky
' Edited by EM to make sure the program didn't stop part of the way through

    Application.DisplayAlerts = False
    
'     Dim i As Integer
'     Dim nTables As Integer
'     nTables = ActiveDocument.Tables.Count
    
'    For i = 3 To nTables
    
    ActiveDocument.Tables(i).Range.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NaN%"
        .Replacement.Text = "--"
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
    
'    Next

    
End Sub

Sub number_questions()
'
' Numbers questions in the survey preview
' Run as part of the final cleaning macro.
'
    With ActiveDocument
    
    Dim Q As Long
    Q = 1
    
    Dim nTables As Long
    nTables = .Tables.Count

    For i = 1 To nTables
        ncol = .Tables(i).Columns.Count
        
    If ncol = 1 Then
        'delete data export tag
        qText = .Tables(i).Cell(2, 1).Range.Text
        qNum = CStr(Q)
        qTextNum = qNum + ". " + qText
        .Tables(i).Cell(2, 1).Range.Select
        'MsgBox qText
        'MsgBox Right(qText, 2)
        ' MsgBox qTextNum
        Selection.Delete
        .Tables(i).Cell(2, 1).Range.Text = Left(qTextNum, Len(qTextNum) - 2)
        .Tables(i).Cell(2, 1).Range.Select
        With Selection.Find
            .Text = "^p"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute

    Q = Q + 1
     
    End If
    Next
    
    End With

End Sub

Sub remove_denominatorRow()

    Dim i As Integer
    Dim nTables As Integer
    
    With ActiveDocument
    
    nTables = .Tables.Count

    For i = 1 To nTables
        .Tables(i).Select
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        
        With Selection.Find
            .Text = "Denominator Used:"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
'        Selection.find.Execute
        If Selection.Find.Execute Then Selection.Rows.Delete

'        Selection.Rows.Delete
    Next
    
    End With

End Sub

Sub remove_questionInfo_row()
'
' Removes question data export tags from the question info tables in the survey preview
' Called as part of the final cleaning up macro
'
    With ActiveDocument
    
    Dim nTables As Long
    nTables = .Tables.Count
    
    For i = 1 To nTables
        ncol = .Tables(i).Columns.Count
        
'        Delete first row of the question info (data export tag)
'        This will only appear in question info in the preview; all others have 3+ columns
'        This can be used for appendices to remove first row from coded and full text comments
        
        If ncol <= 2 Then
            'delete data export tag
            .Tables(i).Rows(1).Select
            Selection.Rows.Delete
                    
        End If
    Next
            
    End With
    
End Sub

Sub define_appendix_table_style()

    'If the style exists from a previous run, delete and redefine
    On Error Resume Next
    ActiveDocument.Styles("Appendix_style_table").Delete

    ActiveDocument.Styles.Add Name:="Appendix_style_table", Type:=wdStyleTypeTable

    With ActiveDocument.Styles("Appendix_style_table")
        With .Font
            .Name = "Arial"
            .Size = 10
            .Color = wdColorAutomatic
        End With
        
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .RightIndent = InchesToPoints(0.1)
            .LeftIndent = InchesToPoints(0.1)
        End With
        
        With .Table
            
             ' Not sure what these do; want to keep rows from breaking,
             'and possibly keep tables together(?)
            .AllowPageBreaks = False
            .AllowBreakAcrossPage = False
        
            .RowStripe = 1
            .ColumnStripe = 0
            
            .LeftPadding = InchesToPoints(0)
            .RightPadding = InchesToPoints(0)
            '.Spacing = InchesToPoints(0)
    
            With .Condition(wdOddRowBanding)
                With .Shading
                    .Texture = wdTextureNone
                    .ForegroundPatternColor = wdColorAutomatic
                    .BackgroundPatternColor = -738132173
                End With
                                
                With .Borders(wdBorderLeft)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = wdColorAutomatic
                End With
                
                With .Borders(wdBorderRight)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = wdColorAutomatic
                End With
                
                With .Borders(wdBorderVertical)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth050pt
                    .Color = wdColorAutomatic
                End With
            
            End With

            'Adjust borders
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
            With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
            'Vertical borders should be included for coded comment appendices
            With .Borders(wdBorderVertical)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
            'For coded comments, need to change style of the last row to adjust
            With .Condition(wdLastRow)
                .Font.Bold = True
                .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                .Borders(wdBorderTop).LineWidth = wdLineWidth050pt
            End With
            
            With .Condition(wdLastColumn)
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
                   
          'Format Header to have bottom border
            With .Condition(wdFirstRow).Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
        End With

    End With

End Sub

Sub alphabetize_table(i As Integer)
Attribute alphabetize_table.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.alphabetize_table"
'
' alphabetize_table Macro
'From recorded macro; has not yet been tested or incorporated into macro
'

'Sort verbatim text appendices alphabetically
    With ActiveDocument
    
        Dim nTables As Long
        nTables = .Tables.Count
    
'        Dim i As Long
'        For i = 1 To nTables
'        i = 1
            nrow = .Tables(i).Rows.Count
            ncol = .Tables(i).Columns.Count
            
            If (nrow > 6) Then
                With .Tables(i)
                    Set responseRows = .Rows(7).Range
                    If ncol = 1 Then
                        responseRows.End = .Rows(nrow).Range.End
                    ElseIf ncol = 2 Then
                        responseRows.End = .Rows(nrow - 1).Range.End
                    End If
                End With
                
                responseRows.Select
                If (ncol = 1) Then
                    Selection.Sort ExcludeHeader:=False, _
                        FieldNumber:="Column 1", _
                        SortFieldType:=wdSortFieldAlphanumeric, _
                        SortOrder:=wdSortOrderAscending, _
                        LanguageID:=wdEnglishUS, subFieldNumber:="Paragraphs"
                ElseIf (ncol = 2) Then
                    Selection.Sort ExcludeHeader:=False, _
                        FieldNumber:="Column 2", _
                        SortFieldType:=wdSortFieldNumeric, _
                        SortOrder:=wdSortOrderDescending, _
                        FieldNumber2:="Column 1", _
                        SortFieldType2:=wdSortFieldAlphanumeric, _
                        SortOrder2:=wdSortOrderAscending, _
                        LanguageID:=wdEnglishUS, subFieldNumber:="Paragraphs"
                End If
            
            End If
'        Next
    End With
End Sub

Sub Appendix_Merge_Header(i As Integer)
Attribute Appendix_Merge_Header.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Appendix_Merge_Header"
'
' Appendix_Merge_Header Macro
'
'
With ActiveDocument

'Dim nTables As Long
'nTables = .Tables.Count
'Dim i As Long
'For i = 1 To nTables

ncol = .Tables(i).Columns.Count

If ncol = 2 Then
    .Tables(i).Rows(1).Select
    Selection.Cells.Merge
End If


Set mergeCells = .Tables(i).Rows(2).Range
mergeCells.End = .Tables(i).Rows(5).Range.End
mergeCells.Select
Selection.Cells.Merge

With Selection.ParagraphFormat
    .SpaceBefore = 0
    .SpaceAfter = 5
End With

.Tables(i).Rows(2).Height = 1

End With

End Sub

Sub duplicateHeaderText(i As Integer)

'The program produces coded comment tables with header text printed twice
'Before we merge the cells, we need to delete the duplicate text
'This macro will remove the text in the header rows of the second column

With ActiveDocument

'Dim nTables As Long
'nTables = .Tables.Count
'Dim i As Long
'For i = 1 To nTables

    ncol = .Tables(i).Columns.Count

'Clear text from coded comment tables; likely, this should be its own macro
    If ncol = 2 Then
        Set duplicateHead = .Tables(i).Columns(2).Cells(1).Range
        duplicateHead.End = .Tables(i).Columns(2).Cells(4).Range.End
        duplicateHead.Select
        duplicateHead.Delete
    End If

'Next

End With

End Sub


Sub appendix_table_formatting_CBB()

'Created by CB;

'Dim i As Long
'i = 1

    With ActiveDocument
    
    
    Set entireDoc = .Range
    With entireDoc
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Color = wdColorAutomatic
        .ParagraphFormat.Alignment = wdAlignLeft
    End With
    
    Dim nTables As Long
    nTables = .Tables.Count
'   MsgBox ("Number of Tables: " & nTables)
    Debug.Print nTables
    
    For i = 1 To nTables
        nrow = .Tables(i).Rows.Count
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

'This tells me the first row after the page break
'I need to expand this to find a row after EACH page break
'Which will provide the formatting I need

Sub BorderAtBreak()

Dim nTables As Long
    nTables = ActiveDocument.Tables.Count
'    MsgBox ("Number of Tables: " & nTables)
    
'    For i = 1 To nTables
For t = 1 To nTables

    Dim r As Range
    Dim tblStartPage As Long
    Dim oTable As Table
    Dim oTableRange As Range
    Dim oRow As Row
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
'        MsgBox ("Start Page: " & tblStartPage)
    '    Set refPage = tblStartPage
    ' loop through each row checking if it is the same page
    For Each oRow In oTable.Rows
       Set r = oRow.Range
       r.Collapse 1
'       MsgBox ("Investigate row " & oRow.Index)
       If r.Information(wdActiveEndPageNumber) <> _
          tblStartPage Then
'          MsgBox ("Row " & oRow.Index & " is AFTER the page break.")
          b = oRow.Index - 1
'          MsgBox (b)
          Set bottomCell = ActiveDocument.Range(Start:=oTable.Cell(b, 1).Range.Start, _
                End:=oTable.Cell(b + 1, 1).Range.End)
    
            With bottomCell.Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleSingle
                .Color = wdColorAutomatic
            End With
          
          i = i + 1
          
          tblStartPage = tblStartPage + 1
'          MsgBox ("New Table Start Page: " & tblStartPage)
       
       End If
    Next
    


Next

End Sub

Sub fix_page_breaks()

'Macro written by CBB to adjust page breaks in appendix tables
' Will need to adapt code to work with preview tables as well
' Version with EM edits

    With ActiveDocument
        Dim nTables As Long
        nTables = .Tables.Count
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        
        With Selection.Find
            .Text = "Responses: "
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With

    
    For i = 1 To nTables
    
        nrow = .Tables(i).Rows.Count
        
        'Determine page of first row in table

        .Tables(i).Rows(1).Select
        FirstRowPage = Selection.Information(wdActiveEndPageNumber)
        Debug.Print "FirstRowPage: " + Str(FirstRowPage)

        'Need to determine whether there are actual responses
        
        .Tables(i).Select
        
        If Selection.Find.Execute Then
            ResponseRow = Selection.Information(wdEndOfRangeRowNumber)
            ResponseRowPage = Selection.Information(wdActiveEndPageNumber)
            
            Debug.Print "ResponseRow: " + Str(ResponseRow)
            Debug.Print "nrow: " + Str(nrow)
            Debug.Print "ResponseRowPage: " + Str(ResponseRowPage)
            
            If ResponseRow = nrow Then
                If ResponseRowPage <> FirstRowPage Then
                    .Tables(i).Rows(1).Select
                    Selection.InsertBreak (wdPageBreak)
                End If
            
            ElseIf ResponseRow < nrow Then
                .Tables(i).Rows(ResponseRow + 1).Select
                FirstCommentPage = Selection.Information(wdActiveEndPageNumber)
                Debug.Print "FirstCommentPage: " + Str(FirstCommentPage)
                
                If FirstRowPage <> FirstCommentPage Then
                    .Tables(i).Rows(1).Select
                    Selection.InsertBreak (wdPageBreak)
                End If
            End If
        End If
                              
    Next
    
    End With

End Sub


Sub fix_page_breaks_CBB_orig()

'Macro written by CBB to adjust page breaks in appendix tables
' Will need to adapt code to work with preview tables as well

    With ActiveDocument
        Dim nTables As Long
        nTables = .Tables.Count
    
    For i = 1 To nTables
        nrow = .Tables(i).Rows.Count
        
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
                 Selection.MoveUp Unit:=wdLine, Count:=1
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

Sub Add_Extra_Borders()
    With ActiveDocument
            Dim nTables As Long
            nTables = .Tables.Count
        
        For i = 1 To nTables
            nrow = .Tables(i).Rows.Count
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

'This macro will iterate through and ensure that questions and tables are on the same page
'The macro can be run multiple times
'If running a second time, the macro will first iterate through and check for page breaks
'Existing breaks will be removed, and additional breaks added to keep everything running
'   smoothly and in order of what should be happenning.


    With ActiveDocument
    Dim nTables As Long
    nTables = .Tables.Count
        
    For i = 1 To nTables
    
        nrow = .Tables(i).Rows.Count
        ncol = .Tables(i).Columns.Count
                
        'Question tables will have 1 row, others will have multiple
        
        Debug.Print "Table " + Str(i) + "(" + Str(nrow) + "x"; Str(ncol) + ")"
        
        If ncol > 1 Then
            .Tables(i).Rows(1).Select
        
            answerRow1 = Selection.Information(wdActiveEndPageNumber)
        
            .Tables(i).Rows(nrow).Select
            answerRowN = Selection.Information(wdActiveEndPageNumber)
            
            .Tables(i - 1).Rows(1).Select
            questionRow1 = Selection.Information(wdActiveEndPageNumber)
            
            qRows = .Tables(i - 1).Rows.Count
            .Tables(i - 1).Rows(qRows).Select
            questionRowN = Selection.Information(wdActiveEndPageNumber)
            
            If .Tables(i - 1).Columns.Count <> 1 Then
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

nsection = ActiveDocument.Sections.Count
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

Sub preview_remove_block_titles()

'This macro will remove the section indicators (block titles from .qsf)
'They are currently input into the document as heading 5
'We want to delete the row of text with heading 5 and the next row

With Selection.Find
    .ClearFormatting
    .Style = ActiveDocument.Styles("Heading 5")
    .Replacement.ClearFormatting
    .Text = ""
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With


npar = ActiveDocument.Paragraphs.Count
Debug.Print (npar)
For i = 1 To npar
    Debug.Print "Paragraph" + Str(i)
    ActiveDocument.Paragraphs(i).Range.Select
    Selection.HomeKey Unit:=wdLine
    Selection.Find.Execute

    If Selection.Find.Found = True Then
        Selection.Find.Parent.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
        Selection.Find.Parent.Delete
    Else: Exit For
    End If

Next
        
End Sub

Sub TableCellPadding()

'For Lauren to run after previews have been generated
'Will adjust cell padding for all tables
'Need to add this to initial macro for others to run

With ActiveDocument
    nTables = .Tables.Count
    For i = 1 To nTables
        ncol = .Tables(i).Columns.Count
        nrow = .Tables(i).Rows.Count
        
        If ncol > 1 Then
            With .Tables(i)
                .LeftPadding = InchesToPoints(0.08)
                .RightPadding = InchesToPoints(0.08)
                .TopPadding = InchesToPoints(0.01)
                .BottomPadding = InchesToPoints(0.01)
                
                
            End With
        End If
    Next

End With

End Sub

Sub DefaultParagraphSpacing()

With ActiveDocument
    
'    .Paragraphs.LineUnitAfter = 0
'    .Paragraphs.LineUnitBefore = 0
    
    .Paragraphs.SpaceAfterAuto = False
    .Paragraphs.SpaceAfter = 0
    .Paragraphs.SpaceBeforeAuto = False
    .Paragraphs.SpaceBefore = 0
    
End With
    
End Sub

Sub remove_blockHeaders_HTML()

    With ActiveDocument
    
    Dim loopCount As Integer
    loopCount = 1
    
    
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
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Execute
    
    Do While Selection.Find.Found = True And loopCount < 1000
    
        Debug.Print iCount
        Selection.Expand wdParagraph
        Selection.Delete
        Selection.EndOf
        Selection.HomeKey Unit:=wdStory
        Selection.Find.Execute
    Loop
    
    
    
    End With

End Sub

Sub removeTotalRowCoded()

With ActiveDocument
    
    Dim loopCount As Integer
    loopCount = 1
    
    Dim nTables As Integer
    Dim nrow As Integer
        
    nTables = ActiveDocument.Tables.Count
    
    For i = 1 To nTables
        
        nrow = ActiveDocument.Tables(i).Rows.Count
    
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
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Execute
    
    Do While Selection.Find.Found = True And loopCount < 1000
    
        Debug.Print iCount
        Selection.Expand wdParagraph
        Selection.Delete
        Selection.EndOf
        Selection.HomeKey Unit:=wdStory
        Selection.Find.Execute
    Loop
    
    
    
    End With


End Sub


Sub left_right_padding_change()

    Dim i As Integer
    Dim ncol As Integer
    Dim nTables As Integer
    
    With ActiveDocument
    
    nTables = .Tables.Count
        
    For i = 1 To nTables
        ncol = .Tables(i).Columns.Count
    
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

Sub replace_newline()

    Dim wrdDoc As Document
    Set wrdDoc = ActiveDocument
    wrdDoc.Content.Select

'Replace new line character (^l) with carraige return (^p)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        'oryginal
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True

    End With

GoHere:
    Selection.Find.Execute Replace:=wdReplaceAll

    If Selection.Find.Execute = True Then
        GoTo GoHere
    End If

End Sub



Sub remove_extra_carraigeReturn()

    With ActiveDocument
    
    Dim para As Paragraph
    Dim i As Integer
    Dim rng As Range
    Dim nextPar As Range
    Dim j As Integer
    Dim loopCount As Integer
   
    nParagraphs = .Paragraphs.Count
    
    i = 1
    
    Do While i < .Paragraphs.Count
    
        Debug.Print ("Paragraph: " & i)
    
        Set rng = .Paragraphs(i).Range
        rng.Select
        Debug.Print (rng)
        
        If Selection.Text = Chr(13) Or Selection.Text = Chr(160) & Chr(13) Then
        
            Debug.Print ("True: " & i)
            If i = .Paragraphs.Count Then End
            j = i + 1
            Set nextPar = .Paragraphs(j).Range
            nextPar_Text = nextPar.Text
            loopCount = 1
            
            Do While (nextPar_Text = Chr(13) Or nextPar_Text = Chr(160) & Chr(13)) And loopCount < 10 And j <= .Paragraphs.Count
            
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


Sub format_See_Appendix(i)

    With ActiveDocument
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        
    With Selection.Find
        .Text = "See Appendix."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    If .Tables(i).Columns.Count = 1 Then
        
        .Tables(i).Select
        
        If Selection.Find.Execute Then
            Selection.Paragraphs.Indent
            Selection.InsertRowsAbove
        End If
            
    End If
    
    End With

End Sub

Sub trials()
    Dim i As Integer
    Dim nrow As Integer
    Dim ncol As Integer
    
    i = 7
    nrow = 12
    ncol = 6

    With ActiveDocument
       Debug.Print (.Tables.Count & " Tables")
       Debug.Print (.Tables(7).Rows.Count & " Rows")
       Debug.Print (.Tables(7).Columns.Count & " Cols")
       Call format_matrix_table(i, nrow, ncol)
    End With
        
    
    
End Sub

Sub RemoveEmptyParagraphs()

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .Text = "^p^$"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Italic = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineSingle
    End With
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "^p^$"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
 
End Sub

Sub Remove_Export_Tag()

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Export Tag: "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Do While Selection.Find.Execute
        Selection.Rows.Delete
    Loop
    
End Sub
