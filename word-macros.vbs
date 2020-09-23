' CHAIRMAN'S TOOLBAR macro collection
'
' (last updated: 2020-09-20)
'
' This is the macro collection for the Chairman's Toolbar
' It was created by Jan Ignatius (Nokia Bell Labs), with help from Jan Kåll (Nokia)
' These macros were used to help the chair and vice chair of the IFA working group in ETSI ISG NFV standardization org.
' The macros assume that the agenda document contains tables of submitted contributions, the exact layout is critical:
'
'               | UID | Title | Allocation | Source | Abstract| Status | Notes |
'
' NOTE: No other tables with exactly 7 columns should be used in the agenda, otherwise some of the macros will fail.
'
' The supported workflow is as follows:
' - separate excel macro was used to wrangle the .xls file from portal.etsi.org to match the above layout.
' - each contribution starts with an empty status, with white/no-color background
' - contributions can be either approved, (feat)agreed, almost feat agreed, revised, noted, postponed, withdrawn or given a temporary
'   flag of "return"
' - the TODO-tool assumes we want the agenda to only contain contributions with a status of approved, noted, withdrawn or postponed
'   and highlights all contributions that do not match those states
' - the linking tool is useful when approving agreed documents - if a link exists then the file is available.
'
' There are some leftover macros that are not visible in the toolbar, e.g. "email approval" and "further discussion required"
' The file linking macro uses the windows username to quickly populate the correct folder for contributions, for others it will ask
' the user to point it to the right directory.
'
' Known bugs:
' - the unlinking macro is buggy, haven't figured out why some of the links in the UID column are not cleared away.
' - the tabulation of the pop-up window and printed statistics is not always neatly lined up.
' - there is not nearly enough error catching in the macros, e.g. the statistics tools get really unhappy if you have other tables
'   with 7 rows that don't conform to the layout listed above.
'
' TODO:
' - combine "agreed", "almost feat agreed" and "agreed megaCR" functions - there's a lot of copy&pasting done currently between them
' - optimize the TODO-tool as it now iterates all cells, we could just concentrate on the 6th cell of a row
' - LinkFiles and todo-tool has some copy&pasted code between them, we could clean this up

Public Current_Meeting_Number As Integer
Public IFA_Document_Directory As String
Public User_Not_Ready As Boolean


Function countStatusPopup()
Dim popup As Boolean
popup = True
s = createStatistics(popup)
End Function

Function countStatus()
Dim popup As Boolean
popup = False
s = createStatistics(popup)
End Function


Function createStatistics(popup As Boolean) As Boolean
'
' countStatus Macro
'
' Creates a set of statistics that can be inserted to the agenda or displayed in a pop-up box.

  Dim tbl As Table
  Dim r As Row
  Dim s As String ' Used to check the cell content in the "Status" cell for the document
  Dim countApproved As Integer: countApproved = 0
  Dim countAgreed As Integer: countAgreed = 0
  Dim countTBRnotHandled As Integer: countTBRnotHandled = 0  ' only count "To be revised" if the revision has not yet been handled.
  Dim countNoted As Integer: countNoted = 0
  Dim countEmpty As Integer: countEmpty = 0 'Only counted if there is a document number in the first cell of the row
  Dim countPostponed As Integer: countPostponed = 0
  Dim countWithdrawn As Integer: countWithdrawn = 0
  Dim countReturn As Integer: countReturn = 0
  Dim countAgreedFeat As Integer: countAgreedFeat = 0
  Dim countAlmostAgreedFeat As Integer: countAlmostAgreedFeat = 0
  Dim countAgreedMegaCR As Integer: countAgreedMegaCR = 0
  Dim countApprovedMegaCR As Integer: countApprovedMegaCR = 0
  Dim countOthers As Integer: countOthers = 0
  Dim countStatus As Integer: countStatus = 0
  Dim orangeDoc As Integer: orangeDoc = 0 ' Denotes that the document skipped due to lack of time for one meeting cycle
  Dim redDoc As Integer: redDoc = 0 ' Denotes that the document skipped due to lack of time for two meeting cycles
  
  Dim countAlldocs As Integer: countAlldocs = 0 ' Used when showing number of all documents in the list and % handled in summary
  Dim iTbl As Integer: iTbl = 0                 ' Identifies the current table being checked
  Dim revisionStatus As String                  ' Contains the Status of the document below the row "To be revised"
  Dim firstCell As String                       ' Used when checking if there is a doc number in the first cell of the row
  Dim boxTitle As String                        ' Title of message box showing the results of this macro
  Dim outputSelect As Integer

  Dim dt As Object, utc As Date                 ' Getting UTC time to be used when printing Status summary in the document
    Set dt = CreateObject("WbemScripting.SWbemDateTime")
    dt.SetVarDate Now
    utc = dt.GetVarDate(False)
  
boxTitle = "Document list statistics"
' Lets hide the change marks and only look at the final mode.
With ActiveWindow.View
.ShowRevisionsAndComments = False
.RevisionsView = wdRevisionsViewFinal
End With


For Each tbl In ActiveDocument.Tables   ' Iterate through all tables in the document
    iTbl = iTbl + 1                     ' Keeping track which table is being checked
    If tbl.Columns.Count <> 7 Then      ' This is quite a weak check, the macro goes wrong if any other new table happens to have 7 columns.
        GoTo NextIteration              ' This is not a correct table
    End If
    
    For Each r In tbl.Rows
        firstCell = Left(r.Cells(1).Range.Text, Len(r.Cells(1).Range.Text) - 2) ' Read possible document number in the first cell
        ' The Left formula cuts off the line break and end-of-cell characters
        If r.IsFirst Then
            If firstCell = "Uid" Then GoTo NextRow  ' If we are on a contribution table, the first row contains the header and the first cell
                                                    ' of the header row should read "Uid". This means we skip this row
        End If
        If firstCell = "" Then GoTo NextRow         ' If the first cell is empty on any row, we skip that row
        
        If r.Cells(6).Shading.BackgroundPatternColor = wdColorGray15 Then 'Checks if the document is to be revised based on the cell color.
                                                                          ' We could also check the status field but that sometimes contains
                                                                          ' extra notes so the color check is more reliable.
            revisionStatus = ActiveDocument.Tables(iTbl).Cell(r.Index + 1, 6).Range.Text 'Read Status text of document below current row
            revisionStatus = Left(revisionStatus, Len(revisionStatus) - 2) 'Remove line break and end of cell characters
            If revisionStatus <> "" Then
                countAlldocs = countAlldocs + 1 'Adding also document "To be revised" to the total
                GoTo NextRow                    'The revision of this document has been handled; skip the row
            End If 'The document status "To be revised" is only counted below when there is no Status text for the revised doc
        End If

        s = LCase(r.Cells(6).Range.Text)        ' All lower case
        s = Replace(s, Chr(13), vbNullString)   ' Remove line breaks
        s = Replace(s, Chr(7), vbNullString)    ' Remove end of cell character
        countAlldocs = countAlldocs + 1         ' Unconditional counting of all docs ending up here. Only shown as % in summary.
                
        If InStr(s, "agreed megacr") > 0 Then
            countAgreedMegaCR = countAgreedMegaCR + 1
        ElseIf InStr(s, "approved megacr") > 0 Then
            countApprovedMegaCR = countApprovedMegaCR + 1
        ElseIf InStr(s, "approved") > 0 Then
            countApproved = countApproved + 1
        ElseIf InStr(s, "almost agreed feat") > 0 Then
            countAlmostAgreedFeat = countAlmostAgreedFeat + 1
        ElseIf InStr(s, "agreed feat") > 0 Then
            countAgreedFeat = countAgreedFeat + 1
        ' IMPORTANT: The ^above^ Feat checks need to be first, because the macro just looks for hits in a string. In the above two, the word "agreed" appears
        ' and would therefore count as just a normal agreed document if we don't check for "feat agreed"-hits first. Note that same problem applies also between
        ' "almost feat agreed" and "feat agreed", we need to check the former first, otherwise the latter gets the hit for both cases.
        ElseIf InStr(s, "agreed") > 0 Then
            countAgreed = countAgreed + 1
        ElseIf InStr(s, "to be revised") > 0 Then
            countTBRnotHandled = countTBRnotHandled + 1
        ElseIf InStr(s, "noted") > 0 Then
            countNoted = countNoted + 1
        ElseIf InStr(s, "postponed") > 0 Then
            countPostponed = countPostponed + 1
        ElseIf InStr(s, "withdrawn") > 0 Then
            countWithdrawn = countWithdrawn + 1
        ElseIf InStr(s, "return") > 0 Then
            countReturn = countReturn + 1
        ElseIf s = "" Then
            countEmpty = countEmpty + 1
        Else
            countOthers = countOthers + 1
        End If
        'Counting rows with red or orange color in the STATUS cell of the document row. The colors of the other cells on the same row are ignored!
        ' TODO: we may need to change this check to be against the first column as that's where the flag is "stored" between meetings.
        If r.Cells(6).Shading.BackgroundPatternColor = -721354957 Then ' light red from the default selection palette, the status cell is in column 6
            redDoc = redDoc + 1
        End If
        If r.Cells(6).Shading.BackgroundPatternColor = -654246093 Then ' light orange from the default selection palette
            orangeDoc = orangeDoc + 1
        End If
        
NextRow:
    Next r
NextIteration:
    Next tbl
  
    countStatus = countApproved + countAgreed + countNoted + countPostponed + countWithdrawn + countTBR + countAgreedFeat + countAlmostAgreedFeat + countAgreedMegaCR + countApprovedMegaCR + countOthers
    documentsHandled = countAlldocs - countTBRnotHandled - countEmpty ' Self-explanatory

    ' Added macro text below for printing Summary Status in the document or on screen.
    If popup = False Then
    ' Stop
    ' NOTE: remove "'" on the ^above^ line before "Stop" if you need to check how the macro types text in your documnent
        Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Summary status for current document list"  ' First checking if this [new] main headline is found in the document
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    
If Selection.Find.Found = True Then ' Then checking if the status summary was already printed before, delete old text if exists
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Find.ClearFormatting
    Mainheadline = True
    With Selection.Find
        .Text = "Document list statistics"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    
If Selection.Find.Found = True Then ' Deleting old statistics if already exists - TODO: we should agree the deleted text in case change marks is active?
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=18, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, Count:=1
    End If ' Printing new document status list below, including utc time stamp

    Selection.TypeText Text:=boxTitle + " at " + CStr(utc) + Chr(13) + _
          "All docs & revs:" + Chr(9) + Chr(9) + CStr(countAlldocs) + Chr(11) + _
          "Approved:" + Chr(9) + Chr(9) + CStr(countApproved) + Chr(11) + _
          "Approved MC:" + Chr(9) + Chr(9) + CStr(countApprovedMegaCR) + Chr(11) + _
          "Agreed:" + Chr(9) + Chr(9) + Chr(9) + CStr(countAgreed) + Chr(11) + _
          "Agreed MC:" + Chr(9) + Chr(9) + CStr(countAgreedMegaCR) + Chr(11) + _
          "Almost Agreed Feat:" + Chr(9) + CStr(countAlmostAgreedFeat) + Chr(11) + _
          "Agreed Feat:" + Chr(9) + Chr(9) + CStr(countAgreedFeat) + Chr(11) + _
          "Noted:" + Chr(9) + Chr(9) + Chr(9) + CStr(countNoted) + Chr(11) + _
          "Return:" + Chr(9) + Chr(9) + Chr(9) + CStr(countReturn) + Chr(11) + _
          "Postponed:" + Chr(9) + Chr(9) + CStr(countPostponed) + Chr(11) + _
          "Other status:" + Chr(9) + Chr(9) + CStr(countOthers) + Chr(11) + _
          "Withdrawn:" + Chr(9) + Chr(9) + CStr(countWithdrawn) + Chr(11) + _
          "Revisions NOT handled:" + CStr(countTBRnotHandled) + Chr(11) + _
          "All NOT handled:" + Chr(9) + CStr(countEmpty) + Chr(11) + _
          "TOTAL handled:       " + Chr(9) + CStr(documentsHandled) + " (i.e. " + CStr(Int(documentsHandled / countAlldocs * 100)) + _
          "% of all.)" + Chr(11) + _
          "Skipped for 2 meetings:" + Chr(9) + CStr(redDoc) + Chr(11) + _
          "Skipped for 1 meeting:" + Chr(9) + CStr(orangeDoc) + Chr(11)
Else ' When the new Status summary headline was NOT found in the document
    MsgBox "Header: 'Summary status for current document list' is missing"

End If ' NOTE: there should be a headline "Summary status for current document list" with 16 lines available below the headline
       ' This headline could be placed after the 'Page break after the IPR Policy box, before the 'Section brake
       ' (the 'Page break and 'Section Breaks indications gets visible in the document by ticking show/hide ['pi'] on the Home menu bar)
       
    Else ' User prefers a popup-window for the statistics
    MsgBox ("All docs & revisions:" + Chr(9) + CStr(countAlldocs) + Chr(13) + _
          "Approved:" + Chr(9) + Chr(9) + CStr(countApproved) + Chr(13) + _
          "Approved MC:" + Chr(9) + Chr(9) + CStr(countApprovedMegaCR) + Chr(13) + _
          "Agreed:" + Chr(9) + Chr(9) + Chr(9) + CStr(countAgreed) + Chr(13) + _
          "Agreed MC:" + Chr(9) + Chr(9) + CStr(countAgreedMegaCR) + Chr(13) + _
          "Almost Agreed Feat:" + Chr(9) + CStr(countAlmostAgreedFeat) + Chr(13) + _
          "Agreed Feature:" + Chr(9) + Chr(9) + CStr(countAgreedFeat) + Chr(13) + _
          "Noted:" + Chr(9) + Chr(9) + Chr(9) + CStr(countNoted) + Chr(13) + _
          "Return:" + Chr(9) + Chr(9) + Chr(9) + CStr(countReturn) + Chr(13) + _
          "Postponed:" + Chr(9) + Chr(9) + CStr(countPostponed) + Chr(13) + _
          "Other status:" + Chr(9) + Chr(9) + CStr(countOthers) + Chr(13) + _
          "Withdrawn:" + Chr(9) + Chr(9) + CStr(countWithdrawn) + Chr(13) + _
          "Revisions NOT handled:" + Chr(9) + CStr(countTBRnotHandled) + Chr(13) + _
          "All NOT handled:" + Chr(9) + Chr(9) + CStr(countEmpty) + Chr(13)) + _
          "TOTAL handled:" + Chr(9) + Chr(9) + CStr(documentsHandled) + " (i.e. " + CStr(Int(documentsHandled / countAlldocs * 100)) + _
          "% of all.)" + Chr(13) + _
          "Skipped for 2 meetings:" + Chr(9) + CStr(redDoc) + Chr(13) + _
          "Skipped for 1 meeting:" + Chr(9) + CStr(orangeDoc), , boxTitle
          ' countAlldocs counts all document rows and this includes the Docs and Revisions not yet handled.
    End If
Cancelled:

' Lets return the view back to normal, showing all markup and revisions.
With ActiveWindow.View
.ShowRevisionsAndComments = True
.RevisionsView = wdRevisionsViewMarkupAll
End With

End Function

Sub NewRevision()
'
' Appends a new row after the selected row and marking the current row's document as revised
'
'
    
    ' Make sure you are in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If
    
' lets hide the change marks and only look at the final mode. This is needed so we get the correct revision count.
With ActiveWindow.View
.ShowRevisionsAndComments = False
.RevisionsView = wdRevisionsViewFinal
End With

    RowNo = Selection.Information(wdEndOfRangeRowNumber)
     
    ' Copy contribution number and increase revision number
    s1 = Selection.Tables(1).Cell(RowNo, 1).Range.Text
    s2 = Replace(s1, Chr(13), vbNullString)   ' remove line breaks
    s3 = Left(s2, Len(s2) - 1)                ' remove rightmost character
    pos = InStr(s3, "r")
    If pos > 0 Then                           ' check if this is already a revision
        I = Right(s3, Len(s2) - pos - 1)
        If IsNumeric(I) Then
            s4 = Left(s3, pos - 1) + "r" + CStr(CInt(I) + 1)
        Else
            MsgBox "Problem with revision number"
            Exit Sub
        End If
    Else
        s4 = s3 + "r1"
    End If
   
 ' lets return the view back to normal, showing all markup and revisions.
With ActiveWindow.View
.ShowRevisionsAndComments = True
.RevisionsView = wdRevisionsViewMarkupAll
End With
   
   
    ' Time to modify the table and insert the new row.
    Selection.InsertRowsBelow 1

    Selection.Tables(1).Cell(RowNo + 1, 1).Range.Text = s4
    Selection.Tables(1).Cell(RowNo, 6).Range.Text = "To be revised"
    
    ' Copy the other cells to the new row
    Selection.Tables(1).Cell(RowNo + 1, 2).Range.Text = Selection.Tables(1).Cell(RowNo, 2).Range.Text
    Selection.Tables(1).Cell(RowNo + 1, 3).Range.Text = Selection.Tables(1).Cell(RowNo, 3).Range.Text
    Selection.Tables(1).Cell(RowNo + 1, 4).Range.Text = Selection.Tables(1).Cell(RowNo, 4).Range.Text
    Selection.Tables(1).Cell(RowNo + 1, 5).Range.Text = Selection.Tables(1).Cell(RowNo, 5).Range.Text

    Selection.Tables(1).Cell(RowNo, 6).Range.HighlightColorIndex = wdNoHighlight     ' remove any old highlight color
    Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = False                       ' remove bolding (which may have been inherited from earlier rows)
    Selection.Tables(1).Rows(RowNo).Select                                           ' Select the whole row
    Selection.Shading.BackgroundPatternColor = wdColorGray15                         ' make background gray
    
    Selection.Tables(1).Cell(RowNo + 1, 7).Select       ' Move pointer to column 7 ("Notes") of the new row
    Selection.Collapse 1                                ' drop the selection, leave the cursor in the new location
    
End Sub


Sub FurtherDiscussionRequired()
'
' Mark status as Further Discussion Required
'
'
            
    ' Make sure you are in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If

    RowNo = Selection.Information(wdEndOfRangeRowNumber)
    Selection.Tables(1).Cell(RowNo, 6).Range.Text = "Further Discussion Required"   ' replace content of the status cell with "Further Discussion Required"
    Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = False                      ' remove bolding (which may have been inherited from earlier rows)
    Selection.Tables(1).Cell(RowNo, 6).Range.HighlightColorIndex = wdYellow         ' Add highlight color green
   
    Selection.Tables(1).Cell(RowNo, 7).Select       ' Move pointer to column 7 ("status") of the new row
    Selection.Collapse 1                            ' drop the selection, leave the cursor in the new location
    
End Sub

Sub ReturnToContribution()
'
' Mark status as RETURN
'
'
            
    ' Make sure you are in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If

    RowNo = Selection.Information(wdEndOfRangeRowNumber)
    Selection.Tables(1).Cell(RowNo, 6).Range.Text = "RETURN"                    ' replace content of the status cell with "Return"
    Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = False                  ' remove bolding (which may have been inherited from earlier rows)
    Selection.Tables(1).Cell(RowNo, 6).Range.HighlightColorIndex = wdYellow     ' Add highlight color yellow
   
    Selection.Tables(1).Cell(RowNo, 7).Select       ' Move pointer to column 7 ("status") of the new row
    Selection.Collapse 1                            ' drop the selection, leave the cursor in the new location
    
End Sub

Sub ResetMeetingNumber()

    Current_Meeting_Number = 0
    
End Sub

Sub ResetIFADocumentDirectory()

    IFA_Document_Directory = ""
    
End Sub

Function FindOpenItems() As Boolean

Dim available_only As Boolean
available_only = True
s = FindToDo(available_only)

End Function

Function FindAllOpenItems() As Boolean

Dim available_only As Boolean
available_only = False
s = FindToDo(available_only)

End Function

Function FindToDo(available_only As Boolean) As Boolean

'' This functions will find all the contributions with status marked as "Agreed", "Almost Agreed FEAT", "Further Discussion Required", "RETURN" or empty.
'' We will ignore rows with a status but no UID. We will also check if the corresponding file is available in the user-defined IFA Document Directory.
'' This will help the chairman to find the documents that still need further treatment during the meeting.
'' The User_Not_Ready boolean flag is used to keep track if the user has cancelled-out from setting the IFA Document Directory.
'' We need to get out of the loop if that is the case, so as to not spam the user with the same dialog box.

Dim myTables As Table
Dim myCells As Cell
Dim myRows As Row
Dim myUID As String
Dim mySkipped As Integer
Dim s As String
Dim userId As String

'' Lets hide the revision marks so the macro can evaluate the status cells correctly.
With ActiveWindow.View
.ShowRevisionsAndComments = False
.RevisionsView = wdRevisionsViewFinal
End With

'' Let's check all tables for unfinished contributions and check if the file is available
'' We will stop at each row that has a status of Agreed, Almost Agreed FEAT, Further Discussion Required or RETURN.

mySkipped = 0

For Each myTables In ActiveDocument.Tables

    For Each myCells In myTables.Range.Cells
        
        '' Check for a cell containing a status of unfinished document on the sixth column
        '' TODO: this goes through all the cells and takes a while - we could optimize and only look at the sixth cell of a given row.
        s = LCase(myCells.Range.Text)           ' All lower case
        s = Replace(s, Chr(13), vbNullString)   ' Remove line breaks
        s = Replace(s, Chr(7), vbNullString)    ' Remove end of cell character
        If (s = "agreed" Or s = "agreed megacr" Or s = "almost agreed feat" Or s = "further discussion required" Or s = "return" Or s = "") And myCells.ColumnIndex = 6 And User_Not_Ready = False Then
            '' Lets check the row in more detail
            myCells.Row.Select
            myUID = Selection.Cells(1).Range.Text
            
            myUID = Replace(myUID, Chr(13), vbNullString)   '' Remove line breaks
            myUID = Replace(myUID, Chr(7), vbNullString)    '' Remove end of cell character
            
            If myUID <> "" Then                             '' Let's skip rows with empty UID fields
            
                '' Let's check if the file exists in the IFA Document Folder
                '' We've hardcoded the default directories that the IFA chair's use to sync the ETSI FTP-folder(s) to make
                '' things a bit easier.
                If available_only Then
                    userId = UserNameId()
                    '' MsgBox "Checkbox status is " & useDefaultDirectory & " and the current user's ID is " & userId, vbInformation
                    '' ^Above^ line is for debugging
                    If (userId = "jignatiu" And useDefaultDirectory) Then
                        IFA_Document_Directory = "C:\Users\jignatiu\OneDrive - Nokia\3_RESOURCES\IFA-FTP-2020\"
                    End If
                    '' If (userId = "jfuller" And useDefaultDirectory) Then
                    ''    IFA_Document_Directory = "C:\Users\jfuller\Desktop\ETSI_NFV\IFA\05-CONTRIBUTIONS\2018\"
                    '' End If
                    '' ^Above^ was for our previous vice chair and can be kept as a template
                    If DoesFileExist(myUID) Then
                        myCells.Row.Select  '' lets highlight the row for the user to consider
                        ''Allow the user to continue search
                        If MsgBox("UNFINISHED and AVAILABLE item has been found and highlighted. Keep searching forward?", vbYesNo) = vbNo Then
                            GoTo foundit
                        End If
                    Else
                        mySkipped = mySkipped + 1
                    End If
                Else
                    myCells.Row.Select  '' lets highlight the row for the user to consider
                        ''Allow the user to continue search
                        If MsgBox("UNFINISHED item has been found and highlighted. Keep searching forward?", vbYesNo) = vbNo Then
                            GoTo foundit
                        End If
                End If
            End If
        End If
       Next '' Cell
    Next    '' Table

If User_Not_Ready <> True And mySkipped <> 0 And available_only Then
    s = CStr(mySkipped)
    MsgBox prompt:="Search completed." + Chr(13) + Chr(13) + "(Note: I found " + s + " unfinished and UNAVAILABLE document(s). Time to re-sync?)", buttons:=vbOKOnly, Title:="Note"
End If

If User_Not_Ready <> True And mySkipped = 0 Then
    s = CStr(mySkipped)
    MsgBox prompt:="Search completed.", buttons:=vbOKOnly, Title:="Note"
End If

'' A bit hacky but this breaks the search loop and moves the cursor to the comment field.
foundit:

    Selection.Collapse 1                            '' drop the selection, move the cursor in the new location
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.Collapse 1                            '' drop the selection, move the cursor in the new location (This is needed in case the comments field has text in it).
    
'' Lets make the change marks and comments visible again.
User_Not_Ready = False

With ActiveWindow.View
.ShowRevisionsAndComments = True
.RevisionsView = wdRevisionsViewMarkupAll
End With

End Function



Sub CreateTableTemplateForImport()
'
' CreateTableTemplateForImport Macro
' This macro creates a template for a table and resizes it to nicely fit the IFA excel imported contribution lists
' You may still need to manually disable the "auto resize to fit content -flag from the table properties (under "options")
' We will also check that the orientation of the page is in Landscape, if not, we'll create new page with the correct orientation.
'
    
    '' Are we current in Portrait mode? If so, lets create a new page and set that to landscape
     If Selection.PageSetup.Orientation = wdOrientationPortrait Then
        Selection.InsertBreak Type:=wdSectionBreakNextPage
        Selection.PageSetup.Orientation = wdOrientLandscape
    End If
    
    '' Add a new 7-column table. "Word8TableBehaviour" should prevent auto-resize for content.
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=7, DefaultTableBehavior:=wdWord8TableBehavior

    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With

    '' Lets resize the columns so they can fit the expected content.
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=69.05, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=90.3, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=60.15, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=134.65, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=75.35, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(6).SetWidth ColumnWidth:=59.5, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(7).SetWidth ColumnWidth:=210.65, RulerStyle:=wdAdjustNone

    '' Now we can name each column
    Selection.Tables(1).Cell(RowNo, 1).Select
    Selection.Range.Text = "Uid"

    Selection.Tables(1).Cell(RowNo, 2).Select
    Selection.Range.Text = "Title"

    Selection.Tables(1).Cell(RowNo, 3).Select
    Selection.Range.Text = "Allocation"

    Selection.Tables(1).Cell(RowNo, 4).Select
    Selection.Range.Text = "Submitted By (Source)"

    Selection.Tables(1).Cell(RowNo, 5).Select
    Selection.Range.Text = "Abstract"

    Selection.Tables(1).Cell(RowNo, 6).Select
    Selection.Range.Text = "Status"

    Selection.Tables(1).Cell(RowNo, 7).Select
    Selection.Range.Text = "Notes"
    
    '' Drop selection and leave cursor to Notes section.
    Selection.Collapse 1


End Sub

Function GetFolder(Optional Title As String, Optional RootFolder As Variant) As String
On Error Resume Next
GetFolder = CreateObject("Shell.Application").BrowseForFolder(0, Title, 0, RootFolder).Items.Item.Path
End Function

Function DoesFileExist(myFilter As String) As Boolean

Dim myFile As String
Dim strFolder As String

If IFA_Document_Directory = "" Then
    If MsgBox("The IFA Document Folder HAS NOT been specified, would you like to set it now?", vbYesNo) = vbYes Then
        IFA_Document_Directory = GetFolder("Please navigate to your IFA DOCUMENT DIRECTORY:") + "\"
    Else
        DoesFileExist = False
        User_Not_Ready = True
        
        Exit Function
    End If
End If

'' Iterate through all the documents in the user provided document folder and look for a match to the UID
myFile = Dir(IFA_Document_Directory & "*.*")
Do While Len(myFile) > 0
    If myFile Like "*" + myFilter + "*" Then
       DoesFileExist = True
    End If
    myFile = Dir()
Loop

End Function


Function GetFilePath(myFilter As String) As String

Dim myFile As String
Dim strFolder As String

If IFA_Document_Directory = "" And User_Not_Ready = False Then
    If MsgBox("The IFA Document Folder HAS NOT been specified, would you like to set it now?", vbYesNo) = vbYes Then
        IFA_Document_Directory = GetFolder("Please navigate to your IFA DOCUMENT DIRECTORY:") + "\"
    Else
        User_Not_Ready = True
        
        Exit Function
    End If
End If

myFile = Dir(IFA_Document_Directory & "*.*")
Do While Len(myFile) > 0
    If myFile Like "*" + myFilter + "*" Then
       GetFilePath = myFile
    End If
    myFile = Dir()
Loop

End Function

Function SetAsAgreed()
'
' Query the user for the current meeting number, assume approval to happen at the next meeting
' If this macro has been run before, we'll keep using the same meeting numbers without asking the user again.
' In case of user error, the value can be reset with "ResetMeetingNumber" macro.
' If user enters "999" as value for the meeting, we will change the behaviour to better match the needs of face2face meetings.

    If Current_Meeting_Number = 0 Then          '' The global variable starts with the value 0, and if so, we need to query the user for the current meeting #
    
        Text = InputBox("What is the current IFA meeting number? For face-to-face meeting, enter 999.", "Document agreed, candidate for approval")
        
            If Len(Text) = 0 Then
                Exit Function
            Else
                If IsNumeric(Text) Then
                    Current_Meeting_Number = CInt(Text)
                Else
                    MsgBox "Please Enter Numeric meeting Number"
                    Exit Function
                End If
            End If
    End If
        
    '' Make sure you are in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Function
    End If

    '' Keep track of the row number and add the new status for column 6
    RowNo = Selection.Information(wdEndOfRangeRowNumber)
    Selection.Tables(1).Cell(RowNo, 6).Range.Text = "Agreed"
    
    Selection.Tables(1).Cell(RowNo, 7).Select
    
        If Selection.Range.Text <> "" Then              '' We have existing text in the comment field. Lets make sure we don't delete that and only append more text as needed.
            If Current_Meeting_Number = 999 Then
                '' Current agreement is not to add any more status-related text to the comment field in a f2f meeting. Below is the more verbose alternative.
                '' Selection.Range.Text = "Agreed, will be up for approval later this week." & vbCrLf & "" & vbCrLf & "" & Selection.Range.Text
            Else
                '' This is for the conference calls.
                Selection.Range.Text = "Agreed in IFA#" & Current_Meeting_Number & ". Candidate for Approval in IFA#" & Current_Meeting_Number + 1 & "." & vbCrLf & "" & vbCrLf & "" & Selection.Range.Text
            End If
        
        Else
            If Current_Meeting_Number = 999 Then        '' Comments field is empty so lets just add the needed text
                '' Current agreement is not to add any more status-related text to the comment field in a f2f meeting. Below is the more verbose alternative.
                '' Selection.Range.Text = "Agreed, will be up for approval later this week."
            Else
                '' This is for the conference calls.
                Selection.Range.Text = "Agreed in IFA#" & Current_Meeting_Number & ". Candidate for Approval in IFA#" & Current_Meeting_Number + 1 & "."
            End If
        End If
    
    
    Selection.Tables(1).Cell(RowNo, 6).Range.HighlightColorIndex = wdNoHighlight        ' remove any old highlight color
    Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = False                          ' remove bolding (which may have been inherited from earlier rows)
    Selection.Tables(1).Rows(RowNo).Select                                              ' Select the whole row
    Selection.Shading.BackgroundPatternColor = wdColorLightGreen                        ' make background light green
    
    
    Selection.Tables(1).Cell(RowNo, 7).Select       ' Move pointer to column 7 ("status") of the new row
    Selection.Collapse 1                            ' drop the selection, leave the cursor in the new location



End Function

Function SetAsApproved()

Dim s As Boolean
Dim longrgb, r, g, b As Integer
Dim bolded As Boolean

bolded = True
longrgb = 65280     '' color equals to wdColorBrightGreen
s = ModifyRow("Approved", longrgb, r, g, b, bolded)

End Function

Function SetAsApprovedMegaCR()

Dim s As Boolean
Dim longrgb, r, g, b As Integer
Dim bolded As Boolean

bolded = True
longrgb = 0
r = 189
g = 146
b = 239
s = ModifyRow("Approved MegaCR", longrgb, r, g, b, bolded)

End Function

Function SetAsNoted()

Dim s As Boolean
Dim longrgb, r, g, b As Integer
Dim bolded As Boolean

bolded = False

longrgb = 39423     '' color equals to wdColorLightOrange
s = ModifyRow("Noted", longrgb, r, g, b, bolded)

End Function

Function SetAsPostponed()

Dim s As Boolean
Dim longrgb, r, g, b As Integer
Dim bolded As Boolean

bolded = False

longrgb = 10079487  '' color equals to wdColorTan
s = ModifyRow("Postponed", longrgb, r, g, b, bolded)

End Function

Function SetAsWithdrawn()

Dim s As Boolean
Dim longrgb, r, g, b As Integer
Dim bolded As Boolean

bolded = False

longrgb = 8421504    '' color equals to wdColorGray50
s = ModifyRow("Withdrawn", longrgb, r, g, b, bolded)

End Function

Function SetAsAlmostAgreedFeat()

'
' Query the user for the current meeting number, assume approval to happen at the next meeting
' If this macro has been run before, we'll keep using the same meeting numbers without asking the user again.
' In case of user error, the value can be reset with "ResetMeetingNumber" macro.
' If user enters "999" as value for the meeting, we will change the behaviour to better match the needs of face2face meetings.

    If Current_Meeting_Number = 0 Then          '' The global variable starts with the value 0, and if so, we need to query the user for the current meeting #
    
        Text = InputBox("What is the current IFA meeting number? For face-to-face meeting, enter 999.", "Document agreed, candidate for approval")
        
            If Len(Text) = 0 Then
                Exit Function
            Else
                If IsNumeric(Text) Then
                    Current_Meeting_Number = CInt(Text)
                Else
                    MsgBox "Please Enter Numeric meeting Number"
                    Exit Function
                End If
            End If
    End If
        
    '' Make sure you are in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Function
    End If

    '' Keep track of the row number and add the new status for column 6
    RowNo = Selection.Information(wdEndOfRangeRowNumber)
    Selection.Tables(1).Cell(RowNo, 6).Range.Text = "Almost Agreed FEAT"
    
    Selection.Tables(1).Cell(RowNo, 7).Select
    
        If Selection.Range.Text <> "" Then              '' We have existing text in the comment field. Lets make sure we don't delete that and only append more text as needed.
            If Current_Meeting_Number = 999 Then
                '' Current agreement is not to add any more status-related text to the comment field in a f2f meeting. Below is the more verbose alternative.
                '' Selection.Range.Text = "Agreed, will be up for approval later this week." & vbCrLf & "" & vbCrLf & "" & Selection.Range.Text
            Else
                '' This is for the conference calls.
                Selection.Range.Text = "Almost Agreed FEAT in IFA#" & Current_Meeting_Number & ". Candidate for Agreed FEAT in IFA#" & Current_Meeting_Number + 1 & "." & vbCrLf & "" & vbCrLf & "" & Selection.Range.Text
            End If
        
        Else
            If Current_Meeting_Number = 999 Then        '' Comments field is empty so lets just add the needed text
                '' Current agreement is not to add any more status-related text to the comment field in a f2f meeting. Below is the more verbose alternative.
                '' Selection.Range.Text = "Agreed, will be up for approval later this week."
            Else
                '' This is for the conference calls.
                Selection.Range.Text = "Almost Agreed FEAT in IFA#" & Current_Meeting_Number & ". Candidate for Agreed FEAT in IFA#" & Current_Meeting_Number + 1 & "."
            End If
        End If
    
    
    Selection.Tables(1).Cell(RowNo, 6).Range.HighlightColorIndex = wdNoHighlight        ' remove any old highlight color
    Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = False                          ' remove bolding (which may have been inherited from earlier rows)
    Selection.Tables(1).Rows(RowNo).Select                                              ' Select the whole row
    Selection.Shading.BackgroundPatternColor = RGB(221, 203, 242)                       ' make background light purple
    
    
    Selection.Tables(1).Cell(RowNo, 7).Select       ' Move pointer to column 7 ("status") of the new row
    Selection.Collapse 1                            ' drop the selection, leave the cursor in the new location



End Function

Function SetAsAgreedMegaCR()

'
' Query the user for the current meeting number, assume approval to happen at the next meeting
' If this macro has been run before, we'll keep using the same meeting numbers without asking the user again.
' In case of user error, the value can be reset with "ResetMeetingNumber" macro.
' If user enters "999" as value for the meeting, we will change the behaviour to better match the needs of face2face meetings.

    If Current_Meeting_Number = 0 Then          '' The global variable starts with the value 0, and if so, we need to query the user for the current meeting #
    
        Text = InputBox("What is the current IFA meeting number? For face-to-face meeting, enter 999.", "Document agreed, candidate for approval")
        
            If Len(Text) = 0 Then
                Exit Function
            Else
                If IsNumeric(Text) Then
                    Current_Meeting_Number = CInt(Text)
                Else
                    MsgBox "Please Enter Numeric meeting Number"
                    Exit Function
                End If
            End If
    End If
        
    '' Make sure you are in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Function
    End If

    '' Keep track of the row number and add the new status for column 6
    RowNo = Selection.Information(wdEndOfRangeRowNumber)
    Selection.Tables(1).Cell(RowNo, 6).Range.Text = "Agreed MegaCR"
    
    Selection.Tables(1).Cell(RowNo, 7).Select
    
        If Selection.Range.Text <> "" Then              '' We have existing text in the comment field. Lets make sure we don't delete that and only append more text as needed.
            If Current_Meeting_Number = 999 Then
                '' Current agreement is not to add any more status-related text to the comment field in a f2f meeting. Below is the more verbose alternative.
                '' Selection.Range.Text = "Agreed, will be up for approval later this week." & vbCrLf & "" & vbCrLf & "" & Selection.Range.Text
            Else
                '' This is for the conference calls.
                Selection.Range.Text = "Agreed MegaCR in IFA#" & Current_Meeting_Number & ". Candidate for Approved MegaCR in IFA#" & Current_Meeting_Number + 1 & "." & vbCrLf & "" & vbCrLf & "" & Selection.Range.Text
            End If
        
        Else
            If Current_Meeting_Number = 999 Then        '' Comments field is empty so lets just add the needed text
                '' Current agreement is not to add any more status-related text to the comment field in a f2f meeting. Below is the more verbose alternative.
                '' Selection.Range.Text = "Agreed, will be up for approval later this week."
            Else
                '' This is for the conference calls.
                Selection.Range.Text = "Agreed MegaCR in IFA#" & Current_Meeting_Number & ". Candidate for Approved MegaCR in IFA#" & Current_Meeting_Number + 1 & "."
            End If
        End If
    
    
    Selection.Tables(1).Cell(RowNo, 6).Range.HighlightColorIndex = wdNoHighlight        ' remove any old highlight color
    Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = False                          ' remove bolding (which may have been inherited from earlier rows)
    Selection.Tables(1).Rows(RowNo).Select                                              ' Select the whole row
    Selection.Shading.BackgroundPatternColor = RGB(221, 203, 242)                       ' make background light purple
    
    
    Selection.Tables(1).Cell(RowNo, 7).Select       ' Move pointer to column 7 ("status") of the new row
    Selection.Collapse 1                            ' drop the selection, leave the cursor in the new location



End Function


Function SetAsAgreedFeat()

Dim s As Boolean
Dim longrgb, r, g, b As Integer
Dim bolded As Boolean

bolded = True

longrgb = 0
r = 189
g = 146
b = 239
s = ModifyRow("Agreed FEAT", longrgb, r, g, b, bolded)

End Function

Function SetAsEmailApproval()

Dim s As Boolean
Dim longrgb, r, g, b As Integer
Dim bolded As Boolean

bolded = True

longrgb = 16763904  ''color equals to wdColorSkyBlue
s = ModifyRow("Email Approval", longrgb, r, g, b, bolded)

End Function

Function ModifyRow(status As String, longrgb, r, g, b As Integer, bolded As Boolean) As Boolean

'' This function updated the status field and changes the table row's background color.
'' We support both 24bit long RGB and 8bit RGB colors.
'' We have to support both because there is no easy way to point to some of the colors in the
'' the default palette, exept with 24bit colors.

Dim myStatus As String
Dim myR As Integer
Dim myG As Integer
Dim myB As Integer
Dim myLongRGN As Integer

myStatus = status
myLongRGB = longrgb
myR = r
myG = g
myB = b
            
    ' Make sure you are in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Function
    End If

    RowNo = Selection.Information(wdEndOfRangeRowNumber)
    Selection.Tables(1).Cell(RowNo, 6).Range.Text = myStatus                        ' replace content of the status cell
    If bolded Then
        Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = True                   ' Bold the Status text if requested
    Else
        Selection.Tables(1).Cell(RowNo, 6).Range.Font.Bold = False                  ' Remove bolded text which may have been inherited from earlier rows for the Status text
    End If
       
    Selection.Tables(1).Cell(RowNo, 6).Range.HighlightColorIndex = wdNoHighlight    ' remove any old highlight color

    Selection.Tables(1).Rows(RowNo).Select                                          ' Select the whole row
    If myLongRGB = 0 Then
        Selection.Shading.BackgroundPatternColor = RGB(myR, myG, myB)               ' change row background color (RGB)
    Else
        Selection.Shading.BackgroundPatternColor = myLongRGB                        ' change row background color (24bit)
    End If
    
    Selection.Tables(1).Cell(RowNo, 7).Select       ' Move pointer to column 7 ("status") of the new row
    Selection.Collapse 1                            ' drop the selection, leave the cursor in the new location

End Function

Function UserNameId() As String
    UserNameId = Environ("USERNAME") 'this picks up the login id
    '' Alternative is Word.Application.UserName , which picks up the user name set in Office
End Function

Function LinkAllFiles() As String

    Dim removeAllLinks As Boolean
    Dim newFilesOnly As Boolean
    
    newFilesOnly = False
    removeAllLinks = False
    s = LinkFiles(newFilesOnly, removeAllLinks)
    
End Function

Function LinkNewFiles() As String

    Dim removeAllLinks As Boolean
    Dim newFilesOnly As Boolean
    
    newFilesOnly = True
    removeAllLinks = False
    s = LinkFiles(newFilesOnly, removeAllLinks)
    
End Function

Function RemoveFileLinks() As String

    Dim removeAllLinks As Boolean
    Dim newFilesOnly As Boolean
    
    newFilesOnly = False
    removeAllLinks = True
    s = LinkFiles(newFilesOnly, removeAllLinks)

End Function

Function LinkFiles(newFilesOnly As Boolean, removeAllLinks As Boolean) As String
   
Dim myTables As Table
Dim myCells As Cell
Dim myRows As Row
Dim myUID As String
Dim s As String
Dim userId As String
Dim myFile As String
Dim myPath As String
Dim newOnly As Boolean
Dim removeAll As Boolean

newOnly = False
removeAll = False
removeAll = removeAllLinks
newOnly = newFilesOnly

'' Lets hide the revision marks so the macro can evaluate the status cells correctly.
With ActiveWindow.View
.ShowRevisionsAndComments = False
.RevisionsView = wdRevisionsViewFinal
Application.ScreenUpdating = False
End With

'' Lets check all tables for UIDs and check if the corresponding file is available

For Each myTables In ActiveDocument.Tables

    For Each myCells In myTables.Range.Cells
        
        '' Check for a cell containing a recognized status on the sixth column
        '' This way we know for sure that we are working on a correct table
        
        s = LCase(myCells.Range.Text)           ' All lower case
        s = Replace(s, Chr(13), vbNullString)   ' Remove line breaks
        s = Replace(s, Chr(7), vbNullString)    ' Remove end of cell character
        If (s = "agreed" Or _
            s = "agreed megacr" Or _
            s = "almost agreed feat" Or _
            s = "agreed feat" Or _
            s = "further discussion required" Or _
            s = "return" Or _
            s = "approved" Or _
            s = "approved megacr" Or _
            s = "noted" Or _
            s = "withdrawn" Or _
            s = "postponed" Or _
            s = "") And myCells.ColumnIndex = 6 And User_Not_Ready = False Then
            
            '' Lets check the row in more detail
            myCells.Row.Select
            myUID = Selection.Cells(1).Range.Text
            
            myUID = Replace(myUID, Chr(13), vbNullString)   '' Remove line breaks
            myUID = Replace(myUID, Chr(7), vbNullString)    '' Remove end of cell character
            
            If myUID <> "" Then                             '' Lets skip rows with empty UID fields
            
                '' Lets check if the file exists in the IFA Document Folder
                '' We've hardcoded the default directories that the IFA chair's use to sync the ETSI FTP-folder(s) to make
                '' things a bit easier.
                
                If newOnly And Selection.Range.Hyperlinks.Count = 0 Then   '' Lets only add the link to UIDs missing a link
                    userId = UserNameId()
                    '' MsgBox "Checkbox status is " & useDefaultDirectory & " and the current user's ID is " & userId, vbInformation
                    If (userId = "jignatiu" And useDefaultDirectory) Then
                        IFA_Document_Directory = "C:\Users\jignatiu\OneDrive - Nokia\3_RESOURCES\IFA-FTP-2019\"
                    End If
                    '' If (userId = "jfuller" And useDefaultDirectory) Then
                    ''     IFA_Document_Directory = "C:\Users\jfuller\Desktop\ETSI_NFV\IFA\05-CONTRIBUTIONS\2018\"
                    '' End If
                    If DoesFileExist(myUID) And User_Not_Ready = False Then
                        myFile = GetFilePath(myUID)
                        myPath = IFA_Document_Directory + myFile
                        '' DEBUG MsgBox myPath
                        ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=myPath, ScreenTip:=myFile
                    End If
                End If
                    
                
                If newOnly = False And removeAllLinks = False Then           '' Lets replace all the links with new ones
                    If Selection.Range.Hyperlinks.Count > 0 Then
                        Selection.Range.Hyperlinks(1).Delete '' remove old link
                    End If
                
                    myCells.Row.Select
                    userId = UserNameId()
                    '' MsgBox "Checkbox status is " & useDefaultDirectory & " and the current user's ID is " & userId, vbInformation
                    
                    If (userId = "jignatiu" And useDefaultDirectory) Then
                        IFA_Document_Directory = "C:\Users\jignatiu\OneDrive - Nokia\3_RESOURCES\IFA-FTP-2019\"
                    End If
                    
                    '' If (userId = "jfuller" And useDefaultDirectory) Then
                    ''     IFA_Document_Directory = "C:\Users\jfuller\Desktop\ETSI_NFV\IFA\05-CONTRIBUTIONS\2018\"
                    '' End If
                    
                    If DoesFileExist(myUID) And User_Not_Ready = False Then
                        myFile = GetFilePath(myUID)
                        myPath = IFA_Document_Directory + myFile
                        '' DEBUG MsgBox myPath
                        ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=myPath, ScreenTip:=myFile
                    End If
                End If
                
                If newOnly = False And removeAllLinks = True Then
                    If Selection.Range.Hyperlinks.Count > 0 Then
                        Selection.Range.Hyperlinks(1).Delete '' remove old link
                    End If
                End If
                
            End If
        End If
       Next '' Cell
    Next    '' Table

'' Lets make the change marks and comments visible again.
User_Not_Ready = False

With ActiveWindow.View
.ShowRevisionsAndComments = True
.RevisionsView = wdRevisionsViewMarkupAll
Application.ScreenUpdating = True
End With

End Function

Sub SetTrackColorRed()
'
' ChangeTrackColor Macro
' When we havea contribution with lots of different contributors, the text is difficult to read with
' change marks on. This makes all the changed text appear in red only.
'
    With Options
        .InsertedTextMark = wdInsertedTextMarkUnderline
        .InsertedTextColor = wdRed
        .DeletedTextMark = wdDeletedTextMarkStrikeThrough
        .DeletedTextColor = wdRed
        .RevisedPropertiesMark = wdRevisedPropertiesMarkNone
        .RevisedPropertiesColor = wdAuto
        .RevisedLinesMark = wdRevisedLinesMarkLeftBorder
        .CommentsColor = wdRed
        .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve
    End With
    ActiveWindow.View.RevisionsMode = wdInLineRevisions
    With Options
        .MoveFromTextMark = wdMoveFromTextMarkDoubleStrikeThrough
        .MoveFromTextColor = wdGreen
        .MoveToTextMark = wdMoveToTextMarkDoubleUnderline
        .MoveToTextColor = wdGreen
        .InsertedCellColor = wdCellColorNoHighlight
        .MergedCellColor = wdCellColorLightYellow
        .DeletedCellColor = wdCellColorPink
        .SplitCellColor = wdCellColorLightOrange
    End With
    With ActiveDocument
        .TrackMoves = True
        .TrackFormatting = True
    End With
End Sub

Sub SetTrackColorByAuthor()
'
' TrackColorByAuthor Macro
'
'
    With Options
        .InsertedTextMark = wdInsertedTextMarkUnderline
        .InsertedTextColor = wdByAuthor
        .DeletedTextMark = wdDeletedTextMarkStrikeThrough
        .DeletedTextColor = wdByAuthor
        .RevisedPropertiesMark = wdRevisedPropertiesMarkNone
        .RevisedPropertiesColor = wdAuto
        .RevisedLinesMark = wdRevisedLinesMarkLeftBorder
        .CommentsColor = wdByAuthor
        .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve
    End With
    ActiveWindow.View.RevisionsMode = wdInLineRevisions
    With Options
        .MoveFromTextMark = wdMoveFromTextMarkDoubleStrikeThrough
        .MoveFromTextColor = wdGreen
        .MoveToTextMark = wdMoveToTextMarkDoubleUnderline
        .MoveToTextColor = wdGreen
        .InsertedCellColor = wdCellColorNoHighlight
        .MergedCellColor = wdCellColorLightYellow
        .DeletedCellColor = wdCellColorPink
        .SplitCellColor = wdCellColorLightOrange
    End With
'    With ActiveDocument
'        .TrackMoves = True
'        .TrackFormatting = True
'    End With
End Sub
