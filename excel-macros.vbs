
'' ###################### Create_full_table ######################

Sub Create_full_table(control As IRibbonControl)
' Sub IFA_contribs_delete_columns_width_wrap_v1()
' Takes ETSI "contributions" Excel file and:
' 1) removes unwanted columns
' 2) adjusts column widths
' 3) wraps text
' 4) applies auto height to rows
' 5) fixes contributor details
' 6) adds Status and Notes columns
' 7) creates borders
'
' NOTE: the macro assumes that the sheet is called "Contributions"

Dim x, y As Range, z As Integer
'' set an array of columns to remove. pk_contribution, technical_body, meeting, status_comment, for (11), decision_requested, contact, file_url
'' old x = Array(2, 3, 7, 10, 11, 12, 14, 18)
x = Array(2, 3, 7, 10, 12, 14, 18)

Set y = Columns(x(0))
For z = 1 To UBound(x)
    Set y = Union(y, Columns(x(z)))
Next z
y.Select
Selection.Delete Shift:=xlToLeft

Columns("A:A").ColumnWidth = 2
Columns("B:B").ColumnWidth = 10.5
Columns("C:C").ColumnWidth = 6
Columns("D:D").ColumnWidth = 20
Columns("E:E").ColumnWidth = 9
Columns("F:F").ColumnWidth = 8
Columns("G:G").ColumnWidth = 11
Columns("H:H").ColumnWidth = 10
Columns("I:I").ColumnWidth = 31
Columns("J:J").ColumnWidth = 11
Columns("K:K").ColumnWidth = 11

    
    Cells.Select
    Worksheets("Contributions").Activate
    ' Worksheets("Sheet2").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlLTR
        .MergeCells = False
    End With
    Selection.Rows.AutoFit
   
'' New part of the macro continues from here onwards (12.4.2018)

' Sort into ascending doc number order

Range("B1", Range("K1").End(xlDown)).Sort key1:=Range("B1"), order1:=xlAscending, Header:=xlYes

' Sort into ascending agenda number order

Range("B1", Range("K1").End(xlDown)).Sort key1:=Range("E1"), order1:=xlAscending, Header:=xlYes

' Merge contributor and company in one column

    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "=""SubmittedBy""&CHAR(10)&""(Source)"""
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=RC[4]&CHAR(10)&CHAR(10)&""(""&RC[1]&"")"""
    lastRow = Cells.Find("*", [A1], , , xlByRows, xlPrevious).Row
    Selection.AutoFill Destination:=Range("H2:H" & lastRow)
    
' Note: the linebreak code above differs between platforms.
' For Windows use CHAR(10)
' For Mac use CHAR(13)
    
    
' Note that all columns from H onwards have been shifted by one
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = True

'Marks Revised documents in Grey (EntireRow) and Reserved document in Light Blue

Dim myLastRow As Long

myLastRow = Worksheets("Contributions").Range("I1").End(xlDown).Row

    For i = 2 To myLastRow
          If Worksheets("Contributions").Cells(i, 6).Value = "Revised" Then
                Cells(i, 2).EntireRow.Interior.ColorIndex = 16
          ElseIf Worksheets("Contributions").Cells(i, 6).Value = "Reserved" Then
                Cells(i, 2).Interior.ColorIndex = 8
          End If
    Next i

    Range("M1").Value = "Status"
    Range("N1").Value = "Notes"

    Set r = Range("A1:N1").CurrentRegion
    With r.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    Columns("M:M").ColumnWidth = 11
    Columns("N:N").ColumnWidth = 31

'' Added 26.04.2019 - Highlight contributions marked for information or discussion in the notes field
Dim a As Variant
Dim NumRows As Long

NumRows = Range("G1", Range("G1").End(xlDown)).Rows.Count
For x = 2 To NumRows
    a = Cells(x, 7).Value
    If a = "Information" Then
        Cells(x, 14).Value = "For Information"
    End If
    If a = "Discussion" Then
        Cells(x, 14).Value = "For Discussion"
    End If
    
Next x

End Sub

'' ###################### F2F_format ######################
'
Sub F2F_format(control As IRibbonControl)
'
' Online_IFA_meeting_1 Macro
' 1) Formats cells widths and hides unwanted cells,
' 2) Sorts into ascending document number and agenda order,
' 3) Merges company and contributor information
' 4) Greys out "Revised" contribution rows
' 5) Marks "Reserved" contribution number in Light Blue
'
' NOTE: the macro assumes that the sheet is called "Contributions"

' Initial formatting

Dim x, y As Range, z As Integer
x = Array(2, 3, 7, 10, 11, 12, 14, 18)
Set y = Columns(x(0))
For z = 1 To UBound(x)
    Set y = Union(y, Columns(x(z)))
Next z
y.Select
Selection.Delete Shift:=xlToLeft

Columns("A:A").ColumnWidth = 2
Columns("B:B").ColumnWidth = 10.5
Columns("C:C").ColumnWidth = 6
Columns("D:D").ColumnWidth = 20
Columns("E:E").ColumnWidth = 9
Columns("F:F").ColumnWidth = 8
Columns("G:G").ColumnWidth = 11
Columns("H:H").ColumnWidth = 10
Columns("I:I").ColumnWidth = 31
Columns("J:J").ColumnWidth = 11
    
    Cells.Select
    Worksheets("Contributions").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlLTR
        .MergeCells = False
    End With
    Selection.Rows.AutoFit

' Sort into ascending doc number order

Range("B1", Range("J1").End(xlDown)).Sort key1:=Range("B1"), order1:=xlAscending, Header:=xlYes

' Sort into ascending agenda number order

Range("B1", Range("J1").End(xlDown)).Sort key1:=Range("E1"), order1:=xlAscending, Header:=xlYes


' Merge contributor and company in one column

   Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "=""SubmittedBy""&CHAR(10)&""(Source)"""
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=RC[3]&CHAR(10)&CHAR(10)&""(""&RC[-1]&"")"""
    Selection.AutoFill Destination:=Range("H2:H1000")
    
' Note: the linebreak code above differs between platforms.
' For Windows use CHAR(10)
' For Mac use CHAR(13)
    
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    
'Marks Revised documents in Grey (EntireRow) and Reserved document in Light Blue

Dim myLastRow As Long

myLastRow = Worksheets("Contributions").Range("I1").End(xlDown).Row

    For i = 2 To myLastRow
          If Worksheets("Contributions").Cells(i, 6).Value = "Revised" Then
                Cells(i, 2).EntireRow.Interior.ColorIndex = 16
          ElseIf Worksheets("Contributions").Cells(i, 6).Value = "Reserved" Then
                Cells(i, 2).Interior.ColorIndex = 8
          End If
    Next i
    
End Sub

'' ###################### HighlightDiffsMultipleSheets ######################
'
' Source: https://answers.microsoft.com/en-us/mac/forum/macoffice2011-macexcel/macro-to-highlight-differences-between-worksheets/4b3134b7-9d2b-42da-b051-e8caed725ded
' Only highlights text (string) changes, but this is enough for us as the raw import from ETSI uses "General" category for all cell types.
'
' NOTE: the macro assumes that the sheets that are being compared are called "Sheet1" and "Sheet2"

 Public Sub HighlightDiffsMultipleSheets(control As IRibbonControl)
        Const csFormulaBase As String = "COUNTIF('*S0*'!*R1*, *R2*)"
        Const csJoin As String = " + "
        Dim i As Long
        Dim sFormula As String
        Dim vSheets As Variant
        
        vSheets = Array("Sheet1", "Sheet2")
        sFormula = vbNullString
        
        For i = LBound(vSheets) To UBound(vSheets)
            With Worksheets(vSheets(i)).UsedRange
                sFormula = sFormula & csJoin & Replace( _
                    Replace(Replace(csFormulaBase, "*R2*", _
                        .Cells(1).Address(False, False)), _
                        "*R1*", .Cells.Address), "*S0*", .Parent.Name)
            End With
        Next i
        sFormula = "=" & Mid(sFormula, Len(csJoin) + 1) & " = 1"
                        
        For i = LBound(vSheets) To UBound(vSheets)
            With Worksheets(vSheets(i)).UsedRange.Cells.FormatConditions
                .Delete
                .Add Type:=xlExpression, Formula1:=sFormula
                With .Item(1)
                    .SetFirstPriority
                    With .Interior
                        .PatternColorIndex = 35
                        .ColorIndex = 35
                    End With
                    .StopIfTrue = False
                End With
            End With
        Next i
    End Sub

