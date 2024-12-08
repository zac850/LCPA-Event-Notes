Attribute VB_Name = "Module1"
' Version 3.0 - No longer iterates through entire sheets, handles > 150 lines
' Written by Zac Spitzer, July 19 2023
' Last updated December 6 2024  ZS

Sub UpdateAllSimpler()
Application.ScreenUpdating = False
ErrorCheck
DeleteOldSchedule
ClearBackEnd
If SheetExists("PRODUCTION SCHEDULE") Then
    copyProduction
    End If
If SheetExists("GE AND OPS SCHEDULE") Then
    copyGE
    End If
If SheetExists("PROGRAMMING SCHEDULE") Then
    copyProgramming
    End If
If SheetExists("Extra Schedule 1") Then
    copyExtra1
    End If
If SheetExists("Extra Schedule 2") Then
    copyExtra2
    End If
If SheetExists("Extra Schedule 3") Then
    copyExtra3
    End If
SortBackEnd
DayBreaks
CopyScheduleToFront
Application.ScreenUpdating = True
End Sub

Sub copyProduction()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim scanlastrow As Long
Dim RangeEnd As Long
Dim StartF As Range
Dim EndF As Range
Dim ScanRange As Range

Worksheets("PRODUCTION SCHEDULE").Activate

scanlastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("PRODUCTION SCHEDULE")
Set StartF = Wksht.Range("F9")
Set EndF = Wksht.Range("F" & scanlastrow)
Set ScanRange = Wksht.Range(StartF, EndF)

Range("G:G").EntireColumn.Hidden = False
For Each Cell In ScanRange
 If Cell.Value = "yes" Or _
    Cell.Value = "y" Or _
    Cell.Value = "YES" Or _
    Cell.Value = "Yes" Or _
    Cell.Value = "Y" Or _
    Cell.Value = "True" Then
    Cell.Select
    Range("G" & (ActiveCell.Row)).Value = "Production"
  Cell.EntireRow.Copy
         Sheets("BackEnd").Activate
         lastrow = Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row
         Range("A" & lastrow).PasteSpecial xlPasteValuesAndNumberFormats
         Worksheets("PRODUCTION SCHEDULE").Activate
 End If
Next
   Range("G:G").EntireColumn.Hidden = True
End Sub

Sub copyGE()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim scanlastrow As Long
Dim RangeEnd As Long
Dim StartF As Range
Dim EndF As Range
Dim ScanRange As Range

Worksheets("GE AND OPS SCHEDULE").Activate

scanlastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("GE AND OPS SCHEDULE")
Set StartF = Wksht.Range("F9")
Set EndF = Wksht.Range("F" & scanlastrow)
Set ScanRange = Wksht.Range(StartF, EndF)

Range("G:G").EntireColumn.Hidden = False
For Each Cell In ScanRange
 If Cell.Value = "yes" Or _
    Cell.Value = "y" Or _
    Cell.Value = "YES" Or _
    Cell.Value = "Yes" Or _
    Cell.Value = "Y" Or _
    Cell.Value = "True" Then
    Cell.Select
    Range("G" & (ActiveCell.Row)).Value = "GE OPS"
  Cell.EntireRow.Copy
         Sheets("BackEnd").Activate
         lastrow = Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row
         Range("A" & lastrow).PasteSpecial xlPasteValuesAndNumberFormats
         Worksheets("GE AND OPS SCHEDULE").Activate
 End If
Next
   Range("G:G").EntireColumn.Hidden = True
End Sub

Sub copyProgramming()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim scanlastrow As Long
Dim RangeEnd As Long
Dim StartF As Range
Dim EndF As Range
Dim ScanRange As Range

Worksheets("PROGRAMMING SCHEDULE").Activate

scanlastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("PROGRAMMING SCHEDULE")
Set StartF = Wksht.Range("F9")
Set EndF = Wksht.Range("F" & scanlastrow)
Set ScanRange = Wksht.Range(StartF, EndF)

Range("G:G").EntireColumn.Hidden = False
For Each Cell In ScanRange
 If Cell.Value = "yes" Or _
    Cell.Value = "y" Or _
    Cell.Value = "YES" Or _
    Cell.Value = "Yes" Or _
    Cell.Value = "Y" Or _
    Cell.Value = "True" Then
    Cell.Select
    Range("G" & (ActiveCell.Row)).Value = "Programming"
  Cell.EntireRow.Copy
         Sheets("BackEnd").Activate
         lastrow = Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row
         Range("A" & lastrow).PasteSpecial xlPasteValuesAndNumberFormats
         Worksheets("PROGRAMMING SCHEDULE").Activate
 End If
Next
   Range("G:G").EntireColumn.Hidden = True
End Sub

Sub copyExtra1()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim scanlastrow As Long
Dim RangeEnd As Long
Dim StartF As Range
Dim EndF As Range
Dim ScanRange As Range

Worksheets("Extra Schedule 1").Activate

scanlastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("Extra Schedule 1")
Set StartF = Wksht.Range("F9")
Set EndF = Wksht.Range("F" & scanlastrow)
Set ScanRange = Wksht.Range(StartF, EndF)

Range("G:G").EntireColumn.Hidden = False
For Each Cell In ScanRange
 If Cell.Value = "yes" Or _
    Cell.Value = "y" Or _
    Cell.Value = "YES" Or _
    Cell.Value = "Yes" Or _
    Cell.Value = "Y" Or _
    Cell.Value = "True" Then
    Cell.Select
    Range("G" & (ActiveCell.Row)).Value = "Extra1"
  Cell.EntireRow.Copy
         Sheets("BackEnd").Activate
         lastrow = Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row
         Range("A" & lastrow).PasteSpecial xlPasteValuesAndNumberFormats
         Worksheets("Extra Schedule 1").Activate
 End If
Next
   Range("G:G").EntireColumn.Hidden = True
End Sub

Sub copyExtra2()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim scanlastrow As Long
Dim RangeEnd As Long
Dim StartF As Range
Dim EndF As Range
Dim ScanRange As Range

Worksheets("Extra Schedule 2").Activate

scanlastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("Extra Schedule 2")
Set StartF = Wksht.Range("F9")
Set EndF = Wksht.Range("F" & scanlastrow)
Set ScanRange = Wksht.Range(StartF, EndF)

Range("G:G").EntireColumn.Hidden = False
For Each Cell In ScanRange
 If Cell.Value = "yes" Or _
    Cell.Value = "y" Or _
    Cell.Value = "YES" Or _
    Cell.Value = "Yes" Or _
    Cell.Value = "Y" Or _
    Cell.Value = "True" Then
    Cell.Select
    Range("G" & (ActiveCell.Row)).Value = "Extra2"
  Cell.EntireRow.Copy
         Sheets("BackEnd").Activate
         lastrow = Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row
         Range("A" & lastrow).PasteSpecial xlPasteValuesAndNumberFormats
         Worksheets("Extra Schedule 2").Activate
 End If
Next
   Range("G:G").EntireColumn.Hidden = True
End Sub

Sub copyExtra3()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim scanlastrow As Long
Dim RangeEnd As Long
Dim StartF As Range
Dim EndF As Range
Dim ScanRange As Range

Worksheets("Extra Schedule 3").Activate

scanlastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("Extra Schedule 3")
Set StartF = Wksht.Range("F9")
Set EndF = Wksht.Range("F" & scanlastrow)
Set ScanRange = Wksht.Range(StartF, EndF)

Range("G:G").EntireColumn.Hidden = False
For Each Cell In ScanRange
 If Cell.Value = "yes" Or _
    Cell.Value = "y" Or _
    Cell.Value = "YES" Or _
    Cell.Value = "Yes" Or _
    Cell.Value = "Y" Or _
    Cell.Value = "True" Then
    Cell.Select
    Range("G" & (ActiveCell.Row)).Value = "Extra3"
  Cell.EntireRow.Copy
         Sheets("BackEnd").Activate
         lastrow = Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row
         Range("A" & lastrow).PasteSpecial xlPasteValuesAndNumberFormats
         Worksheets("Extra Schedule 3").Activate
 End If
Next
   Range("G:G").EntireColumn.Hidden = True
End Sub

Sub SortBackEnd()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim RangeEnd As Long
Dim StartSort As Range
Dim EndSort As Range
Dim SortRange As Range

Worksheets("Backend").Activate

lastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("Backend")
Set StartSort = Wksht.Range("A1")
Set EndSort = Wksht.Range("G" & lastrow)
Set SortRange = Wksht.Range(StartSort, EndSort)

SortRange.Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo, Key2:=Range("B1"), Order2:=xlAscending
End Sub

Sub CopyScheduleToFront()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim RangeEnd As Long
Dim StartCopy As Range
Dim EndCopy As Range
Dim CopyRange As Range

Worksheets("BackEnd").Activate

lastrow = Cells(Rows.Count, "F").End(xlUp).Row
Set Wksht = Sheets("Backend")
Set StartCopy = Wksht.Range("B1")
Set EndCopy = Wksht.Range("G" & lastrow)
Set CopyRange = Wksht.Range(StartCopy, EndCopy)

CopyRange.Copy
Worksheets("EVENT OVERVIEW").Activate
Range("A24").PasteSpecial xlPasteValuesAndNumberFormats
Range("E:F").EntireColumn.Hidden = True
Range("A22").Select
Application.CutCopyMode = False

End Sub

Sub DeleteOldSchedule()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim RangeEnd As Long
Dim BeginRow As Range
Dim EndRow As Range
Dim DeleteRange As Range

Worksheets("EVENT OVERVIEW").Activate
Range("E:F").EntireColumn.Hidden = False

lastrow = WorksheetFunction.Max(Cells(Rows.Count, "E").End(xlUp).Row, 24)
Set Wksht = Sheets("EVENT OVERVIEW")
Set BeginRow = Wksht.Range("A24")
Set EndRow = Wksht.Range("G" & lastrow)
Set DeleteRange = Wksht.Range(BeginRow, EndRow)

DeleteRange.ClearContents

End Sub

Sub ClearBackEnd()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim RangeEnd As Long
Dim BeginRow As Range
Dim EndRow As Range
Dim DeleteRange As Range

Worksheets("Backend").Activate

lastrow = Cells(Rows.Count, "F").End(xlUp).Row

Set Wksht = Sheets("Backend")
Set BeginRow = Wksht.Range("A1")
Set EndRow = Wksht.Range("G" & lastrow)
Set DeleteRange = Wksht.Range(BeginRow, EndRow)

DeleteRange.ClearContents
End Sub

Sub DayBreaks()
    Dim i As Long
    i = 1
    Do While Not IsEmpty(Worksheets("BackEnd").Range("A" & i))
        If Worksheets("BackEnd").Range("A" & i + 1) <> Worksheets("BackEnd").Range("A" & i) Then
            Rows(i + 1).Insert
            Range("D" & (i + 1)).Value(xlRangeValueXMLSpreadsheet) = Range("A" & (i + 2)).Value(xlRangeValueXMLSpreadsheet)
            Range("D" & (i + 1)).NumberFormat = "dddd mmmm d, yyyy"
            i = i + 1
        End If
        i = i + 1
    Loop
If IsEmpty(Range("A1")) = False Then
   Range("A1").EntireRow.Insert
   Range("D1").Value(xlRangeValueXMLSpreadsheet) = Range("A2").Value(xlRangeValueXMLSpreadsheet)
   Range("D1").NumberFormat = "dddd mmmm d, yyyy"
End If
End Sub

Sub ErrorCheck()  'Error check if sheets got renamed!
    If SheetExists("Event Overview") Then
        Else
            MsgBox "Event Overview tab does not exist or was renamed. The macro cannot run without this tab. Add/Rename sheet to 'EVENT OVERVIEW', or call Zac."
        End
        End If
    If SheetExists("BACKEND") Then
        Else
            MsgBox "Backend tab does not exist. This tab should be hidden. The macro cannot run without this tab. Add/Rename sheet to 'Backend', or call Zac."
        End
        End If
    If SheetExists("Production Schedule") Then
        Else
            MsgBox "Production Schedule tab could not be found, so was not included. If it was intentionally deleted, carry on. If not, rename sheet to 'PRODUCTION SCHEDULE', or call Zac."
        End If
    If SheetExists("GE AND OPS SCHEDULE") Then
        Else
            MsgBox "GE and Ops tab could not be found, so was not included. If it was intentionally deleted, carry on. If not, rename sheet to 'GE AND OPS SCHEDULE', or call Zac."
        End If
    If SheetExists("PROGRAMMING SCHEDULE") Then
        Else
            MsgBox "Programming Schedule tab could not be found, so was not included. If it was intentionally deleted, carry on. If not, rename sheet to 'PROGRAMMING SCHEDULE', or call Zac."
            End If

    If SheetExists("Extra Schedule 1") Then
            Else
                MsgBox "Extra Schedule 1 tab could not be found, so was not included. If it was intentionally deleted, carry on. If not, rename sheet to 'Extra Schedule 1', or call Zac."
                End If
    If SheetExists("Extra Schedule 2") Then
            Else
                MsgBox "Extra Schedule 2 tab could not be found, so was not included. If it was intentionally deleted, carry on. If not, rename sheet to 'Extra Schedule 2', or call Zac."
                End If
    If SheetExists("Extra Schedule 3") Then
            Else
                MsgBox "Extra Schedule 3 tab could not be found, so was not included. If it was intentionally deleted, carry on. If not, rename sheet to 'Extra Schedule 3', or call Zac."
                End If

End Sub

Function SheetExists(sheetName As String, Optional Wb As Workbook) As Boolean 'Function for error check sub
    If Wb Is Nothing Then Set Wb = ThisWorkbook
    On Error Resume Next
    SheetExists = (LCase(Wb.Sheets(sheetName).Name) = LCase(sheetName))
    On Error GoTo 0
Exit Function
End Function








