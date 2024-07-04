Attribute VB_Name = "Module1"
' Version 2.0 - Works if sheet not found
' Written by Zac Spitzer, July 19 2023
' Last updated April 1 2024  ZS

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
Dim rw As Long, Cell As Range
Dim lastrow As Long
Worksheets("PRODUCTION SCHEDULE").Activate
Range("G:G").EntireColumn.Hidden = False
For Each Cell In Range("F:F")
rw = Cell.Row
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
Dim rw As Long, Cell As Range
Worksheets("GE AND OPS SCHEDULE").Activate
Range("G:G").EntireColumn.Hidden = False
For Each Cell In Range("F:F")
rw = Cell.Row
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
Dim rw As Long, Cell As Range
Worksheets("PROGRAMMING SCHEDULE").Activate
   Range("G:G").EntireColumn.Hidden = False
For Each Cell In Range("F:F")
rw = Cell.Row
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
Dim rw As Long, Cell As Range
Dim lastrow As Long
Worksheets("Extra Schedule 1").Activate
Range("G:G").EntireColumn.Hidden = False
For Each Cell In Range("F:F")
rw = Cell.Row
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
Dim rw As Long, Cell As Range
Dim lastrow As Long
Worksheets("Extra Schedule 2").Activate
Range("G:G").EntireColumn.Hidden = False
For Each Cell In Range("F:F")
rw = Cell.Row
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
Dim rw As Long, Cell As Range
Dim lastrow As Long
Worksheets("Extra Schedule 3").Activate
Range("G:G").EntireColumn.Hidden = False
For Each Cell In Range("F:F")
rw = Cell.Row
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
Worksheets("BackEnd").Activate
Range("A1:G500").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo, Key2:=Range("B1"), Order2:=xlAscending
End Sub

Sub CopyScheduleToFront()
Worksheets("BackEnd").Activate
Range("B1:G150").Copy
Worksheets("EVENT OVERVIEW").Activate
Range("A24").PasteSpecial xlPasteValuesAndNumberFormats
Range("E:F").EntireColumn.Hidden = True
Range("A22").Select
Application.CutCopyMode = False

End Sub

Sub DeleteOldSchedule()
Worksheets("EVENT OVERVIEW").Activate
Range("E:F").EntireColumn.Hidden = False
Range("A24:G150").ClearContents
End Sub

Sub ClearBackEnd()
Worksheets("BackEnd").Activate
Range("A1:G150").ClearContents
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






