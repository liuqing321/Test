Attribute VB_Name = "Module11"
Sub PasteFilteredResult()

'Step1 - Copy and paste the filtered result to a new sheet.
'The new worksheet will be used for adding employee name and division in summary report

' add new sheet

Sheets.Add.Name = "result"

'filter the report
Worksheets("Page 1").Range("A1").AutoFilter Field:=7, Criteria1:=Array("Awaiting User Info", "Open"), Operator:=xlFilterValues


'copy the filtered result and paste it to the result page

Worksheets("Page 1").Range("W:W").Copy


Worksheets("result").Range("A1").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


Worksheets("Page 1").Range("A:V").Copy


Worksheets("result").Range("B1").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False




End Sub

Sub PasteCleanData()

' Step2 - After cleaning the data in CCAF report and cc profiles, copy and paste the cc#s, CostCenterName,TargetRange,BasicType,LOB and Operation(Column W to Columns AB) columns to summary report.


'filter the report
Worksheets("Page 1").Range("A1").AutoFilter Field:=7, Criteria1:=Array("Awaiting User Info", "Open"), Operator:=xlFilterValues

'copy the filtered result and paste it to the result page



Worksheets("Page 1").Range("H:H").Copy
    
Worksheets("demo").Range("A2").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


Worksheets("Page 1").Range("W:AB").Copy
    
Worksheets("demo").Range("B2").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False




End Sub
Sub VlookupFromCCprofiles()

'Step 3 - Compare the CCAF report to CC Profiles by using vlookup function

On Error Resume Next
For i = 2 To Worksheets("demo").UsedRange.Rows.Count


Worksheets("demo").Cells(i, 8).Value = Application.WorksheetFunction.VLookup( _
Worksheets("demo").Cells(i, 2).Value, Worksheets("CC Profile Aug").Range("A:k"), 1, False)

Worksheets("demo").Cells(i, 9).Value = Application.WorksheetFunction.VLookup( _
Worksheets("demo").Cells(i, 2).Value, Worksheets("CC Profile Aug").Range("A:k"), 9, False)


Worksheets("demo").Cells(i, 10).Value = Application.WorksheetFunction.VLookup( _
Worksheets("demo").Cells(i, 2).Value, Worksheets("CC Profile Aug").Range("A:k"), 8, False)

Worksheets("demo").Cells(i, 11).Value = Application.WorksheetFunction.VLookup( _
Worksheets("demo").Cells(i, 2).Value, Worksheets("CC Profile Aug").Range("A:k"), 10, False)


Worksheets("demo").Cells(i, 12).Value = Application.WorksheetFunction.VLookup( _
Worksheets("demo").Cells(i, 2).Value, Worksheets("CC Profile Aug").Range("A:k"), 6, False)

Next

'The operation Id needs to be extract mannuallty to avoid error before running the next macro
'The CCs# needs to be coverted to text format


End Sub

Sub Addchecks()

'Step4 - Add check section to the summary report to reflect the comparision result.
'Employee information and division information are also added


Worksheets("demo").Range("M2").Value = "CCs#"
Worksheets("demo").Range("N2").Value = "Target Range"
Worksheets("demo").Range("O2").Value = "Current Methodology"
Worksheets("demo").Range("P2").Value = "LOB"
Worksheets("demo").Range("Q2").Value = "Operations"
Worksheets("demo").Range("R2").Value = "Check"
Worksheets("demo").Range("S2").Value = "Division"
Worksheets("demo").Range("T2").Value = "Name"
Worksheets("demo").Range("U2").Value = "Status"
Worksheets("demo").Range("V2").Value = "LastUpdateDate"
Worksheets("demo").Range("W2").Value = "Comments"



LastRow = Worksheets("demo").Cells(Rows.Count, 1).End(xlUp).Row

For i = 3 To LastRow

If Worksheets("demo").Cells(i, 2).Value = Worksheets("demo").Cells(i, 8).Value Then

Worksheets("demo").Cells(i, 13).Value = "TRUE"

Else: Worksheets("demo").Cells(i, 13).Value = "FALSE"

End If


If Worksheets("demo").Cells(i, 4).Value = Worksheets("demo").Cells(i, 9).Value Then

Worksheets("demo").Cells(i, 14).Value = "TRUE"

Else: Worksheets("demo").Cells(i, 14).Value = "FALSE"

End If


If Worksheets("demo").Cells(i, 5).Value = Worksheets("demo").Cells(i, 10).Value Then

Worksheets("demo").Cells(i, 15).Value = "TRUE"

Else: Worksheets("demo").Cells(i, 15).Value = "FALSE"

End If


If Worksheets("demo").Cells(i, 6).Value = Worksheets("demo").Cells(i, 11).Value Then

Worksheets("demo").Cells(i, 16).Value = "TRUE"

Else: Worksheets("demo").Cells(i, 16).Value = "FALSE"

End If


If Worksheets("demo").Cells(i, 7).Value = Worksheets("demo").Cells(i, 12).Value Then

Worksheets("demo").Cells(i, 17).Value = "TRUE"

Else: Worksheets("demo").Cells(i, 17).Value = "FALSE"

End If



Next

'Add additional information like division, name, status, etc. . find the the match in "result" sheet

On Error Resume Next

For j = 3 To Worksheets("demo").UsedRange.Rows.Count

Worksheets("demo").Cells(j, 19).Value = Application.WorksheetFunction.VLookup( _
Worksheets("demo").Cells(j, 2).Value, Worksheets("result").Range("A:W"), 14, False)

Worksheets("demo").Cells(j, 20).Value = Application.WorksheetFunction.VLookup( _
Worksheets("demo").Cells(j, 2).Value, Worksheets("result").Range("A:W"), 23, False)




Next

' Add the dropdown list and check summmary to the report


Dim last_row  As Long

last_row = Cells(Rows.Count, 1).End(xlUp).Row

Range("U3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="Confirming Info,Change Request Sent,Action Completed"


Range("U3").AutoFill Destination:=Range("U3:U" & last_row)

last_row = Cells(Rows.Count, 1).End(xlUp).Row

Range("R3").Formula = "=if(countif(M3:Q3,""TRUE"")<5,""N"",""Y"")"

Range("R3").AutoFill Destination:=Range("R3:R" & last_row)

End Sub

Sub conditionalformat()

' Step5 - Add conditinal format the comparative report

Dim myrange As Range
Dim condition1 As FormatCondition, condition2 As FormatCondition, condition3 As FormatCondition

lr = ActiveSheet.UsedRange.Rows.Count

Set myrange = Range("M3:Q3" & lr)
myrange.FormatConditions.Delete

Set condition1 = myrange.FormatConditions.Add(xlCellValue, xlEqual, "TRUE")

Set condition2 = myrange.FormatConditions.Add(xlCellValue, xlEqual, "FALSE")

Set condition3 = myrange.FormatConditions.Add(xlCellValue, xlEqual, "=0")

With condition1
.Interior.Color = vbGreen

End With

With condition2
.Interior.Color = vbRed
     
End With

With condition3
.Interior.Color = vbWhite
     
End With

End Sub

Sub FilteredByDivision()

'Step 6 - filter the summary report by division, and paste the filtered result to a new tab name after the division
'Use the dropdown list at cell P1 to filter the summary report

division = Cells(1, 16).Value

Worksheets("Demo").Range("A2").AutoFilter Field:=21, Criteria1:=Cells(1, 16).Value, Operator:=xlFilterValues

Worksheets("Demo").Range("A:Y").Copy

'Name the new sheet by division name and paste the filtered result to the new sheet

If Len(division) < 31 Then

Sheets.Add.Name = division

Worksheets(division).Range("A1").PasteSpecial Paste:=xlPasteValues


Application.CutCopyMode = False


ElseIf Len(division) > 31 Then

sheetname = Left(Cells(1, 16).Value, 5)

Sheets.Add.Name = sheetname

Worksheets(sheetname).Range("A1").PasteSpecial Paste:=xlPasteValues


Application.CutCopyMode = False

End If


End Sub
