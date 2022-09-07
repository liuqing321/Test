Attribute VB_Name = "Module1"
Sub copypaste()

'filter the report
Worksheets("Page 1").Range("A1").AutoFilter Field:=7, Criteria1:=Array("Awaiting User Info", "Open"), Operator:=xlFilterValues

'copy the filtered result and paste it to the result page
Worksheets("Page 1").Range("W:AB").Copy _
    Destination:=Worksheets("test").Range("A2")
    

End Sub


Sub lookupccs()

On Error Resume Next
Dim i As Integer

For i = 2 To Worksheets("test").UsedRange.Rows.Count


Worksheets("test").Cells(i, 8).Value = Application.WorksheetFunction.VLookup( _
Worksheets("test").Cells(i, 1).Value, Worksheets("CC Profile Single Month").Range("G:K"), 1, False)

Worksheets("test").Cells(i, 9).Value = Application.WorksheetFunction.VLookup( _
Worksheets("test").Cells(i, 1).Value, Worksheets("CC Profile Single Month").Range("G:K"), 3, False)


Worksheets("test").Cells(i, 10).Value = Application.WorksheetFunction.VLookup( _
Worksheets("test").Cells(i, 1).Value, Worksheets("CC Profile Single Month").Range("G:K"), 2, False)

Worksheets("test").Cells(i, 11).Value = Application.WorksheetFunction.VLookup( _
Worksheets("test").Cells(i, 1).Value, Worksheets("CC Profile Single Month").Range("G:K"), 4, False)

Worksheets("test").Cells(i, 12).Value = Application.WorksheetFunction.VLookup( _
Worksheets("test").Cells(i, 7).Value, Worksheets("CC Profile Single Month").Range("G:K"), 5, False)
Next

    Dim j As Integer
    For j = 2 To Worksheets("test").UsedRange.Rows.Count

    Worksheets("test").Cells(j, 13).Value = WorksheetFunction.Text(Worksheets("test").Cells(j, 8).Value, "00000")

    Next



End Sub

Sub matchccprofiles()

Worksheets("test").Range("H2").Value = "CCs#"
Worksheets("test").Range("I2").Value = "TargetRange"
Worksheets("test").Range("J2").Value = "Current Methodology"
Worksheets("test").Range("K2").Value = "LOB"
Worksheets("test").Range("L2").Value = "Operations"

End Sub
