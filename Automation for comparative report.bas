Attribute VB_Name = "Module1"
Sub copypaste()

'filter the report
Worksheets("Page 1").Range("A1").AutoFilter Field:=7, Criteria1:=Array("Awaiting User Info", "Open"), Operator:=xlFilterValues

'copy the filtered result and paste it to the result page
Worksheets("Page 1").Range("W:AB").Copy _
    Destination:=Worksheets("Result").Range("A:F")
    

End Sub


Sub matchccprofiles()

Worksheets("Result").Range("I1").Value = "CCs#"
Worksheets("Result").Range("J1").Value = "Current Methodology"
Worksheets("Result").Range("K1").Value = "TargetRange"
Worksheets("Result").Range("L1").Value = "LOB"
Worksheets("Result").Range("M1").Value = "Operations"

End Sub
