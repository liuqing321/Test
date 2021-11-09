Attribute VB_Name = "Module1"
Sub copyfunc()

Dim wscopy As Worksheet
Dim wsdest As Worksheet
Dim lcopylastrow As Long
Dim ldestlastrow As Long

Set wscopy = ActiveSheet
Set wsdest = ActiveWorkbook.Worksheets("summary")

lcopylastrow = wscopy.Cells(wscopy.Rows.Count, "A").End(xlUp).Row

ldestlastrow = wsdest.Cells(wsdest.Rows.Count, "A").End(xlUp).Offset(1).Row

'3. Copy & Paste Data
  wscopy.Range("A2:H" & lcopylastrow).Copy _
    wsdest.Range("A" & ldestlastrow)

End Sub
