# VBA_Segregate-data_by_Column
There is a simple VBA code to segregate data by selecting desired column.
Option Explicit

Sub DeleteSheet()
Dim ws As Worksheet
Application.DisplayAlerts = False
Application.ScreenUpdating = False
For Each ws In Worksheets
If ws.Name <> "Source Data" Then
ws.Delete
Else
End If
Next ws
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub


Sub seggregation3()
Dim wb As Workbook, ws As Worksheet, SeggregateBy As String, x As String, lr As Integer, address As String, search As String, cn As String
Set wb = ThisWorkbook
Set ws = wb.Sheets("Source Data")
Application.DisplayAlerts = False
Application.ScreenUpdating = False
ws.Range("a1").Select
On Error Resume Next
search = Application.InputBox(prompt:="Select column")
SeggregateBy = Cells.Find(what:=search, after:=ws.Range("a1"), searchorder:=xlColumns, MatchCase:=False).Select
address = Selection.address
cn = Selection.Column
Sheets.Add(before:=Sheets(1)).Name = "Support"
ws.Range("a1").AutoFilter = False
ws.Range(address).EntireColumn.Copy
Sheets("Support").Range("a1").PasteSpecial xlPasteValues
Sheets("Support").Range("a1").RemoveDuplicates Columns:=1, Header:=xlYes
lr = Sheets("Support").Range("a" & Rows.Count).End(xlUp).Row
For lr = 2 To lr
ws.Range("a1").AutoFilter = False
x = Sheets("Support").Range("a" & lr)
Sheets.Add(after:=Sheets(Sheets.Count)).Name = (ws.Range(address).Value & "-" & x)
ws.Range("a1").AutoFilter field:=cn, Criteria1:=x
ws.Range("a1").CurrentRegion.Copy
Sheets(ws.Range(address).Value & "-" & x).Range("a1").PasteSpecial xlPasteValues
Sheets(ws.Range(address).Value & "-" & x).Range("a1").PasteSpecial xlPasteFormats
Sheets(ws.Range(address).Value & "-" & x).Range("a1").Select
Sheets(ws.Range(address).Value & "-" & x).Range("a1").CurrentRegion.EntireColumn.AutoFit
Next lr

MsgBox "Data Segregated by " & ws.Range(address).Value
ws.Range("a1").AutoFilter = False
Sheets("Support").Delete
ws.Select
ws.Range("a1").Select
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub


