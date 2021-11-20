Attribute VB_Name = "Module4"
Function winsor(x As Range, alph As Double)
Dim low, up As Double
Dim i, n As Integer
n = x.Rows.Count
Dim U() As Double
ReDim U(1 To n, 1 To 1)
Dim C() As Double
ReDim C(1 To n, 1 To 1)
low = Application.WorksheetFunction.Percentile(x, alph)
up = Application.WorksheetFunction.Percentile(x, 1 - alph)
For i = 1 To n
C(i, 1) = Application.WorksheetFunction.Max(x(i), low)
U(i, 1) = Application.WorksheetFunction.Min(C(i, 1), up)
Next i
winsor = U


End Function

