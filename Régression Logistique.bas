Attribute VB_Name = "Module3"
Function logit(x As Range, y As Range)
Dim n As Integer
Dim k As Integer
Dim U As Double
Dim r As Double
Dim f As Double
Dim q As Double
Dim p As Double
Dim i As Integer, s As Integer, j As Integer, m As Integer, z As Integer
Dim eps As Double
eps = 0.000001


n = x.Rows.Count
k = x.Columns.Count
Dim b() As Double
ReDim b(1 To k + 1)
Dim grad() As Double
ReDim grad(1 To k + 1, 1 To 1)

For i = 2 To k + 1
b(i) = 0
Next i
p = 0
For i = 1 To n
p = p + y(i)
Next i
p = p / n
b(1) = Log(p / (1 - p))
Dim norm As Double
iter = 0
norm = 1
While (iter < 15) And (norm > eps)
Dim h() As Double
ReDim h(1 To n, 1 To 1)
Dim v() As Double
ReDim v(1 To n, 1 To k + 1)

For i = 1 To n
v(i, 1) = 1
For j = 1 To k
v(i, j + 1) = x(i, j)
Next j
Next i

For i = 1 To n
o = 0
For j = 1 To k + 1
o = o + v(i, j) * b(j)
Next j
h(i, 1) = o
Next i

For j = 1 To k + 1
q = 0
For s = 1 To n
r = 1 / (Exp(-h(s, 1)) + 1)
q = q + (y(s) - r) * v(s, j)
Next s
grad(j, 1) = q
Next j
'logit = grad



Dim hes() As Double
ReDim hes(1 To k + 1, 1 To k + 1)

For m = 1 To k + 1
For j = 1 To k + 1
U = 0
For s = 1 To n
f = 1 / (Exp(-h(s, 1)) + 1)
U = U + (1 - f) * f * v(s, j) * v(s, m)
Next s
hes(m, j) = -U
Next j
Next m
'logit = hes

Dim C()
C = Application.WorksheetFunction.MInverse(hes)
Dim e()
e = Application.WorksheetFunction.MMult(C, grad)
norm = 0
For i = 1 To k + 1:
b(i) = b(i) - e(i, 1)
norm = norm + Abs(e(i, 1))
Next i
iter = iter + 1
Wend


'______ test wald'

Dim t(), se() As Double
ReDim t(1 To k + 1)
ReDim se(1 To k + 1)
Dim pval() As Double
ReDim pval(1 To k + 1)
For j = 1 To k + 1
se(j) = Sqr(-C(j, j))
t(j) = b(j) / se(j)
pval(j) = (1 - Application.WorksheetFunction.NormSDist(Abs(t(j)))) * 2
Next j
'logit = b

Dim ybar As Double
ybar = Application.WorksheetFunction.Average(y)

Dim aa As Double
aa = n * (ybar * Log(ybar) + (1 - ybar) * Log(1 - ybar))
Dim tt() As Double
ReDim tt(1 To n)
Dim a As Double
For i = 1 To n
a = 0
For j = 1 To k + 1
a = a + v(i, j) * b(j)
Next j
tt(i) = a
Next i

Dim cc, pp As Double
cc = 0

For i = 1 To n
pp = 1 / (Exp(-tt(i)) + 1)
cc = cc + y(i) * Log(pp) + (1 - y(i)) * Log(1 - pp)
Next i

Dim psR As Double
psR = 1 - (cc / aa)

Dim leg() As Double
ReDim leg(1 To 4, 1 To k + 1)
For i = 1 To k + 1
leg(1, i) = b(i)
leg(2, i) = se(i)
leg(3, i) = pval(i)
leg(4, 1) = psR
Next i
'logit = cc
'logit = leg
logit = h


End Function




