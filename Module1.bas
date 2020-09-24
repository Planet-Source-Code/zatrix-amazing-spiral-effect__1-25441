Attribute VB_Name = "Module1"
'****************************************************************'
'This code is copyright to Alex K!                               '
'zatrix@load.com                                                 '
'You my use this code anyway you wish but you cannot change it!  '
'Unless you find a bug!                                          '
'If you find a bug please report it at zatrix@loa.com            '
'*****************************************************************
Global re As Integer
Global gr As Integer
Global bl As Integer
Global SStop As Boolean
Sub DrawRoullette()
Dim R1, R2, r, pi
R1 = Form1.HScroll1.Value
R2 = Form1.HScroll2.Value - 80
If R2 = 0 Then R2 = 10
r = Form1.HScroll3.Value
pi = 4 * Atn(1)
Dim loop1, loop2
Dim t, X, Y As Double
Dim Rotations As Integer
If Int(R1 / R2) = R1 / R2 Then
Rotations = 1
Else
Rotations = Abs(R2 / 10)
If Int(R2 / 10) <> R2 / 10 Then Rotations = 10 * Rotations
End If
For loop1 = 1 To Rotations
If BreaNow Then
Form1.Command1.Caption = "Start"
BreakNow = False
Exit Sub
End If
For loop2 = 0 To 2 * pi Step pi / (4 * 360)
If SStop = True Then Exit Sub
t = loop1 * 2 * pi + loop2
X = (R1 + R2) * Cos(t) - (R2 + r) * Cos(((R1 + R2) / R2) * t)
Y = (R1 + R2) * Sin(t) - (R2 + r) * Sin(((R1 + R2) / R2) * t)
Form1.Picture1.PSet (Form1.Picture1.ScaleWidth / 2 + X, Form1.Picture1.ScaleHeight / 2 + Y), RGB(re, gr, bl)
Next
DoEvents
Next
Form1.Command1.Caption = "Start"
BreakNow = False
End Sub
