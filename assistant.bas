Attribute VB_Name = "assistant"
Sub disable()
For temp = 0 To 17
calc.but(temp).Enabled = False
Next
calc.functions.Enabled = False
calc.values.Enabled = False
calc.signs.Enabled = False
End Sub

Sub numwo0()
For temp = 0 To 10
If Not temp = 10 Then
calc.but(temp).Enabled = True
End If
Next
End Sub

Sub num()
For temp = 0 To 10
calc.but(temp).Enabled = True
Next
End Sub

Sub asdm()
For temp = 11 To 14
calc.but(temp).Enabled = True
Next
End Sub

Sub pm()
calc.but(11).Enabled = True
calc.but(12).Enabled = True
End Sub

Sub funct()
calc.functions.Enabled = True
End Sub

Sub constant()
calc.values.Enabled = True
End Sub

Sub variable()
If fun = True Then
calc.signs.Enabled = True
End If
End Sub

Sub stb()
calc.but(16).Enabled = True
End Sub

Sub endb()
calc.but(17).Enabled = True
End Sub

Sub sqrs()
calc.but(15).Enabled = True
End Sub

Function crad(number As Double)
crad = ((22 / 7) / 180) * number
End Function

Function arcsin(number As Double)
intvl = 3
For temp = -22 / 7 + (10 / (10 ^ intvl)) To (22 / 7) Step (1 / (10 ^ intvl))
If Round(Sin(temp), intvl) = Round(number, intvl) - 0.001 Then
arcsin = Round(temp, intvl)
End If
Next
End Function

Function arccos(number As Double)
intvl = 3
For temp = -22 / 7 + (10 / (10 ^ intvl)) To (22 / 7) Step (1 / (10 ^ intvl))
If Round(Cos(temp), intvl) = Round(number, intvl) Then
arccos = Round(temp, intvl)
End If
Next
End Function


Function arcsins(number As Double)
number = Round(number, 3)
For temp = -22 / 7 To 22 / 7 Step 0.001
If Round(Sin(temp), 3) = Round(number, 3) Then
arcsins = Round(temp, 3)
End If
Next
End Function
Function arccoss(number As Double)
number = Round(number, 3)
For temp = -22 / 7 To 22 / 7 Step 0.001
If Round(Cos(temp), 3) = Round(number, 3) Then
arccoss = Round(temp, 3)
End If
Next
End Function
