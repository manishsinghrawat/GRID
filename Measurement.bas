Attribute VB_Name = "Measurement"
Function angle(X1, Y1, X, Y, X2, Y2)
On Error Resume Next
'get angle A
mem = Sqr((X1 - X) ^ 2 + (Y1 - Y) ^ 2)
If Y1 > Y Then
aa = 44 / 7 - arccos((X1 - X) / mem)
Else
aa = arccos((X1 - X) / mem)
End If

'get angle B
mem = Sqr((X2 - X) ^ 2 + (Y2 - Y) ^ 2)
If Y2 - Y > 0 Then
bb = 44 / 7 - arccos((X2 - X) / mem)
Else
bb = arccos((X2 - X) / mem)
End If

angle = Abs(aa - bb)
End Function

