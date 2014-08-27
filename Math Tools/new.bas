Attribute VB_Name = "new"
Function precal(eqn As String, var As String, val As String)
Dim eqn As String

'let x be the Known variable
eqn = UCase(strx)

Do While Not InStr(1, eqn, "PI") = 0
eqn = Replace(eqn, "PI", CStr(22 / 7), 1)
Loop

Do While Not InStr(1, eqn, "E") = 0
eqn = Replace(eqn, "E", "2.71828", 1)
Loop

eqn = "(" + eqn + ")"

'main procedure
'following procedure will replace any x present in String
Do While Not InStr(1, eqn, UCase(varx)) = 0
eqn = Left(eqn, InStr(1, eqn, UCase(varx)) - 1) + CStr(var) + Right(eqn, Len(eqn) - InStr(1, eqn, UCase(varx)))
Loop
cal (eqn)
End Function
Function cal(eqn As String)
On Error GoTo error
Dim res As Double
Dim num(1) As Double

While InStr(1, eqn, "(")
starting = Len(eqn) - InStrRev(eqn, "(", 1) + 1
ending = InStr(starting, eqn, ")")

bracket = Mid(eqn, starting + 1, ending - starting - 1) ' since formula for length is ending-starting+1

res = cal(bracket)
If starting > 3 Then
func = Mid(eqn, starting - 4, 3)
num = appfunc(res, func)
Else
num(0) = res
num(1) = 0
End If

If num(1) = 0 Then
eqn = Left(eqn, starting - 1) + CStr(num(0)) + Right(eqn, Len(eqn) - ending)
Else
eqn = Left(eqn, starting - 4) + CStr(num(0)) + Right(eqn, Len(eqn) - ending)
End If
Wend

solu (eqn)
solveop (eqn)
End Function


Function solveop(eqn As String, op As String)
Do While Not InStr(2, strx, op) = 0
'get divident
For temp = 1 To Len(strx) + 3
If Not InStr(2, strx, "/") - temp < 1 Then
If IsNumeric(Mid(strx, InStr(2, strx, "/") - temp, temp)) = True Then
tx = Mid(strx, InStr(2, strx, "/") - temp, temp)
End If
End If
Next

' get divisor
For temp = 1 To Len(strx) + 3
If Not Mid(strx, InStr(2, strx, "/") + temp, 1) = "-" Then
If Not Mid(strx, InStr(2, strx, "/") + temp, 1) = "+" Then
If IsNumeric(Mid(strx, InStr(2, strx, "/") + 1, temp)) = True Then
tx1 = Mid(strx, InStr(2, strx, "/") + 1, temp)
End If
End If
End If
Next

'set temporary variables
tempx = Len(tx)
tempy = Len(tx1)

ext = CStr(CDbl(tx) / CDbl(tx1))
If ext >= 0 Then
sign = "+"
Else
sign = ""
End If
ext = sign + CStr(ext)
strx = Left(strx, InStr(2, strx, "/") - tempx - 1) + ext + Right(strx, Len(strx) - tempy - InStr(2, strx, "/"))
Loop

End Function

Function solu(eqn As String)
'solve unary operators
eqn = replaceall(eqn, "++", "+")
eqn = replaceall(eqn, "--", "+")
eqn = replaceall(eqn, "+-", "-")
eqn = replaceall(eqn, "-+", "-")

eqn = replaceall(eqn, "--", "+")
eqn = replaceall(eqn, "++", "+")
eqn = replaceall(eqn, "+-", "-")
eqn = replaceall(eqn, "-+", "-")

eqn = replaceall(eqn, "+-", "-")
eqn = replaceall(eqn, "--", "+")
eqn = replaceall(eqn, "++", "+")
eqn = replaceall(eqn, "-+", "-")

eqn = replaceall(eqn, "-+", "-")
eqn = replaceall(eqn, "--", "+")
eqn = replaceall(eqn, "+-", "-")
eqn = replaceall(eqn, "++", "+")
solop = eqn

End Function
Function appfunc(num As Double, func As String)
Dim res(1) As Double
res(0) = num
res(1) = 1
appfunc = res
End Function
