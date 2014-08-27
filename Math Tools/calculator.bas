Attribute VB_Name = "calculator"
 
 ' part of my list of standard functions
Function calculate(strx As String, varx As String, var As String)
On Error GoTo error
Dim mem As String
'let x be the Known variable
mem = UCase(strx)

Do While Not InStr(1, mem, "PI") = 0
mem = Replace(mem, "PI", CStr(22 / 7), 1)
Loop

Do While Not InStr(1, mem, "E") = 0
mem = Replace(mem, "E", "2.7", 1)
Loop



mem = "(" + mem + ")"
'main procedure
'following procedure will replace any x present in String
Do While Not InStr(1, mem, UCase(varx)) = 0

mem = Left(mem, InStr(1, mem, UCase(varx)) - 1) + CStr(var) + Right(mem, Len(mem) - InStr(1, mem, UCase(varx)))

Loop


Do While Not InStr(1, mem, "(") = 0
'reset all variables

'following will help in retrival of brackets
'maximum number of brackets allowed are 10

'******************************************************************
'set variable value
tx = 0
tx1 = 0
tex = 1

For temp = 1 To 10
If Not InStr(tex, mem, "(") = 0 Then
tx = InStr(tex, mem, "(")
tex = tx + 1
End If
Next

tex = 1
For temp1 = 1 To 10
If Not InStr(tex, mem, ")") = 0 Then
If InStr(tex, mem, ")") > tx Then
tx1 = InStr(tex, mem, ")")
End If
tex = tx + 1
End If
Next

bracket = Mid(mem, tx + 1, tx1 - tx - 1)
start = tx - 1
ending = Len(mem) - tx1
'*******************************************************************

'----------------------------------------------------------------------
tx = 0
tx1 = 0
'solve the exponential functions
Do While Not InStr(1, bracket, "^") = 0

'number of digits in temp will be 10
For temp = 1 To Len(bracket) + 3
If Not InStr(1, bracket, "^") - temp < 1 Then
If IsNumeric(Mid(bracket, InStr(1, bracket, "^") - temp, temp)) = True Then
tx = Mid(bracket, InStr(1, bracket, "^") - temp, temp)
End If
End If
Next

' get power value
For temp = 1 To Len(bracket) + 3
If Not Mid(bracket, InStr(1, bracket, "^") + temp, 1) = "-" Then
If Not Mid(bracket, InStr(1, bracket, "^") + temp, 1) = "+" Then
If IsNumeric(Mid(bracket, InStr(1, bracket, "^") + 1, temp)) = True Then
tx1 = Mid(bracket, InStr(1, bracket, "^") + 1, temp)
End If
End If
End If
Next
tex = tx1
ext = tx ^ tx1
bracket = Left(bracket, InStr(1, bracket, "^") - Len(tx) - 1) + CStr(ext) + Right(bracket, Len(bracket) - Len(tex) - InStr(1, bracket, "^"))
Loop
'--------------------------------------------------------------------
'Solve Bracket Complete


strx = bracket
'solve arithmetic====================================================

'division section //////////////////////////////////////////////////
Do While Not InStr(2, strx, "/") = 0
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
'/////////////////////////////////////////////////////////////////

' Multiplication section *****************************************
Do While Not InStr(2, strx, "*") = 0
'get first number
For temp = 1 To Len(strx) + 3
If Not InStr(2, strx, "*") - temp < 1 Then
If IsNumeric(Mid(strx, InStr(2, strx, "*") - temp, temp)) = True Then
tx = Mid(strx, InStr(2, strx, "*") - temp, temp)
End If
End If
Next

' get second number
For temp = 1 To Len(strx) + 3
If Not Mid(strx, InStr(2, strx, "*") + temp, 1) = "-" Then
If Not Mid(strx, InStr(2, strx, "*") + temp, 1) = "+" Then
If IsNumeric(Mid(strx, InStr(2, strx, "*") + 1, temp)) = True Then
tx1 = Mid(strx, InStr(2, strx, "*") + 1, temp)
End If
End If
End If
Next

'set temporary variables
tempx = Len(tx)
tempy = Len(tx1)

ext = CStr(CDbl(tx) * CDbl(tx1))
If ext >= 0 Then
sign = "+"
Else
sign = ""
End If
ext = sign + CStr(ext)
strx = Left(strx, InStr(2, strx, "*") - tempx - 1) + ext + Right(strx, Len(strx) - tempy - InStr(2, strx, "*"))
Loop
'*******************************************************************
'set temporary string to solve large equations
Dim emu As String
'set value of emulater variable
emu = strx
'replace E and Positive or negative sign
Do While Not InStr(1, emu, "E") = 0
If Not InStr(1, emu, "E+") = 0 Then
emu = Replace(emu, "E+", "Pl", 1)
ElseIf Not InStr(1, emu, "E-") = 0 Then
emu = Replace(emu, "E-", "mi", 1)
Else
emu = Replace(emu, "E", "P", 1)
End If
Loop

' addition  section +++++++++++++++++++++++++++++++++++++++++++++++++
Do While Not InStr(2, emu, "+") = 0
addsgn = InStr(2, emu, "+")
'get first number
For temp = 1 To Len(strx) + 3
If Not addsgn - temp < 1 Then
If IsNumeric(Mid(strx, addsgn - temp, temp)) = True Then
tx = Mid(strx, addsgn - temp, temp)
End If
End If
Next

'get second number
For temp = 1 To Len(strx) + 3
If Not Mid(strx, addsgn + temp, 1) = "-" Then
If Not Mid(strx, addsgn + temp, 1) = "+" Then
If IsNumeric(Mid(strx, addsgn + 1, temp)) = True Then
tx1 = Mid(strx, addsgn + 1, temp)
End If
End If
End If
Next

'set temporary variables
tempx = Len(tx)
tempy = Len(tx1)

ext = CStr(CDbl(tx) + CDbl(tx1))
If ext >= 0 Then
sign = "+"
Else
sign = ""
End If
ext = sign + CStr(ext)
strx = Left(strx, addsgn - tempx - 1) + ext + Right(strx, Len(strx) - tempy - addsgn)

'set value of emulater variable
emu = strx
'replace E and Positive or negative sign
Do While Not InStr(1, emu, "E") = 0
If Not InStr(1, emu, "E+") = 0 Then
emu = Replace(emu, "E+", "Pl", 1)
ElseIf Not InStr(1, emu, "E-") = 0 Then
emu = Replace(emu, "E-", "mi", 1)
Else
emu = Replace(emu, "E", "P", 1)
End If
Loop


Loop
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' subtraction section -----------------------------------------------
Do While Not InStr(2, emu, "-") = 0
minsgn = InStr(2, emu, "-")
'get first number
For temp = 1 To Len(strx) + 3
If Not minsgn - temp < 1 Then
If IsNumeric(Mid(strx, minsgn - temp, temp)) = True Then
tx = Mid(strx, minsgn - temp, temp)
End If
End If
Next

' get second number
For temp = 1 To Len(strx) + 3
If Not Mid(strx, minsgn + temp, 1) = "-" Then
If Not Mid(strx, minsgn + temp, 1) = "+" Then
If IsNumeric(Mid(strx, minsgn + 1, temp)) = True Then
tx1 = Mid(strx, minsgn + 1, temp)
End If
End If
End If
Next

'set temporary variables
tempx = Len(tx)
tempy = Len(tx1)

ext = CStr(CDbl(tx) - CDbl(tx1))
If ext >= 0 Then
sign = "+"
Else
sign = ""
End If
ext = sign + CStr(ext)
strx = Left(strx, minsgn - tempx - 1) + ext + Right(strx, Len(strx) - tempy - minsgn)


'set value of emulater variable
emu = strx
'replace E and Positive or negative sign
Do While Not InStr(1, emu, "E") = 0
If Not InStr(1, emu, "E+") = 0 Then
emu = Replace(emu, "E+", "Pl", 1)
ElseIf Not InStr(1, emu, "E-") = 0 Then
emu = Replace(emu, "E-", "mi", 1)
Else
emu = Replace(emu, "E", "P", 1)
End If
Loop
Loop
'-------------------------------------------------------------------

'convert large equations into known formats
strx = strx
If Not start - 2 < 1 Then

If Mid(mem, start - 2, 3) = "SIN" Then
strx = Sin(strx)
start = start - 3
ElseIf Mid(mem, start - 2, 3) = "COS" Then
strx = Cos(strx)
start = start - 3
ElseIf Mid(mem, start - 2, 3) = "TAN" Then
strx = Tan(strx)
start = start - 3
ElseIf Mid(mem, start - 1, 2) = "LN" Then
strx = Log(strx)
start = start - 2
ElseIf Mid(mem, start - 2, 3) = "ABS" Then
strx = Abs(strx)
start = start - 3
ElseIf Mid(mem, start - 2, 3) = "SGN" Then
strx = Sgn(strx)
start = start - 3
ElseIf Mid(mem, start - 2, 3) = "SQR" Then
strx = Sqr(strx)
start = start - 3
ElseIf Mid(mem, start - 2, 3) = "GIT" Then
If strx < 0 Then
strx = Fix(strx) - 1
Else
strx = Fix(strx)
End If

start = start - 3
strx = CDbl(strx)
End If
End If
'///////////////////////////////////////////////////////////////////
' set bracklet solved value in memory string
mem = Left(mem, start) + strx + Right(mem, ending)

'///////////////////////////////////////////////////////////////////

Loop

error:
If err.number = 0 Then
result = mem
Else
result = "..."
End If
calculate = result

End Function

