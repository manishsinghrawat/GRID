Attribute VB_Name = "postcal"
Public opr(100) As String
Public fnc(100) As String
Dim expression(1000) As String
Function preprocess(eqn As String, var As String, val As String)


'let x be the Known variable
eqn = UCase(eqn)

Do While Not InStr(1, eqn, "PI") = 0
eqn = Replace(eqn, "PI", CStr(22 / 7), 1)
Loop

Do While Not InStr(1, eqn, "E") = 0
eqn = Replace(eqn, "E", "2.71828", 1)
Loop

eqn = "(" + eqn + ")"

'main procedure
'following procedure will replace any x present in String
Do While Not InStr(1, eqn, UCase(var)) = 0
eqn = Left(eqn, InStr(1, eqn, UCase(var)) - 1) + CStr(Round(val, 10)) + Right(eqn, Len(eqn) - InStr(1, eqn, UCase(var)))
Loop

preprocess = eqn
End Function

Function leng(stack() As String)
i = 0
While (Not stack(i) = vbNullString)
i = i + 1
Wend
leng = i
End Function

Function reverse(stack() As String)
Dim newstack(1000) As String

length = leng(stack)

newstack(length) = vbNullString
For i = 0 To length - 1
newstack(length - i - 1) = stack(i)
Next
reverse = newstack
End Function

Function calc(var As String, value As String)
On Error GoTo errr
Dim stack(1000) As String
stacko = expression

i = 0
While (Not stacko(i) = vbNullString)
i = i + 1
Wend
lll = i
i = 0
For i = 0 To lll - 1
stack(i) = stacko(i)
Select Case stack(i)
Case ";"
 stack(i) = CStr(Round(22 / 7, 10))
 Case ","
 stack(i) = CStr("2.71828")
 Case UCase(var)
 stack(i) = CStr(value)
End Select
 
Next
stack(i) = vbNullString

Dim ele As String
Dim o1 As String
Dim o2 As String
Dim newstack(1000) As String
topn = -1

For tops = 0 To leng(stack) - 1
ele = stack(tops)
If isopr(ele) = 1 Then
If isfnc(ele) = True Then
o1 = newstack(topn)
topn = topn - 1
newstack(topn + 1) = CStr(apfnc(o1, ele))
topn = topn + 1
Else
o2 = newstack(topn)
topn = topn - 1
o1 = newstack(topn)
topn = topn - 1
newstack(topn + 1) = CStr(apopr(o1, o2, ele))
topn = topn + 1
End If
Else
newstack(topn + 1) = ele
topn = topn + 1
End If
Next

calc = CDbl(newstack(0))
Exit Function
errr:
calc = "..."
End Function

Function apfnc(o1 As String, char As String) As Double
Dim o11 As Double
o11 = CDbl(o1)
num = Asc(char) - Asc("A")
f = fnc(num)

Select Case f
Case "SIN"
apfnc = Sin(o11)
Case "COS"
apfnc = Cos(o11)
Case "TAN"
apfnc = Tan(o11)
Case "LN"
apfnc = Log(o11)
Case "SQR"
apfnc = Sqr(o11)
Case "ABS"
apfnc = Abs(o11)
Case "SGN"
apfnc = Sgn(o11)
Case "GIT"
If o11 > 0 Then
apfnc = Fix(o11)
Else
apfnc = Fix(o11) - 1
End If
End Select


End Function

Function isfnc(ch As String) As Boolean
If Asc(ch) <= Asc("Z") And Asc(ch) >= Asc("A") Then
isfnc = True
Else
isfnc = False
End If

End Function
Function apopr(o1 As String, o2 As String, char As String) As Double
Dim o11 As Double
Dim o22 As Double

o11 = CDbl(Round(o1, 10))

o22 = CDbl(o2)
Select Case char
Case "^"
apopr = o11 ^ o22
Case "/"
apopr = o11 / o22
Case "*"
apopr = o11 * o22
Case "+"
apopr = o11 + o22
Case "-"
apopr = o11 - o22
End Select

End Function
Function replaceall(str As String, find As String, by As String)

Do While Not InStr(1, str, find) = 0
str = Replace(str, find, by, 1)
Loop
replaceall = str
End Function


Function op()
opr(0) = "("
opr(1) = "-"
opr(2) = "+"
opr(3) = "*"
opr(4) = "/"
opr(5) = "^"
opr(6) = vbNullString
fnc(0) = "SIN"
fnc(1) = "COS"
fnc(2) = "TAN"
fnc(3) = "LN"
fnc(4) = "SQR"
fnc(5) = "ABS"
fnc(6) = "SGN"
fnc(7) = "RND"
fnc(8) = "GIT"
fnc(9) = vbNullString



End Function

Function isopr(str As String)
isopr = 0
For i = 0 To 100
If opr(i) = str Then
isopr = 1
ElseIf opr(i) = vbNullString Then
Exit For
End If
Next

If str = "X" Then
Exit Function
End If

For i = 0 To 100
If Chr(i + Asc("A")) = str Then
isopr = 1
ElseIf fnc(i) = vbNullString Then
Exit For
End If
Next

End Function

Function getpr(str As String)
getpr = -1
For i = 0 To 100
If opr(i) = str Then
getpr = i
ElseIf opr(i) = vbNullString Then
Exit For
End If
Next


If isfnc(str) = True Then
getpr = 1000
End If


End Function

Function crpost(eqn As String)
eqn = replaceall(eqn, "E", ",")
eqn = replaceall(eqn, "PI", ";")

eqn = "(" + UCase(eqn) + ")"

i = 0
While (Not fnc(i) = vbNullString)
eqn = replaceall(eqn, fnc(i), Chr(Asc("A") + i))
i = i + 1
Wend

Dim ostack(1000) As String
Dim otop As Integer
otop = -1

Dim etop As Integer
etop = -1


For start = 1 To Len(eqn)
cr = Mid(eqn, start, 1)

If cr = "(" Then
ostack(otop + 1) = cr
otop = otop + 1

If isfnc(Mid(eqn, start + 1, 1)) = False Or Mid(eqn, start + 1, 1) = "X" Then
If Not Mid(eqn, start + 1, 1) = "(" Then
expression(etop + 1) = Mid(eqn, start + 1, 1)
start = start + 1
etop = etop + 1
End If
End If

ElseIf cr = ")" Then
While Not ostack(otop) = "("
expression(etop + 1) = ostack(otop)
etop = etop + 1
otop = otop - 1
Wend
otop = otop - 1

ElseIf isopr(CStr(cr)) = 1 Then

If getpr(ostack(otop)) > getpr(CStr(cr)) Then
While getpr(ostack(otop)) > getpr(CStr(cr))
expression(etop + 1) = ostack(otop)
etop = etop + 1
otop = otop - 1
Wend
End If
ostack(otop + 1) = cr
otop = otop + 1

If isfnc(Mid(eqn, start + 1, 1)) = False Or Mid(eqn, start + 1, 1) = "X" Then
If Not Mid(eqn, start + 1, 1) = "(" Then
expression(etop + 1) = Mid(eqn, start + 1, 1)
start = start + 1
etop = etop + 1
End If
End If

Else
expression(etop) = expression(etop) + cr
End If

Next
expression(etop + 1) = vbNullString
etop = etop + 1


End Function
