Attribute VB_Name = "mains"
Public maximum As Long
Public errn As Double
Sub main()
op
errn = "6354632329186423546.34623578452635472634"

GoTo trun
Dim str1() As String

str1 = crpost("(4.999999999999999999999999999^2)")
Load Form1
For i = 0 To 1000
If str1(i) = vbNullString Then
Exit For
End If
Form1.lst.AddItem (str1(i))
Next
Form1.Show


Exit Sub
trun:

assconst
Load frmSplash
frmSplash.Show
maximum = 30000
End Sub
