Option Explicit
Dim strA, strB
Dim regEx
Set regEx = New RegExp

strA = "^te3-0st$"
strB = "tE3-0sT"

regEx.Pattern = strA
regEx.Global = False
regEx.IgnoreCase = True

If regEx.Test(strB) Then
  MsgBox "B true"
Else
  MsgBox "B false"
End If
