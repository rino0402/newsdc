Option Explicit
Dim strA, strB
Dim regEx, Matches, Match
Set regEx = New RegExp

strA = "\d+"
strB = "000aaa111"

regEx.Pattern = strA
regEx.Global = True
regEx.IgnoreCase = False

Set Matches = regEx.Execute(strB)
For Each Match in Matches
  WScript.Echo Match.FirstIndex & ":'" & Match.value & "'"
Next
