Function makeItCute(sInput As String) As String
Dim sSpecChar As String
Dim i As Long
sSpecChar = "\/:*?™""®<>|&@#_+`©~;+=^$!"©°'â„¢â€Ã""
For i = 1 To Len(sSpecChar)
sInput = Replace$(sInput, Mid$(sSpecChar, i, 1), "")
Next
makeItCute = sInput
End Function



' =MakeItCute([@[Cell]]) 
' sSpecChar = "REMOVE THESE CHARs"
