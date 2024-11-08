Function makeItCute(sInput As String, pattern As String) As String
   Dim regex As Object
   Set regex = CreateObject("VBScript.RegExp")
   With regex
       .Global = True
       .Pattern = pattern
       makeItCute = .Replace(sInput, "")
   End With
End Function



'USAGE:  =MakeItCute([@[Cell]], "^[a-z]{1,3}") 

