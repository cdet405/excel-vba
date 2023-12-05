Function makeItCute(sInput As String) As String
   Dim regex As Object
   Set regex = CreateObject("VBScript.RegExp")
   With regex
       .Global = True
       .Pattern = "[^a-zA-Z0-9\s]"
       makeItCute = .Replace(sInput, "")
   End With
End Function



'USAGE:  =MakeItCute([@[Cell]]) 

