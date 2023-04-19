' retrieves link from href html imported to excel
' basically opposite of hyperlink()
Public Function REVERSEHYPERLINK(c As Range) As String
    On Error Resume Next
    REVERSEHYPERLINK = c.Hyperlinks(1).Address
End Function
