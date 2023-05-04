Public Function COALESCE(ParamArray Fields() As Variant) As Variant

    Dim v As Variant

    For Each v In Fields
        If "" & v <> "" Then
            Coalesce = v
            Exit Function
        End If
    Next
    Coalesce = ""

End Function

' usage: =COALESCE(A1,B1,C1) returns first non blank cell
