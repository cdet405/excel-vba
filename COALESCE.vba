' usage: =COALESCE(A1,B1,C1) returns first non blank cell
Public Function COALESCE(ParamArray Fields() As Variant) As Variant
    Dim v As Variant
    For Each v In Fields
        If "" & v <> "" Then
            Coalesce = v
            Exit Function
        End If
    Next
    Coalesce = "" 'desired default value if all cells are blank
End Function


