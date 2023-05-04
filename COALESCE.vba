' usage: =COALESCE(A1,B1,C1) returns first non blank cell
Public Function COALESCE(ParamArray Fields() As Variant) As Variant
    Dim v As Variant
    For Each v In Fields
        If "" & v <> "" Then
            COALESCE = v
            Exit Function
        End If
    Next
    COALESCE = "" 'desired default value if all cells are blank
End Function
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
 ' usage: =COALESCE_ARRAY(A1:C1) returns first non blank cell
Public Function COALESCE_ARRAY(rng As Range) As Variant
    Dim cell As Range
    For Each cell In rng.Cells
        If "" & cell.Value <> "" Then
            COALESCE_ARRAY = cell.Value
            Exit Function
        End If
    Next
    COALESCE_ARRAY = "" 'desired default value if all cells are blank
End Function

