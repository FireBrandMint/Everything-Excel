Sub Add1()
    Dim area As Range
    Dim cell As Range
    Dim char As String
    Dim cell_str As String
    
    For Each area In Selection.Areas
        For Each cell In area
            If IsError(cell.Value) Then
                GoTo skip
            End If
            
            cell_str = CStr(cell.Value)
            
            If cell_str = "" Or Not IsNumeric(cell_str) Then
                GoTo skip
            End If
            
            cell.Value = cell.Value + 1
skip:
        Next cell
    Next area
End Sub
