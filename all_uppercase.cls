Sub AllToUppercase()
    Dim area As Range
    Dim cell As Range
    Dim i As Long
    Dim i2 As Long
    Dim i_upper As Long
    Dim is_weird As Boolean
    Dim char As String
    Dim cell_str As String
    Dim separators As Variant
    separators = Array(" ", "/")
    
    For Each area In Selection.Areas
        For Each cell In area
            If IsError(cell.Value) Then
                GoTo skip
            End If
            
            cell_str = CStr(cell.Value)
            i_upper = -1
            
            If cell_str = "" Then
                GoTo skip
            End If
            
            For i = 1 To Len(cell_str)
            
                char = Mid(cell_str, i, 1)
                
                For i2 = LBound(separators) To UBound(separators)
                    If char = separators(i2) Then
                        If Not is_weird And Not i_upper = -1 Then
                            Mid(cell_str, i_upper, 1) = UCase(Mid(cell_str, i_upper, 1))
                        End If
                        
                        is_weird = False
                        i_upper = -1
                        
                        GoTo skip_i
                    End If
                Next i2
                
                If i_upper = -1 Then
                    i_upper = i
                End If
                
                If is_weird Or Not IsLetter(char) Then
                    is_weird = True
                End If
                
skip_i:
            Next i
            
            If Not is_weird And Not i_upper = -1 Then
                Mid(cell_str, i_upper, 1) = UCase(Mid(cell_str, i_upper, 1))
            End If
                
            is_weird = False
            i_upper = -1
            
            cell.Value = cell_str
skip:
        Next cell
    Next area
End Sub

Private Function IsLetter(ByVal character As String) As Boolean
    IsLetter = UCase$(character) <> LCase$(character)
End Function
