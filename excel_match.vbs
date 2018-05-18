Sub Sort()
    N_source = 160
    N_ref = 507
    min_dist = 1000000
    min_index = 0
    
    start_col = 1
    
    source_col_X = start_col
    source_col_Y = start_col + 1
    source_col_Mag = start_col + 2
    
    dest_col_X = start_col + 4
    dest_col_Y = start_col + 5
    dest_col_Mag = start_col + 6
    
    ref_col_X = start_col + 8
    ref_col_Y = start_col + 9
    ref_col_Mag = start_col + 10
    
    For i = 1 To N_source
        For j = 1 To N_ref
            dist = (Cells(i, source_col_X).Value - Cells(j, ref_col_X).Value) ^ 2 + (Cells(i, source_col_Y).Value - Cells(j, ref_col_Y).Value) ^ 2
            If dist < min_dist Then
                min_dist = dist
                min_index = j
            End If
        Next j
        Cells(min_index, dest_col_X).Value = Cells(i, source_col_X).Value
        Cells(min_index, dest_col_Y).Value = Cells(i, source_col_Y).Value
        Cells(min_index, dest_col_Mag).Value = Cells(i, source_col_Mag).Value
        min_dist = 1000000
        min_index = 0
    Next i
End Sub