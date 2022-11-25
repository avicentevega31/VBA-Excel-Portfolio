Option Explicit
Option Base 1

Sub AddFive_Alt()
    Dim i As Integer, j As Integer
    Dim nc As Integer, nr As Integer
    
    nr = Selection.Rows.Count
    nc = Selection.Columns.Count
    
    For i = 1 To nr
        For j = 1 To nc
            Selection.Cells(i, j).Offset(nr + 1, 0) = Selection.Cells(i, j) + 5
        Next j
    Next i
End Sub
