Option Explicit
Option Base 1

Sub Summary_Report()

Dim FileNames As Variant, nw As Integer
Dim S() As Variant
Dim tWB As Workbook, aWB As Workbook
Dim nr As Integer, nc As Integer
Dim w As Integer, i As Integer, j As Integer

'Assignment 3 - STARTER.xlsm
Set tWB = ThisWorkbook  'ThisWorbook: Es el workbook sobre el que se escribe el código VBA.

FileNames = Application.GetOpenFilename( _
            FileFilter:="Excel Files (*.csv), *.csv", _
            Title:="Open File(s)", _
            MultiSelect:=True)

nw = UBound(FileNames)

nr = 4
nc = 2

ReDim S(nw, nr)

For w = 1 To nw
    tWB.Activate
    Workbooks.Open FileNames(w)
    Set aWB = ActiveWorkbook
    
    For i = 1 To nr
        For j = 1 To nc
            S(w, i) = Range("B" & i + 2)
        Next j
    Next i
    
    aWB.Close SaveChanges:=False
    tWB.Activate
Next w

tWB.Sheets("Data").Range("A1:D" & nw) = S
End Sub

Sub Reset_Report()
Cells.Clear
End Sub
