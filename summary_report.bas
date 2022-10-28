Option Explicit
Option Base 1

Sub Summary_Report()
            
'----------------------------------------------------------------------
' Variables
'----------------------------------------------------------------------
            
    Dim FileNames As Variant    'Nombre de los archivos a resumir.
    Dim w As Integer            'Representa cada archivo a resumir.
    Dim nw As Integer           'Numero de los archivos a resumir. Cantidad de w's
    Dim S() As Variant          'Matriz para el contenido de los archivos a resumir.
    Dim tWB As Workbook         'ThisWorbook: Almacene el libro de trabajo donde se escribe el código VBA.
    Dim aWB As Workbook         'ActiveWorkbook: Almacena el libro de trabajo activo o sobreexpuesto ante los demás.
    Dim nr As Integer           'Número de filas del rango a extraer de los archivos a resumir.
    Dim nc As Integer           'Número de columnas del rango a extraer de los archivos a resumir.
    Dim i As Integer            'Variable que almacena cada valor de nr.
    Dim j As Integer            'Variable que almacena cada valor de nc.

'----------------------------------------------------------------------
' Código
'----------------------------------------------------------------------

    Set tWB = ThisWorkbook

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
