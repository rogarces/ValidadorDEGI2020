Module Module1
    Public Function nreg(Hoja As Microsoft.Office.Interop.Excel.Worksheet, nFila As Long, nColumna As Long)

        Do Until Hoja.Cells(nFila, nColumna).value = ""
            nFila = nFila + 1
        Loop
        Return nFila
    End Function
End Module
