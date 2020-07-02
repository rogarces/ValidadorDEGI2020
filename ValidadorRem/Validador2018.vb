Option Explicit On
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Deployment.Application
Public Class Validador2019
    Dim xlExcel As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Public mRow As Integer = 0
    Public newpage As Boolean = True
    Dim NombreArchivo, CodigoEstablec, ValidarNombre, ValidaAno, ValidaMes, ValidaComuna, ValidaVersion, ValidaCodigo, ValidaEstable, ValidaDependencia, ValidaHojaControl, ValidaSerieREM As String
    Public Fechas As DateTime = System.DateTime.Today 'FECHA DEL PC
    Public A01(20), A02(20), A03(25), A04(20), A09(20), REMBM18(10) As Double
    Public A01VAL(100), A02VAL(100) As String
    Private Sub BtnAbrir_Click(sender As System.Object, e As System.EventArgs) Handles BtnAbrir.Click
        With OpenFileDialog1
            .Filter = "Establecimiento(.xlsm)|*.xlsm|Establecimiento(.xlsx)|*.xlsx"
            .DefaultExt = ".xlsm"
            .FileName = ""
        End With

        ' if para abrir un opendialog
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.TxtRuta.Text = ""
            NombreArchivo = Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            TxtRuta.Text = OpenFileDialog1.FileName.ToString
            Me.DataGridView1.Rows.Clear()
            Me.LBEstable.Text = ""
            Me.LBcodigo.Text = ""
            Me.LBSerie.Text = ""
            Me.LBcomuna.Text = ""
            Me.LBdependencia.Text = ""
            Me.LBmes.Text = ""
            Me.LBversion.Text = ""
            Me.LBaño.Text = ""
            Me.LblHojaControl.Text = ""
            Me.LBerrores.Text = ""
            BtnExportar.UseVisualStyleBackColor = True
            Me.LBLprogreso.Text = "0 %"
            ProgressBar1.Value = 0
            'valida series A y P
            Select Case Len(NombreArchivo) ' largo de archivo 9 digitos
                Case Is = 9 ' 9 DIGITOS EN EL LARGO DEL NOMBRE DEL ARCHIVO, SERIE A P D
                    Select Case Mid(NombreArchivo, 7, 1)
                        Case Is = "A"
                            CargaSerieA()
                        Case Is = "P"
                            CargaSerieP()
                        Case Is = "D"
                            CargaSerieD()
                        Case Else
                            MsgBox("El archivo seleccionado no corresponde a serie, Vuelva a seleccionar", MsgBoxStyle.Information, "Advertencia de Serie REM")
                    End Select
                Case Is = 10 '9 DIGITOS EN EL LARGO DEL NOMBRE DEL ARCHIVO
                    Select Case Mid(NombreArchivo, 7, 2)
                        Case Is = "BM"
                            CargaSerieBM()
                        Case Is = "BS"
                            MsgBox("la Serie BS, se encuentra en contrucción", MsgBoxStyle.Information)
                        Case Else
                            MsgBox("El archivo seleccionado no corresponde a serie, Vuelva a seleccionar", MsgBoxStyle.Information, "Advertencia de Serie REM")
                    End Select
                Case Else
                    MsgBox("El archivo seleccionado no corresponde a la serie, Vuelva a seleccionar", MsgBoxStyle.Information, "Advertencia de Serie REM")
            End Select
        Else
            MsgBox("Debe Seleccionar un establecimiento", MsgBoxStyle.Information, "ADVERTENCIA DE CANCELACIÓN")
        End If
    End Sub
    Private Sub BtnExportar_Click(sender As Object, e As EventArgs) Handles BtnExportar.Click
        Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
        Dim libro = ExcelApp.Workbooks.Add
        Dim Planilla = New Microsoft.Office.Interop.Excel.Worksheet

        Dim Fila As Integer = 6
        Dim Columna As Integer = 1
        Dim RowCount As Integer = DataGridView1.Rows.Count - 2
        Dim ColumnCount As Integer = DataGridView1.Columns.Count - 1

        Try
            Planilla = libro.Sheets("Hoja1")

            With Planilla
                ' .Name = "erores"
                .Range("A1").Value = "SERVICIO SALUD OSORNO"
                .Range("A1:F1").MergeCells = True
                .Range("A2").Value = "ESTABLECIMIENTO : " & Me.LBEstable.Text & " [ " & Me.ValidaCodigo & " ]" ' ESTABLECIMIENTO
                .Range("A2:F2").MergeCells = True
                .Range("A3").Value = "COMUNA : " & Me.LBcomuna.Text ' COMUNA
                .Range("A3:E3").MergeCells = True
                .Range("A4").Value = "MES : " & Me.LBmes.Text ' MES
                .Range("A4:E4").MergeCells = True
                .Range("A1:A4").Font.Bold = True
                .Columns("E").AutoFit()
                .Range("F4").Value = "TOTAL ERRORES : " & Me.LBerrores.Text ' TOTAL ERRORES
                .Range("F4").Font.Bold = True
                .Rows.Font.Size = 10
                .Rows.Font.Name = "Calibri"
            End With

            For nColumna As Integer = 0 To ColumnCount
                libro.Worksheets("Hoja1").Cells(5, Columna) = DataGridView1.Columns(nColumna).HeaderText
                ' libro.Worksheets("Hoja1").Cells(6, Columna).Font.Bold = True

                For nFila As Integer = 0 To RowCount
                    libro.Worksheets("Hoja1").Cells(Fila, Columna) = DataGridView1.Rows(nFila).Cells(nColumna).Value
                    Fila = Fila + 1
                Next
                Columna = Columna + 1
                Fila = 6
            Next

            SaveFileDialog1.DefaultExt = "*.xlsx"
            SaveFileDialog1.FileName = "Libro1"
            SaveFileDialog1.Filter = "Libro de Excel (*.xlsx) | *.xlsx"

            ' GUARDAMOS EL ARCHIVO EXCEL DE LAS VALIDACIONES
            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                libro.SaveAs(SaveFileDialog1.FileName)
                MsgBox("Los Registros Fueron Exportados Satisfactoriamente")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            libro.Saved() = True
            ExcelApp.Quit()
            libro = Nothing
            ExcelApp = Nothing
        End Try
    End Sub
    Private Sub Validador2018_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        BtnExportar.Enabled = False
        TxtRuta.Enabled = False
      
        BtnAbrir.Focus()
    End Sub
    Sub REM01()
        Dim ii, C(91), D(91), E(91), F(91), G(91), H(91), I(91), J(91), K(91), L(91), M(91), N(91), O(91), P(91), Q(91), R(91), S(91), T(91), U(91), V(91), W(91), X(91), Y(91), Z(91), AA(91), AB(91), AC(91), AD(91), AE(91), AF(91), AG(91), AH(91), AI(91), AJ(91), AK(91), AL(91), AM(91), AN(91) As Integer
        xlHoja = xlLibro.Worksheets("A01")
        ' SECCIÓN A: CONTROLES DE SALUD SEXUAL Y REPRODUCTIVA
        For ii = 12 To 27
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
        Next
        'SECCIÓN B: CONTROLES DE SALUD SEGÚN CICLO VITAL
        For ii = 31 To 34
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
        Next
        ' SECCIÓN C: CONTROLES SEGÚN PROBLEMA DE SALUD
        'For ii = 39 To 61
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        ' SECCIÓN D: CONTROLES DE SALUD GRUPAL 
        For ii = 65 To 69
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
        Next
        ' SECCIÓN E: CONTROL DE SALUD INTEGRAL DE ADOLESCENTES PROGRAMA JOVEN SANO (Incluídos en la Sección B)
        'For ii = 73 To 82
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1 ************************************************************************************************************************************************************************************
        Select Case E(12)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [01]", "[REVISAR]", "Control De Salud y Reproductiva, Control Preconcepcional  en edades extremas de 10 a 14 años, celda E12", "[" & E(12) & "]")
                End With
        End Select
        Select Case E(13)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [01]", "[REVISAR]", "Control De Salud y Reproductiva, Control Preconcepcional  en edades extremas de 10 a 14 años, celda E13", "[" & E(13) & "]")
                End With
        End Select
        ' 2 ************************************************************************************************************************************************************************************
        Select Case L(12)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [02]", "[REVISAR]", "Control De Salud y Reproductiva, Control Preconcepcional en edades extremas de 45 a 54 años, celda L12", "[" & L(12) & "]")
                End With
        End Select
        Select Case L(13)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [02]", "[REVISAR]", "Control De Salud y Reproductiva, Control Preconcepcional en edades extremas de 45 a 54 años, celda L13", "[" & L(13) & "]")
                End With
        End Select
        Select Case M(12)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [02]", "[REVISAR]", "Control De Salud y Reproductiva, Control Preconcepcional en edades extremas de 45 a 54 años, celda M12", "[" & M(12) & "]")
                End With
        End Select
        Select Case M(13)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [02]", "[REVISAR]", "Control De Salud y Reproductiva, Control Preconcepcional en edades extremas de 45 a 54 años, celda M13", "[" & M(13) & "]")
                End With
        End Select
        ' 3 ************************************************************************************************************************************************************************************
        Select Case N(14)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [03]", "[REVISAR]", "Control De Salud y Reproductiva, Control Prenatal en mujeres de 55 a 59 años, celda N14", "[" & N(14) & "]")
                End With
        End Select
        Select Case N(15)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [03]", "[REVISAR]", "Control De Salud y Reproductiva, Control Prenatal en mujeres de 55 a 59 años, celda N15", "[" & N(15) & "]")
                End With
        End Select
        ' 4 ************************************************************************************************************************************************************************************
        Select Case N(16)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [04]", "[REVISAR]", "Control De Salud y Reproductiva, Control Post Parto y Post Aborto en mujeres de 55 a 59 años, celda N16", "[" & N(16) & "]")
                End With
        End Select
        Select Case N(17)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [04]", "[REVISAR]", "Control De Salud y Reproductiva, Control Post Parto y Post Aborto en mujeres de 55 a 59 años, celda N17.", "[" & N(17) & "]")
                End With
        End Select
        ' 5 ************************************************************************************************************************************************************************************
        Select Case (N(26) + O(26) + P(26) + Q(26) + R(26) + S(26))
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [05]", "[REVISAR]", "Control De Salud y Reproductiva, Regulación de Fecundidad en mujeres de 55 a 80 años y más, celdas N26 a S26", "[" & (N(26) + O(26) + P(26) + Q(26) + R(26) + S(26)) & "]")
                End With
        End Select
        Select Case (N(27) + O(27) + P(27) + Q(27) + R(27) + S(27))
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [A]", "VAL [05]", "[REVISAR]", "Control De Salud y Reproductiva, Regulación de Fecundidad en mujeres de 55 a 80 años y más, celdas N27 a S27", "[" & (N(27) + O(27) + P(27) + Q(27) + R(27) + S(27)) & "]")
                End With
        End Select
        ' 6 ************************************************************************************************************************************************************************************
        A01(1) = (C(18) + C(19) + C(20) + C(21) + F(31) + F(32) + F(33)) ' VA AL REM 05
        ' 7 ************************************************************************************************************************************************************************************
        Select Case (T(31) + T(32) + T(33))
            Case Is < C(69)
                With Me.DataGridView1.Rows
                    .Add("A01", " [B][D]", "VAL [07]", "[ERROR]", " Controles de Salud según Ciclo Vital, 10 a 14 años, celdas T31 a T33 debe ser mayor o igual a la suma de  Control de Salud Integral de Adolescente, celda C69", "[" & (T(31) + T(32) + T(33)) & "-" & C(69) & "]")
                End With
        End Select
        ' 8 ************************************************************************************************************************************************************************************
        Select Case (U(31) + U(32) + U(33))
            Case Is < F(69)
                With Me.DataGridView1.Rows
                    .Add("A01", " [B][D]", "VAL [08]", "[ERROR]", "Controles de Salud según Ciclo Vital, 10 a 14 años, celdas U31 a U33 debe ser mayor o igual a la suma de  Control de Salud Integral de Adolescente, celda F69", "[" & (U(31) + U(32) + U(33)) & "-" & F(69) & "]")
                End With
        End Select
        ' 9 ************************************************************************************************************************************************************************************
        Select Case F(31)
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A01", " [B]", "VAL [09]", "[REVISAR]", "Controles de Salud según Ciclo Vital, celda F31, no debería ser registrado por el médico salvo excepciones", "[" & F(31) & "]")
                End With
        End Select
        ' 10 ************************************************************************************************************************************************************************************
        A01(2) = (O(31) + O(32)) ' VA AL REM03
        ' 11 ************************************************************************************************************************************************************************************
        A01(3) = (P(31) + P(32)) ' VA AL REM03
        '    ' 12 ************************************************************************************************************************************************************************************
        '    A01(4) = (J(31) + J(32)) ' VA AL REM03
        '    ' 13 ************************************************************************************************************************************************************************************
        '    A01(5) = (N(31) + N(32)) ' VA AL REM03

        ' OTRAS ASIGNACIONES
        A01(6) = (H(31) + H(32)) ' VA AL REM03
        A01(7) = (L(31) + L(32)) ' VA AL REM03
        A01(8) = (T(31) + U(31) + T(32) + U(32) + T(33) + U(33)) ' VA AL REM03


        xlHoja = Nothing
    End Sub 'OK
    Sub REM02()
        Dim ii, B(35), C(35), D(35), E(35), F(35), G(35), H(35), I(35), J(35), K(35), L(35), M(35), N(35), O(35), P(35), Q(35), R(35), S(35), T(35), U(35), V(35), W(35), X(35), Y(35), Z(35), AA(35), AB(35), AC(35), AD(35), AE(35), AF(35) As Integer
        xlHoja = xlLibro.Worksheets("A02")
        ' SECCIÓN A: EMP REALIZADO POR PROFESIONAL
        For ii = 11 To 17
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
        Next
        'SECCIÓN B: EMP SEGÚN RESULTADO DEL ESTADO NUTRICIONAL
        For ii = 21 To 25
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
        Next
        'SECCIÓN C: RESULTADOS DE EMP SEGÚN ESTADO DE SALUD
        'For ii = 29 To 30
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        'Next
        ''SECCIÓN D: RESULTADOS DE EMP SEGÚN ESTADO DE SALUD (EXÁMENES DE LABORATORIO)
        'For ii = 34 To 35
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1*************************************************************************************************************************************************************************************
        Select Case B(11)
            Case Is <> B(21)
                With Me.DataGridView1.Rows
                    .Add("A02", " [A][B]", "VAL [01]", "[ERROR]", "El total de EMP Realizados por Profesional, sección A, celda B11 debe ser igual al total de los EMP según Resultado del Estado Nutricional, sección B, celda B21.", "[" & B(11) & "-" & B(21) & "]")
                End With
        End Select
        '2*************************************************************************************************************************************************************************************
        Select Case C(11)
            Case Is <> C(21)
                With Me.DataGridView1.Rows
                    .Add("A02", " [A][B]", "VAL [02]", "[ERROR]", "El total de EMP Realizados por Profesional a Hombres, sección A, celda C11 debe ser igual al total de los EMP según Resultado del Estado Nutricional a Hombres, sección B, celda C21. ", "[" & C(11) & "-" & C(21) & "]")
                End With
        End Select
        '3*************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123402, 123404, 123406, 123407, 123408, 123410, 123411, 123412, 123413, 123414, 123415, 123416, 123417, 123419, 123420, 123422, 123437, 123423, 123424, 123425, 123426, 123436, 123427, 123428, 123430, 123431, 123432, 123434, 123435, 123709, 123705, 123700, 123701  'POSTA SALUD RURAL CUINCO
            Case Else ' RESTO ESTABLECIMIENTOS
                Select Case B(17)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A02", " [A]", "VAL [03]", "[ERROR]", "EMP realizado por Profesional, celda B17 debe corresponder solo a Postas", "[" & B(17) & "]")
                        End With
                End Select
        End Select
        '4y5************************************************************************************************************************************************************************************
        A02(1) = B(11)

        xlHoja = Nothing
    End Sub 'OK
    Sub REM03()
        Dim ii, B(205), C(205), D(205), E(205), F(205), G(205), H(205), I(205), J(205), K(205), L(205), M(205), N(205), O(205), P(205), Q(205), R(205), S(205), T(205), U(205), V(205), W(205), X(205), Y(205), Z(205), AA(205), AB(205), AC(205), AD(205), AE(205), AF(205), AG(205), AH(205), AI(205), AJ(205), AK(205), AL(205), AM(205) As Integer
        xlHoja = xlLibro.Worksheets("A03")
        'SECCIÓN A1: APLICACIÓN Y RESULTADOS DE PAUTA BREVE
        For ii = 12 To 14
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
        Next
        ' SECCIÓN A2: RESULTADOS DE LA APLICACIÓN DE ESCALA DE EVALUACIÓN DEL DESARROLLO PSICOMOTOR
        For ii = 20 To 38
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
        Next
        'SECCIÓN A3: NIÑOS Y NIÑAS CON REZAGO, DÉFICIT U OTRA VULNERABILIDAD DERIVADOS A ALGUNA MODALIDAD DE ESTIMULACIÓN EN LA PRIMERA EVALUACIÓN
        For ii = 43 To 46
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        'SECCIÓN A.4: RESULTADOS DE LA APLICACIÓN DE PROTOCOLO NEUROSENSORIAL
        For ii = 51 To 54
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
        Next
        'SECCIÓN A.5:  LACTANCIA MATERNA EN MENORES CONTROLADOS
        'For ii = 59 To 63
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        'Next
        'SECCION B: EVALUACIÓN, APLICACIÓN Y RESULTADOS DE ESCALAS EN  LA MUJER
        'SECCIÓN B.1: EVALUACIÓN DEL ESTADO NUTRICIONAL A MUJERES CONTROLADAS AL OCTAVO MES POST PARTO
        'For ii = 68 To 72
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        'SECCIÓN B.2: APLICACIÓN DE ESCALA SEGÚN EVALUACION DE RIESGO PSICOSOCIAL ABREVIADA A GESTANTES
        For ii = 75 To 75
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next
        'SECCIÓN B.3: APLICACIÓN DE ESCALA DE EDIMBURGO A GESTANTES Y MUJERES POST PARTO
        For ii = 78 To 81
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next
        'SECCIÓN C: RESULTADOS DE LA EVALUACIÓN DEL ESTADO NUTRICIONAL DEL ADOLESCENTE CON CONTROL SALUD INTEGRAL
        For ii = 85 To 90
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
        Next
        ' SECCIÓN D: OTRAS EVALUACIONES, APLICACIONES Y RESULTADOS DE ESCALAS EN TODAS LAS EDADES															
        ' SECCIÓN D.1: APLICACIÓN DE INSTRUMENTO E INTERVENCIONES BREVES POR PATRÓN DE CONSUMO ALCOHOL y OTRAS SUSTANCIAS (PROGRAMA DIR EX PROGRAMA VIDA SANA ALCOHOL)																					
        For ii = 96 To 103
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        'SECCIÓN D.2: RESULTADOS DE LA APLICACIÓN DE INSTRUMENTO DE VALORACIÓN DE DESEMPEÑO EN COMUNIDAD (IVADEC-CIF)
        'For ii = 108 To 149
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        'Next
        ' SECCION D.3: RESULTADO APLICACIÓN GHQ12 
        'For ii = 154 To 159
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        'Next
        ''SECCION D.4: RESULTADO DE APLICACIÓN DE CONDICIÓN DE FUNCIONALIDAD AL EGRESO PROGRAMA "MÁS ADULTOS MAYORES AUTOVALENTES"
        'For ii = 164 To 169
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        'Next
        '' SECCION D.5: VARIACION  DE RESULTADO DE APLICACIÓN DEL ÍNDICE DE BARTHEL ENTRE EL INGRESO Y EGRESO HOSPITALARIO														
        'For ii = 174 To 177
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        'Next
        ''SECCIÓN D.6: APLICACIÓN DE ESCALA ZARIT ABREVIADO EN CUIDADORES DE PERSONAS CON DEPENDENCIA SEVERA
        'For ii = 183 To 185
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        'Next
        '' SECCION D.7: APLICACIÓN Y RESULTADOS DE PAUTA DE EVALUACION CON ENFOQUE DE RIESGO ODONTOLOGICO (CERO)
        'For ii = 190 To 192
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        'Next
        '' SECCION E: APLICACIÓN DE PAUTA DETECCIÓN DE FACTORES DE RIESGO PSICOSOCIAL INFANTIL
        'For ii = 197 To 197
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        'Next
        'For ii = 201 To 204
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        Select Case A01(2) '  VALIDACION 10 -REM A01
            Case Is < (L(20) + M(20))
                With Me.DataGridView1.Rows
                    .Add("A01", " [A.2][B]", "VAL [10]", "[ERROR]", "La suma de los Controles de Salud según Ciclo Vital, 18-23 meses, celdas O31:O32, deben ser mayor o igual que Sección A.2 REM 03, Resultados de la Aplicación  de Escala evaluación del desarrollo Psicomotor, 18-23 meses celda L20+M20.", "[" & A01(2) & "-" & (L(20) + M(20)) & "]")
                End With
        End Select
        Select Case A01(3) '  VALIDACION 11 - REM A01
            Case Is < (N(20) + O(20))
                With Me.DataGridView1.Rows
                    .Add("A01", " [A.2][B]", "VAL [11]", "[ERROR]", "La suma de los Controles de Salud según Ciclo Vital, 24-47 meses, celdas P31:P32, deben ser mayor o igual que Sección A.2 REM 03, Resultados de la Aplicación  de Escala evaluación del desarrollo Psicomotor, 24-47 meses celda N20+O20 .", "[" & A01(3) & "-" & (N(20) + O(20)) & "]")
                End With
        End Select

        Select Case A02(1) ' VALIDACION 4 - REM A02
            Case Is <> C(96)
                With Me.DataGridView1.Rows
                    .Add("A02", " [A][D.1]", "VAL [4]", "[ERROR]", " EMP Realizados por Profesional, Total de EMP realizados por profesional, celda B11 debe ser igual al N° de Audit (EMP/EMPAM), Sección D.1 APLICACIÓN DE TAMIZAJE PARA EVALUAR EL NIVEL DE RIESGO DE CONSUMO DE  ALCOHOL, TABACO Y OTRAS DROGAS, celda C96 del REMA03 ", "[" & A02(1) & " - " & C(96) & "]")
                End With
        End Select
        Select Case A02(1) ' VALIDACION 5 - REM A02
            Case Is <> (C(96) + C(98))
                With Me.DataGridView1.Rows
                    .Add("A02", " [A][D.1]", "VAL [5]", "[ERROR]", "El total de EMP Realizados por Profesional, sección A, celda B11 debe ser igual al total de los EMP según Resultado del Estado Nutricional, sección B, celda B21", "[" & A02(1) & " - " & (C(96) + C(98)) & "]")
                End With
        End Select
        '    ''*************************************************************************************************************************************************************************************
        '    ''*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case C(20)
            Case Is <> (C(21) + C(22) + C(23) + C(24) + C(25) + C(26) + C(27) + C(28) + C(29) + C(30) + C(31) + C(32) + C(33))
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.2]", "VAL [01]", "[ERROR]", "El total de Resultados de la Aplicación de escala de evaluación del desarrollo psicomotor, La aplicación del test de desarrollo celda C20 debe ser igual al detalle de evaluaciones del test  de desarrollo Psicomotor celdas C21:C33", "[" & C(20) & "-" & (C(21) + C(22) + C(23) + C(24) + C(25) + C(26) + C(27) + C(28) + C(29) + C(30) + C(31) + C(32) + C(33)) & "]")
                End With
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case C(12)
            Case Is <> (C(13) + C(14))
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.1]", "VAL [02]", "[ERROR]", " La Aplicación y resultados de Pauta Breve, celda C12 debe ser igual a sumatoria de los Resultados de la Aplicación de Pauta Breve, celda C13 y C14.", "[" & C(12) & "-" & (C(13) + C(14)) & "]")
                End With
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case C(51)
            Case Is <> (C(52) + C(53) + C(54))
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.4]", "VAL [03]", "[ERROR]", "Los Resultados de la Aplicación de Protocolo Neurosensorial, celda C51 debe ser igual a sumatoria de los Resultados de la Aplicación de Protocolo Neurosensorial, celda C52:C54", "[" & C(51) & "-" & (C(52) + C(53) + C(54)) & "]")
                End With
        End Select
        '4************************************************************************************************************************************************************************************
        Select Case C(22)
            Case Is <> C(43)
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.2][A.3]", "VAL [04]", "[ERROR]", "El Resultados de la aplicación de escala de evaluación del desarrollo Psicomotor, C22 deben ser igual a Niños y Niñas con rezago, Déficit y otra vulnerabilidad, sección A3, celda C43.", "[" & C(22) & "-" & C(43) & "]")
                End With
        End Select
        '5************************************************************************************************************************************************************************************
        Select Case C(23)
            Case Is <> C(44)
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.2][A.3]", "VAL [05]", "[ERROR]", "El Resultados de la aplicación de escala de evaluación del desarrollo Psicomotor, C23 deben ser igual a Niños y Niñas con rezago, Déficit y otra vulnerabilidad, sección A3,  celda C44.", "[" & C(23) & "-" & C(44) & "]")
                End With
        End Select
        '6************************************************************************************************************************************************************************************
        Select Case C(24)
            Case Is <> C(45)
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.2][A.3]", "VAL [06]", "[ERROR]", "El Resultados de la aplicación de escala de evaluación del desarrollo Psicomotor, C24 deben ser igual a Niños y Niñas con rezago, Déficit y otra vulnerabilidad, sección A3, celda C45.", "[" & C(24) & "-" & C(45) & "]")
                End With
        End Select
        '7************************************************************************************************************************************************************************************
        A03(1) = C(43)
        '8************************************************************************************************************************************************************************************
        A03(2) = C(44)
        '9************************************************************************************************************************************************************************************
        A03(3) = C(45)
        '10************************************************************************************************************************************************************************************
        A03(4) = C(46)
        '11************************************************************************************************************************************************************************************
        A03(5) = C(75)
        '12************************************************************************************************************************************************************************************
        Select Case C(75)
            Case Is < D(75)
                With Me.DataGridView1.Rows
                    .Add("A03", " [B.2]", "VAL [12]", "[ERROR]", "La Aplicación de Escala según evaluación de Riesgo Psicosocial abreviada a gestantes, celda C75, Total de Aplicaciones, debe ser  mayor o igual a Riesgo celda D75", "[" & C(75) & "-" & D(75) & "]")
                End With
        End Select
        '13************************************************************************************************************************************************************************************
        Select Case D(75)
            Case Is < E(75)
                With Me.DataGridView1.Rows
                    .Add("A03", " [B.2]", "VAL [13]", "[ERROR]", " La Aplicación de Escala según evaluación de Riesgo Psicosocial abreviada a gestantes, celda D75, Riesgo, debe ser  mayor o igual a Derivadas a Equipo de cabecera celda E75", "[" & D(75) & "-" & E(75) & "]")
                End With
        End Select
        '14************************************************************************************************************************************************************************************
        Select Case C(78)
            Case Is < C(79)
                With Me.DataGridView1.Rows
                    .Add("A03", " [B.3]", "VAL [14]", "[ERROR]", "La Aplicación de escala de Edimburgo a Gestantes y Mujeres Post Parto, Primera Evaluación, C78 debe ser  mayor o igual a  Reevaluación , celda C79.", "[" & C(78) & "-" & C(79) & "]")
                End With
        End Select
        '15************************************************************************************************************************************************************************************
        Select Case C(80) ' VALIDACION REM A01 
            Case Is > A01(6)
                With Me.DataGridView1.Rows
                    .Add("A03", " [B][B.3]", "VAL [15]", "[ERROR]", "La Aplicación de escala de Edimburgo a Gestantes y Mujeres Post Parto, Evaluación a los 2 Meses, celda C80 debe ser menor igual a los Controles de Salud según Ciclo vital del REM01 sección B, celda H31+ H32.", "[" & C(80) & "-" & A01(6) & "]")
                End With
        End Select
        '16************************************************************************************************************************************************************************************
        Select Case C(81) ' VALIDACION REM A01 
            Case Is > A01(7)
                With Me.DataGridView1.Rows
                    .Add("A03", " [B][B.3]", "VAL [16]", "[ERROR]", " La Aplicación de escala de Edimburgo a Gestantes y Mujeres Post Parto, Evaluación a los 6 Meses, celda C81 debe ser menor igual a los Controles de Salud según Ciclo vital del REM01 sección B, celda L31+L32", "[" & C(81) & "-" & A01(7) & "]")
                End With
        End Select
        '17************************************************************************************************************************************************************************************
        Select Case C(85) ' VALIDACION REM A01 
            Case Is <> A01(8)
                With Me.DataGridView1.Rows
                    .Add("A03", " [C][B]", "VAL [17]", "[ERROR]", "Los Resultados de la Evaluación del Estado Nutricional del Adolescente con Control Salud Integral, celda C85, debe ser igual a Controles de salud según ciclo Vital, REM01, Sección B, suma de celdas T31 a U33", "[" & C(85) & "-" & A01(8) & "]")
                End With
        End Select
        '18************************************************************************************************************************************************************************************
        A03(6) = (C(102) + C(103) + C(104)) ' VA AL REM 27

        xlHoja = Nothing
    End Sub 'OK
    Sub REM04()
        Dim ii, B(126), C(126), D(126), E(126), F(126), G(126), H(126), I(126), J(126), K(126), L(126), M(126), N(126), O(126), P(126), Q(126), R(126), S(126), T(126), U(126), V(126), W(126), X(126), Y(126), Z(126), AA(126), AB(126), AC(126), AD(126), AE(126), AF(126), AG(126), AH(126), AI(126), AJ(126), AK(126), AL(126), AM(126), AN(126), AO(126), AP(126), AQ(126), AR(126), AS1(126) As Integer
        xlHoja = xlLibro.Worksheets("A04")
        ' SECCIÓN A: CONSULTAS MÉDICAS 
        For ii = 12 To 26
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        Next

        ' SECCIÓN B: CONSULTAS DE PROFESIONALES NO MÉDICOS 
        For ii = 31 To 44
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        Next

        'SECCIÓN C: CONSULTAS ANTICONCEPCIÓN DE EMERGENCIA (Incluidas en Sección A y B, respectivamente.)
        For ii = 48 To 49
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
        Next

        ' SECCIÓN D: CONSULTAS EN HORARIO CONTINUADO (Incluidas en las consultas de morbilidad de sección A y B)
        'For ii = 54 To 59
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        'Next

        ' SECCIÓN E: CONSULTAS DE MORBILIDAD SOLICITADAS Y RECHAZADAS DENTRO DE LAS 48 HORAS DE SOLICITADA LA ATENCIÓN
        For ii = 63 To 64
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next

        ' SECCIÓN F: CONSULTA  ABREVIADA
        'For ii = 67 To 68
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next

        ' SECCIÓN G: ATENCIONES DE MEDICINA INDIGENA ASOCIADA AL PROGRAMA ESPECIAL DE SALUD Y PUEBLOS ORIGINARIOS
        'For ii = 72 To 72
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next

        ' SECCIÓN H: INTERVENCIÓN INDIVIDUAL DEL USUARIO EN PROGRAMA VIDA SANA						
        For ii = 77 To 80
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
        Next

        ' SECCIÓN I: SERVICIOS FARMACEÚTICOS
        'For ii = 85 To 89
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        'For ii = 91 To 93
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next

        ' SECCIÓN J: DESPACHO DE RECETAS DE PACIENTES AMBULATORIOS					
        For ii = 98 To 101
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
        Next

        ' SECCIÓN K: RONDAS POR TIPO Y PROFESIONAL
        'For ii = 105 To 107
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        'Next

        ' SECCION L: CLASIFICACION DE CONSULTA NUTRICIONAL POR GRUPO DE EDAD (Incluidas en Seccion B)
        For ii = 112 To 114
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
        Next

        ' SECCIÓN M: CONSULTA DE LACTANCIA MATERNA EN MENORES CONTROLADOS
        For ii = 120 To 126
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
        Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case B(32)
            Case Is <> 0
                If B(32) <= B(36) Then
                    With Me.DataGridView1.Rows
                        .Add("A04", "[B]", "VAL [01]", "[ERROR]", "Consultas de Profesionales no Médicos, Matrona/ón (Morb. Ginecológica) de la sección B, celda B32 tiene dato, entonces debe ser mayor que las Otras consultas por la Matrona/ón, celda B36.", "[" & B(32) & "-" & B(36) & "]")
                    End With
                End If
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case B(49)
            Case Is <> 0
                If B(49) <= B(36) Then
                    With Me.DataGridView1.Rows
                        .Add("A04", " [C][B]", "VAL [02]", "[ERROR]", "Consultas Anticoncepción de Emergencia por Matrona/ón, Celda B49 tiene dato, entonces debe ser menor o igual a las Consultas de Profesionales No médicos Matrona/ón de la sección B, celda B36.", "[" & B(49) & "-" & B(36) & "]")
                    End With
                End If
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case (B(37) + B(38) + B(39))
            Case Is <> (B(112) + B(113) + B(114))
                With Me.DataGridView1.Rows
                    .Add("A04", "[B][L]", "VAL [03]", "[ERROR]", "Consultas de Profesionales no médicos, Consultas por Nutricionistas, Celda B37:B39, debe ser igual a Sección L, Clasificación de consulta Nutricional por grupo de edad, Celdas B112:B114", "[" & (B(37) + B(38) + B(39)) & "-" & (B(112) + B(113) + B(114)) & "]")
                End With
        End Select
        '4************************************************************************************************************************************************************************************
        Select Case (C(48) + C(49))
            Case Is > 0
                With Me.DataGridView1.Rows
                    .Add("A04", "[C]", "VAL [04]", "[ERROR]", " Las Consultas Anticoncepción de Emergencia, celdas C48:C49 si es mayor a cero en el grupo de 10 a 14 años.", "[" & (C(48) + C(49)) & "]")
                End With
        End Select
        '5************************************************************************************************************************************************************************************
        Select Case (G(48) + G(49))
            Case Is > 0
                With Me.DataGridView1.Rows
                    .Add("A04", "[C]", "VAL [05]", "[ERROR]", "Las Consultas Anticoncepción de Emergencia, celdas G48:G49 si es mayor a cero en el grupo de 10 a 14 años", "[" & (G(48) + G(49)) & "]")
                End With
        End Select
        '6************************************************************************************************************************************************************************************
        Select Case (B(63) + B(64))
            Case Is < (E(12) + F(12) + G(12) + H(12))
                With Me.DataGridView1.Rows
                    .Add("A04", "[C]", "VAL [06]", "[ERROR]", "Consultas de morbilidad solicitadas y rechazadas dentro de las 48 horas de solicitada la atención, Menores de 5 Años, Total de Atención Solicitada, celda B63:B64 debe ser ≥ que sumatoria  Sección A, Total Consultas Médicas hasta 4 Años, Celda E12:H12 ", "[" & (B(63) + B(64)) & "-" & (E(12) + F(12) + G(12) + H(12)) & "]")
                End With
        End Select
        '7************************************************************************************************************************************************************************************
        Select Case (D(63) + D(64))
            Case Is < (AG(12) + AH(12) + AI(12) + AJ(12) + AK(12) + AL(12) + AM(12) + AN(12))
                With Me.DataGridView1.Rows
                    .Add("A04", "[C]", "VAL [07]", "[ERROR]", " Consultas de morbilidad solicitadas y rechazadas dentro de las 48 horas de solicitada la atención , 65 Años y Más , Total de Atención Solicitada, Celda D63:D64 debe ser ≥ que sumatoria Sección A, Total Consultas Médicas  65 Años hasta 80 y más, Celda AG12:AN12", "[" & (D(63) + D(64)) & "-" & (AG(12) + AH(12) + AI(12) + AJ(12) + AK(12) + AL(12) + AM(12) + AN(12)) & "]")
                End With
        End Select
        '8************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123422, 123423, 123424, 123426, 123427, 123428, 123307, 123309, 123304, 123311, 123312, 123300, 123310, 123305, 123301, 123306, 123302
            Case Else
                Select Case (C(77) + C(78) + C(79) + C(80))
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A04", "[H]", "VAL [08]", "[ERROR]", "Intervención individual del usuario en programa Vida Sana, Celdas C77:C80 Solo deben registrar los Establecimientos: DSM Pto. Octay Y  CESFAM (Purranque, P. Araya, San Pablo, Entre Lagos, Puaucho, Bahía Mansa, Jauregui, V Centenario)", "[" & (C(77) + C(78) + C(79) + C(80)) & "]")
                        End With
                End Select
        End Select
        '9************************************************************************************************************************************************************************************
        Select Case B(101)
            Case Is < F(101)
                With Me.DataGridView1.Rows
                    .Add("A04", "[J]", "VAL [09]", "[REVISAR]", "Despacho de recetas de pacientes ambulatorios, celda B101 los registros de recetas despachadas deben ser al menos igual a las recetas despachadas con oportunidad celda F101", "[" & B(101) & "-" & F(101) & "]")
                End With
        End Select
        '10************************************************************************************************************************************************************************************
        Select Case (C(120) + C(121) + C(122))
            Case Is < (C(123) + C(124) + C(125) + C(126))
                With Me.DataGridView1.Rows
                    .Add("A04", "[M]", "VAL [10]", "[ERROR]", "Consultas de lactancia en menores controlados por Nutricionista, Celda C120:C122, debe ser igual a las consultas por lactancia celdas C123:C126", "[" & (C(120) + C(121) + C(122)) & "-" & (C(123) + C(124) + C(125) + C(126)) & "]")
                End With
        End Select

        'Validacion Recetas
        Select Case C(98)
            Case Is >= B(98)
                With Me.DataGridView1.Rows
                    .Add("A04", "[J]", "RECETA[01]", "[ERROR]", "Despacho Recetas Pacientes Ambulatorios - Crónica, Recetas Despachada parcial (Celda C98), Debe ser menor a recetas despachadas total (Celda B98)", "[" & C(98) & "-" & B(98) & "]")
                End With
        End Select

        Select Case D(98)
            Case Is <= B(98)
                With Me.DataGridView1.Rows
                    .Add("A04", "[J]", "RECETA[02]", "[ERROR]", "Despacho Recetas Pacientes Ambulatorios - Crónica, Prescripciones emitidas (Celda D98) debe ser mayor a recetas despachadas total (Celda B98)", "[" & D(98) & "-" & B(98) & "]")
                End With
        End Select

        Select Case E(98)
            Case Is >= D(98)
                With Me.DataGridView1.Rows
                    .Add("A04", "[J]", "RECETA[03]", "[ERROR]", "Despacho Recetas Pacientes Ambulatorios - Crónica, Prescripciones Rechazadas (Celda E98) debe ser menor a prescripciones emitidas (Celda D98)", "[" & E(98) & "-" & D(98) & "]")
                End With
        End Select

        Select Case F(98)
            Case Is <> (B(98) - C(98))
                With Me.DataGridView1.Rows
                    .Add("A04", "[J]", "RECETA[04]", "[ERROR]", "Despacho Recetas Pacientes Ambulatorios - Crónica, Recetas Despachadas con Oportunidad (Celda F98), debe ser igual a recetas con despacho total menos despacho Parcial (Celdas B98-C98)", "[" & F(98) & "-" & (B(98) - C(98)) & "]")
                End With
        End Select



        A04(1) = B(24)

        xlHoja = Nothing
    End Sub 'OK
    Sub REM05()
        Dim ii, B(304), C(304), D(304), E(304), F(304), G(304), H(304), I(304), J(304), K(304), L(304), M(304), N(304), O(304), P(304), Q(304), R(304), S(304), T(304), U(304), V(304), W(304), X(304), Y(304), Z(304), AA(304), AB(304), AC(304), AD(304), AE(304), AF(304), AG(304), AH(304), AI(304), AJ(304), AK(304), AL(304), AM(304), AN(304), AO(304), AP(304), AQ(304), AR(304), AS1(304), AT(304), AU(304), AV(304), AW(304), AX(304), AY(304), AZ(304) As Integer
        xlHoja = xlLibro.Worksheets("A05")
        ' SECCIÓN A: INGRESOS DE GESTANTES A PROGRAMA PRENATAL
        For ii = 11 To 15
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
        Next

        ' SECCIÓN B: INGRESO DE GESTANTES CON PATOLOGÍA DE ALTO RIESGO OBSTÉTRICO A LA UNIDAD DE ARO  (Nivel Secundario)
        'For ii = 19 To 29
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next

        ' SECCIÓN C: INGRESOS A PROGRAMA DE REGULACION DE FERTILIDAD Y SALUD SEXUAL.
        'For ii = 33 To 52
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        'Next

        ' SECCION D: INGRESOS A PROGRAMA CONTROL DE CLIMATERIO 
        'For ii = 55 To 55
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next

        'SECCIÓN E: INGRESOS  A CONTROL DE SALUD DE RECIÉN NACIDOS
        For ii = 58 To 58
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next

        ' SECCIÓN F:  INGRESOS Y EGRESOS A SALA DE ESTIMULACIÓN SERVICIO ITINERANTE Y ATENCIÓN DOMICILIARIA 
        'For ii = 63 To 66
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        'Next

        ' SECCIÓN F.1: REINGRESOS Y EGRESOS POR SEGUNDA VEZ A MODALIDAD DE ESTIMULACIÓN EN EL CENTRO DE SALUD
        For ii = 71 To 74
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
        Next

        ' SECCION G: INGRESOS DE NIÑOS Y NIÑAS CON NECESIDADES ESPECIALES DE BAJA COMPLEJIDAD A CONTROL DE SALUD INFANTIL EN APS
        'For ii = 79 To 79
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        'Next

        ' SECCIÓN H: INGRESOS AL PSCV
        For ii = 84 To 89
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
        Next

        ' SECCIÓN I: EGRESOS  DEL PSCV
        'For ii = 94 To 99
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        'Next

        ' SECCIÓN J: INGRESOS Y EGRESOS AL PROGRAMA DE PACIENTES CON DEPENDENCIA LEVE, MODERADA Y SEVERA
        For ii = 104 To 108
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        Next

        ' SECCIÓN K: INGRESOS AL PROGRAMA DEL ADULTO MAYOR SEGÚN CONDICIÓN DE FUNCIONALIDAD Y DEPENDENCIA
        For ii = 112 To 121
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
        Next

        ' SECCIÓN L: EGRESOS DEL PROGRAMA DEL ADULTO MAYOR SEGÚN CONDICIÓN DE FUNCIONALIDAD Y DEPENDENCIA
        'For ii = 125 To 134
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        'Next

        ' SECCIÓN M: INGRESOS Y EGRESOS DEL  PROGRAMA MÁS ADULTOS MAYORES AUTOVALENTES
        'For ii = 138 To 145
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        'Next

        ' SECCIÓN N: INGRESOS AL PROGRAMA DE SALUD  MENTAL EN APS /ESPECIALIDAD
        For ii = 151 To 189
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("As" & ii & "").Value
        Next

        ' SECCIÓN O: EGRESOS DEL PROGRAMA DE SALUD  MENTAL POR ALTAS CLÍNICAS EN APS /ESPECIALIDAD
        'For ii = 193 To 231
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("As" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        'Next

        ' SECCIÓN P:  INGRESOS Y EGRESOS AL COMPONENTE ALCOHOL Y DROGA EN APS/ESPECIALIDAD
        'For ii = 235 To 238
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        'Next

        ' SECCIÓN Q:  PROGRAMA DE REHABILITACIÓN (PERSONAS CON TRASTORNOS PSIQUIÁTRICOS)
        'For ii = 243 To 246
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        'Next

        ' SECCIÓN R: INGRESOS Y EGRESOS A PROGRAMA INFECCIÓN POR TRANSMISIÓN SEXUAL (Uso de establecimientos que realizan atención de ITS)
        'For ii = 251 To 269
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("As" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        'Next

        ' SECCIÓN S: INGRESOS Y EGRESOS DEL PROGRAMA DE VIH/SIDA (Uso exclusivo Centros de Atención VIH-SIDA)
        For ii = 274 To 281
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("As" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
            AV(ii) = xlHoja.Range("AV" & ii & "").Value
            AW(ii) = xlHoja.Range("AW" & ii & "").Value
            AX(ii) = xlHoja.Range("AX" & ii & "").Value
        Next

        ' SECCIÓN T: INGRESOS Y EGRESOS POR COMERCIO SEXUAL (Uso exclusivo de Unidades Control Comercio Sexual)
        For ii = 286 To 288
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
        Next

        ' SECCIÓN U. INGRESOS Y EGRESOS PROGRAMA DE ACOMPAÑAMIENTO PSICOSOCIAL EN ATENCIÓN PRIMARIA
        'For ii = 293 To 294
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        'Next
        ' SECCION V 
        'For ii = 299 To 304   
        'C(ii) = xlHoja.Range("C" & ii & "").Value
        'D(ii) = xlHoja.Range("D" & ii & "").Value
        'E(ii) = xlHoja.Range("E" & ii & "").Value
        'F(ii) = xlHoja.Range("F" & ii & "").Value
        'G(ii) = xlHoja.Range("G" & ii & "").Value
        'H(ii) = xlHoja.Range("H" & ii & "").Value
        'I(ii) = xlHoja.Range("I" & ii & "").Value
        'J(ii) = xlHoja.Range("J" & ii & "").Value
        'K(ii) = xlHoja.Range("K" & ii & "").Value
        'L(ii) = xlHoja.Range("L" & ii & "").Value
        'M(ii) = xlHoja.Range("M" & ii & "").Value
        'N(ii) = xlHoja.Range("N" & ii & "").Value
        'O(ii) = xlHoja.Range("O" & ii & "").Value
        'P(ii) = xlHoja.Range("P" & ii & "").Value
        'Q(ii) = xlHoja.Range("Q" & ii & "").Value
        'R(ii) = xlHoja.Range("R" & ii & "").Value
        'S(ii) = xlHoja.Range("S" & ii & "").Value
        'T(ii) = xlHoja.Range("T" & ii & "").Value
        'U(ii) = xlHoja.Range("U" & ii & "").Value
        'V(ii) = xlHoja.Range("V" & ii & "").Value
        'W(ii) = xlHoja.Range("W" & ii & "").Value
        'X(ii) = xlHoja.Range("X" & ii & "").Value
        'Y(ii) = xlHoja.Range("Y" & ii & "").Value
        'Z(ii) = xlHoja.Range("Z" & ii & "").Value
        'AA(ii) = xlHoja.Range("AA" & ii & "").Value
        'AB(ii) = xlHoja.Range("AB" & ii & "").Value
        'AC(ii) = xlHoja.Range("AC" & ii & "").Value
        'AD(ii) = xlHoja.Range("AD" & ii & "").Value
        'AE(ii) = xlHoja.Range("AE" & ii & "").Value
        'AF(ii) = xlHoja.Range("AF" & ii & "").Value
        'AG(ii) = xlHoja.Range("AG" & ii & "").Value
        'AH(ii) = xlHoja.Range("AH" & ii & "").Value
        'AI(ii) = xlHoja.Range("AI" & ii & "").Value
        'AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        'Next
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        ' VALIDACIONES REM 01 *********************************************************************************
        Select Case A01(1) ' VALIDACION 06
            Case Is <> C(58)
                With Me.DataGridView1.Rows
                    .Add("A01", " [A-B][E]", "VAL [06]", "[ERROR]", " Controles de salud sexual y Reproductiva, la Suma de Puérpera con recién Nacidos 10 Días hasta 28 Días, celdas C18 a C21 y Controles de salud sexual y Reproductiva, Menor de 1 Mes, celdas  F31 a F33  deben ser igual a la Sección E REM05, Ingresos a control de salud recién Nacidos, Total menores de 28 Días, celda C58", "[" & A01(1) & " - " & C(58) & "]")
                End With
        End Select
        ' VALIDACIONES REM 03 *********************************************************************************
        Select Case A03(1) ' VALIDACION 07
            Case Is < C(63)
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.3][F]", "VAL [07]", "[ERROR]", "La Derivación niños y niñas con Rezago, celda C43 debe ser mayor o igual a Los Ingresos y Egresos a Sala de Estimulación en el Centro de Salud del REM05, sección F, celda C63", "[" & A03(1) & " - " & C(63) & "]")
                End With
        End Select
        Select Case A03(2) ' VALIDACION 08
            Case Is < C(64)
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.3][F]", "VAL [08]", "[ERROR]", " La Derivación niños y niñas con Rezago, celda C44 debe ser mayor o igual a Los Ingresos y Egresos a Sala de Estimulación en el Centro de Salud del REM05, sección F, celda C64.", "[" & A03(2) & " - " & C(64) & "]")
                End With
        End Select
        Select Case A03(3) ' VALIDACION 09
            Case Is < C(65)
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.3][F]", "VAL [09]", "[ERROR]", " La Derivación niños y niñas con Rezago, celda C45 debe ser mayor o igual a Los Ingresos y Egresos a Sala de Estimulación en el Centro de Salud del REM05, sección F, celda C65.", "[" & A03(3) & " - " & C(65) & "]")
                End With
        End Select
        Select Case A03(4) ' VALIDACION 10
            Case Is < C(66)
                With Me.DataGridView1.Rows
                    .Add("A03", " [A.3][F]", "VAL [10]", "[ERROR]", "La Derivación niños y niñas con Rezago, celda C46 debe ser mayor o igual a Los Ingresos y Egresos a Sala de Estimulación en el Centro de Salud del REM05, sección F, celda C66.", "[" & A03(4) & " - " & C(66) & "]")
                End With
        End Select
        Select Case A03(5) 'VALIDACION 11
            Case Is > C(11)
                With Me.DataGridView1.Rows
                    .Add("A03", " [B.2][A]", "VAL [11]", "[ERROR]", "La Aplicación de Escala según evaluación de Riesgo Psicosocial abreviada a gestantes, celda C75 debe ser menor o igual a Los Ingresos de Gestantes a Programa Prenatal, Total Gestantes Ingresadas en el REM05, sección A, celda C11", "[" & A03(5) & " - " & C(11) & "]")
                End With
        End Select
        '    '*************************************************************************************************************************************************************************************
        '    '********************************************************************** VALIDACIONES *************************************************************************************************
        '    '*************************************************************************************************************************************************************************************
        '    '*************************************************************************************************************************************************************************************
        '    '1************************************************************************************************************************************************************************************
        Select Case (K(11) + L(11) + M(11))
            Case Is > 0
                With Me.DataGridView1.Rows
                    .Add("A05", " [A]", "VAL [01]", "[REVISAR]", "Ingresos de Gestantes a Programa Prenatal, celda K11: M11 entre edades extremas, 45 a 55 años, deben ser corroboradas por profesional a cargo", "[" & (K(11) + L(11) + M(11)) & "]")
                End With
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case (K(12) + L(12) + M(12))
            Case Is > 0
                With Me.DataGridView1.Rows
                    .Add("A05", " [A]", "VAL [02]", "[REVISAR]", "Ingresos de Gestantes a Programa Prenatal, celda K12:M12 entre edades extremas, 45 y 55 años, deben ser corroboradas por profesional a cargo", "[" & (K(12) + L(12) + M(12)) & "]")
                End With
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case (K(13) + L(13) + M(13))
            Case Is > 0
                With Me.DataGridView1.Rows
                    .Add("A05", " [A]", "VAL [03]", "[REVISAR]", "Ingresos de Gestantes a Programa Prenatal, celda K13:M13 entre edades extremas, 45 y 55 años, deben ser corroboradas por profesional a cargo", "[" & (K(13) + L(13) + M(13)) & "]")
                End With
        End Select
        '4************************************************************************************************************************************************************************************
        Select Case (K(14) + L(14) + M(14))
            Case Is > 0
                With Me.DataGridView1.Rows
                    .Add("A05", " [A]", "VAL [04]", "[REVISAR]", "Ingresos de Gestantes a Programa Prenatal, celda K14:M14 entre edades extremas, 45 y 55 años, deben ser corroboradas por profesional a cargo", "[" & (K(14) + L(14) + M(14)) & "]")
                End With
        End Select
        '5************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123404, 123425, 123422, 123423, 123424, 123426, 123427, 123428, 123410, 123434, 123432, 123436, 123435, 123430, 123431, 123402, 123411, 123412, 123413, 123414, 123415, 123416, 123417, 123419, 123420, 123406, 123407, 123408
                Select Case (C(73) + C(74) + C(75) + C(76))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A05", " [F.1]", "VAL [05]", "[ERROR]", "Reingresos y Egresos por segunda vez a Sala de Estimulación en el Centro de Salud, celdas C71 a C74 corresponde el registro a todos los establecimientos con excepción de Postas", "[" & (C(73) + C(74) + C(75) + C(76)) & "]")
                        End With
                End Select
            Case Else ' RESTO DE ESTABLECIMIENTOS

        End Select
        '6************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123302, 123310, 123307, 123309, 123103, 123304, 123305, 123104, 123105, 123311, 123312
            Case Else ' RESTO DE ESTABLECIMIENTOS
                Select Case (C(71) + C(72) + C(73) + C(74))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A05", " [F.1]", "VAL [06]", "[ERROR]", "Reingresos y Egresos por segunda vez a Sala de Estimulación en el Centro de Salud, celdas C71 a C74 deben corresponder solo a Establecimientos con Sala de Estimulación", "[" & (C(73) + C(74) + C(75) + C(76)) & "]")
                        End With
                End Select
        End Select
        '    '7************************************************************************************************************************************************************************************
        Select Case C(84)
            Case Is > (C(85) + C(86) + C(87) + C(88) + C(89))
                With Me.DataGridView1.Rows
                    .Add("A05", " [H]", "VAL [07]", "[ERROR]", "Ingresos al PSCV, celda C84 debe ser menor o igual a la suma del desglose del Programa de Salud Cardiovascular, celdas C85 a 89.", "[" & C(84) & "-" & (C(85) + C(86) + C(87) + C(88) + C(89)) & "]")
                End With
        End Select
        '    '8**********************************************************************************************************************************************************************************
        Select Case (Z(104) + AA(104) + AB(104) + AC(104) + AD(104) + AE(104) + AF(104) + AG(104) + Z(105) + AA(105) + AB(105) + AC(105) + AD(105) + AE(105) + AF(105) + AG(105) + Z(106) + AA(106) + AB(106) + AC(106) + AD(106) + AE(106) + AF(106) + AG(106) + Z(107) + AA(107) + AB(107) + AC(107) + AD(107) + AE(107) + AF(107) + AG(107) + Z(108) + AA(108) + AB(108) + AC(108) + AD(108) + AE(108) + AF(108) + AG(108))
            Case Is < C(120)
                With Me.DataGridView1.Rows
                    .Add("A05", " [J][K]", "VAL [08]", "[ERROR]", "Ingresos y Egresos al programa de pacientes con dependencia Leve, Moderada y Severa, desde 65 Años y más , Celda Z104:AG108 debe ser mayor a Sección K, Ingreso al programa A.M según condición de dependencia, subtotal desagregación Barthel, Celda C120.", "[" & (Z(104) + AA(104) + AB(104) + AC(104) + AD(104) + AE(104) + AF(104) + AG(104) + Z(105) + AA(105) + AB(105) + AC(105) + AD(105) + AE(105) + AF(105) + AG(105) + Z(106) + AA(106) + AB(106) + AC(106) + AD(106) + AE(106) + AF(106) + AG(106) + Z(107) + AA(107) + AB(107) + AC(107) + AD(107) + AE(107) + AF(107) + AG(107) + Z(108) + AA(108) + AB(108) + AC(108) + AD(108) + AE(108) + AF(108) + AG(108)) & "-" & C(122) & "]")
                End With
        End Select

        '    '9**********************************************************************************************************************************************************************************
        Select Case C(151)
            Case Is < C(158)
                With Me.DataGridView1.Rows
                    .Add("A05", " [N]", "VAL [09]", "[ERROR]", "Ingreso al Programa de Salud Mental en APS/Especialidad, Ingresos al programa , celda C151 deben ser Mayor o Igual  al Numero personas que posee uno o más trastornos Mentales, Celda C158.", "[" & C(151) & "-" & C(158) & "]")
                End With
        End Select
        '10**********************************************************************************************************************************************************************************
        Select Case C(158)
            Case Is > (C(159) + C(160) + C(161) + C(162) + C(163) + C(164) + C(165) + C(166) + C(167) + C(168) + C(169) + C(170) + C(171) + C(172) + C(173) + C(174) + C(175) + C(176) + C(177) + C(178) + C(179) + C(180) + C(181) + C(182) + C(183) + C(184) + C(185) + C(186) + C(187) + C(188) + C(189))
                With Me.DataGridView1.Rows
                    .Add("A05", " [N]", "VAL [10]", "[ERROR]", " Ingreso al Programa de Salud Mental en APS/Especialidad, Personas con Diagnostico de Trastornos Mentales celda C158, deben ser menor o igual a  los Diagnósticos de Trastornos Mentales de la sección N, en las Celdas C159 a C189", "[" & D(158) & "-" & (C(159) + C(160) + C(161) + C(162) + C(163) + C(164) + C(165) + C(166) + C(167) + C(168) + C(169) + C(170) + C(171) + C(172) + C(173) + C(174) + C(175) + C(176) + C(177) + C(178) + C(179) + C(180) + C(181) + C(182) + C(183) + C(184) + C(185) + C(186) + C(187) + C(188) + C(189)) & "]")
                End With
        End Select
        '   '11***********************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 200209, 200445, 200477, 200248, 123010 ' DISCRIMINAR HBSJO - COSAM RAHUE - COSAM ORIENTE - AYEKAN- UNIDAD DE MEMORIA - DSSO
            Case Else
                Select Case C(151)
                    Case Is > A04(1)
                        With Me.DataGridView1.Rows
                            .Add("A05", " [N][A]", "VAL [11]", "[ERROR]", "Ingreso al Programa de Salud Mental en APS/Especialidad, Ingresos al programa, Celda C151, de existir registros en  consultas de Salud Mental REM A04, Sección A: Consultas Medica, celda B24, estas deben ser mayor al ingreso de programa de Salud Mental del REM A05", "[" & C(151) & "-" & A04(1) & "]")
                        End With
                End Select
        End Select
        '12***********************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 200248 ' CDR DE ADULTO MAYOR CON DEMENCIA
                Select Case (C(180) + C(181) + C(182) + C(183) + C(184) + C(185) + C(186) + C(187) + C(188) + C(189))
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A05", " [N]", "VAL [12]", "[REVISAR]", "Ingreso al Programa de Salud Mental en APS/Especialidad, Personas con diagnósticos de trastornos mentales, celda C180:C189, debe realizar registro el establecimiento CDR DE ADULTO MAYOR CON DEMENCIA ", "[" & (C(180) + C(181) + C(182) + C(183) + C(184) + C(185) + C(186) + C(187) + C(188) + C(189)) & "]")
                        End With
                End Select
        End Select
        '    '13***********************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123300, 123301, 123302, 123303, 123306, 123310, 123404, 123425, 123700, 123701
                Select Case (C(243) + C(244) + C(245) + C(246))
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A05", " [Q]", "VAL [13]", "[ERROR]", "El Programa de Rehabilitación, celdas C243 a E246 corresponde al HBSJO y DSSO.", "[" & (C(243) + C(244) + C(245) + C(246)) & "]")
                        End With
                End Select
        End Select
        '    '14***********************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100
            Case Else
                Select Case (C(274) + C(275) + C(276) + C(277) + C(278) + C(279) + C(280) + C(281))
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A05", " [S]", "VAL [14]", "[REVISAR]", "El Ingreso y Egreso a Programa de VIH/SIDA, celdas C274 a E281 corresponde solo al HBSJO.", "[" & (C(276) + C(277) + C(278) + C(279) + C(280) + C(281) + C(282) + C(283)) & "]")
                        End With
                End Select
        End Select
        '15***********************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100
            Case Else
                Select Case (C(286) + C(287) + C(280))
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A05", " [T]", "VAL [15]", "[REVISAR]", " El Ingreso y Egreso por Comercio Sexual, celdas C286 a E288 corresponde solo al HBSJO", "[" & (C(286) + C(287) + C(280)) & "]")
                        End With
                End Select
        End Select


        xlHoja = Nothing
    End Sub 'OK 
    Sub REM06()
        Dim ii, B(151), C(151), D(151), E(151), F(151), G(151), H(151), I(151), J(151), K(151), L(151), M(151), N(151), O(151), P(151), Q(151), R(151), S(151), T(151), U(151), V(151), W(151), X(151), Y(151), Z(151), AA(151), AB(151), AC(151), AD(151), AE(151), AF(151), AG(151), AH(151), AI(151), AJ(151), AK(151), AL(151), AM(151), AN(151), AO(151), AP(151), AQ(151), AR(151), AS1(151), AT(151), AU(151), AV(151), AW(151), AX(151), AY(151), AZ(151) As Integer
        xlHoja = xlLibro.Worksheets("A06")
        'SECCIÓN A. ATENCIÓN PRIMARIA
        'SECCIÓN A.1: CONTROLES DE ATENCION PRIMARIA / ESPECIALIDADES
        For ii = 13 To 27
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
        Next
        ' SECCIÓN A.2: CONSULTORÍAS DE SALUD MENTAL
        'For ii = 31 To 31
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        'Next
        '' SECCIÓN B. ATENCIÓN DE ESPECIALIDADES
        '' SECCIÓN B.1: ACTIVIDADES GRUPALES (NÚMERO DE SESIONES)
        'For ii = 36 To 39
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN B.2:  PROGRAMA DE REHABILITACIÓN (PERSONAS CON TRASTORNOS PSIQUIÁTRICOS)
        'For ii = 44 To 45
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        '' SECCIÓN B.3: ACTIVIDADES DE PSIQUIATRÍA FORENSE PARA PERSONAS EN CONFLICTO CON LA JUSTICIA (En lo Penal, Civil, Familiar, etc.)												
        'For ii = 50 To 71
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        '' SECCIÓN B.4: DISPOSITIVOS DE SALUD MENTAL
        'For ii = 75 To 78
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        'Next
        '' SECCIÓN C. ACTIVIDADES COMUNES EN AMBOS TIPOS DE ATENCIÓN
        '' SECCIÓN C.1: ACTIVIDADES DE COORDINACION SECTORIAL, INTERSECTORIAL Y COMUNITARIA
        'For ii = 83 To 89
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        '' SECCIÓN C.2: INFORMES A TRIBUNALES
        'For ii = 93 To 98
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        '' SECCIÓN D. PLANES Y EVALUACIONES PROGRAMA DE ACOMPAÑAMIENTO PSICOSOCIAL EN ATENCIÓN PRIMARIA
        'For ii = 103 To 104
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        'Next
        ''SECCIÓN E: PLANES DE CUIDADO INTEGRAL (PCI)
        'For ii = 109 To 109
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        'Next
        '' SECCIÓN F: PERSONAS CON EVALUACION Y CONFIRMACION DIAGNOSTICA EN APS
        'For ii = 113 To 113
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN G: REGISTRO DE ATENCIONES PROFESIONALES PLAN DE DEMENCIA EN ATENCIÓN PRIMARIA
        'For ii = 118 To 123
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        'Next
        '' SECCIÓN H. EVALUACIONES PROGRAMA PLAN NACIONAL DE DEMENCIA
        'For ii = 128 To 130
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        'Next
        'For ii = 134 To 134
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case CodigoEstablec '
            Case 123103, 123102
                Select Case (C(13) + C(14) + C(15) + C(16) + C(17) + C(18) + C(19) + C(20) + C(21) + C(22) + C(23) + C(24) + C(25) + C(26) + C(27))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A06", " [A.1]", "VAL [01]", "[ERROR]", "Controles de atención Primaria/ Especialidades, el registro de controles por Profesionales, Celdas C13 a E27,  No lo deben registrar los establecimientos HRN - HPU", "[" & (C(13) + C(14) + C(15) + C(16) + C(17) + C(18) + C(19) + C(20) + C(21) + C(22) + C(23) + C(24) + C(25) + C(26) + C(27)) & "]")
                        End With
                End Select
            Case Else ' RESTO DE ESTABLECIMIENTOS

        End Select


        xlHoja = Nothing
    End Sub 'OK
    Sub REM07()
        Dim ii, B(117), C(117), D(117), E(117), F(117), G(117), H(117), I(117), J(117), K(117), L(117), M(117), N(117), O(117), P(117), Q(117), R(117), S(117), T(117), U(117), V(117), W(117), X(117), Y(117), Z(117), AA(117), AB(117), AC(117), AD(117), AE(117), AF(117), AG(117), AH(117), AI(117), AJ(117), AK(117), AL(117), AM(117), AN(117), AO(117), AP(117), AQ(117), AR(117), AS1(117), AT(117), AU(117), AV(117), AW(117), AX(117), AY(117), AZ(117) As Integer
        xlHoja = xlLibro.Worksheets("A07")
        ' SECCIÓN A : CONSULTAS MÉDICAS
        For ii = 11 To 71
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
            AV(ii) = xlHoja.Range("AV" & ii & "").Value
            AW(ii) = xlHoja.Range("AW" & ii & "").Value
        Next
        ' SECCIÓN B: ATENCIONES MEDICAS POR PROGRAMAS Y POLICLINICOS ESPECIALISTAS ACREDITADOS
        'For ii = 75 To 90
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        'Next
        '' SECCIÓN C: CONSULTAS Y CONTROLES POR OTROS PROFESIONALES EN ESPECIALIDAD (NIVEL SECUNDARIO)
        'For ii = 95 To 106
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        'Next
        '' SECCIÓN D: CONSULTAS INFECCIÓN TRANSMISIÓN SEXUAL (ITS) Y CONTROLES DE SALUD SEXUAL EN EL NIVEL SECUNDARIO 
        'For ii = 111 To 121
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case (AR(58) + AS1(58))
            Case Is <> 0
                With Me.DataGridView1.Rows
                    .Add("A07", " [A]", "VAL [01]", "[REVISAR]", "Consultas Médicas por Oftalmología , Total Interconsultas generadas en APS para derivación especialidad, Celdas AR58:AS58 con registros. Se recuerda que solo se debe registrar producción Propia - GES en esta sección. Prestaciones pertenecientes a Programa de Resolutividad - UAPO se registran en REM29", "[" & (AR(58) + AS1(58)) & "]")
                End With
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123300, 123301, 123302, 123303, 123701 ' JAUREGUI, M.LOPETEGUI, OVEJERIA, RAHUE ALTO, CECOF MANUEL RODRIGUEZ
            Case Else 'RESTO DE ESTABLECIMIENTOS
                Select Case B(11)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A07", " [A]", "VAL [02]", "[REVISAR]", " Consultas Médicas de Especialidad, Solo deben registrar en Pediatría B11 los establecimientos : CESFAM: M.LOPETEGUI, P.JAUREGUI, OVEJERIA Y RAHUE ALTO; CESCOF M. RODRUIGUEZ y SAFU siempre que cuenten con médicos especialistas", "[" & (B(11)) & "]")
                        End With
                End Select
        End Select
        Select Case CodigoEstablec
            Case 123100, 123300, 123301, 123302, 123303, 123701 ' JAUREGUI, M.LOPETEGUI, OVEJERIA, RAHUE ALTO, CECOF MANUEL RODRIGUEZ
            Case Else 'RESTO DE ESTABLECIMIENTOS
                Select Case B(65)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A07", " [A]", "VAL [02]", "[REVISAR]", " Consultas Médicas de Especialidad, Solo deben registrar en Medicina Familiar B65 los establecimientos : CESFAM: M.LOPETEGUI, P.JAUREGUI, OVEJERIA Y RAHUE ALTO; CESCOF M. RODRUIGUEZ y SAFU siempre que cuenten con médicos especialistas", "[" & B(65) & "]")
                        End With
                End Select
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123011, 200477, 200209
            Case Else 'RESTO DE ESTABLECIMIENTOS
                Select Case B(42)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A07", " [A]", "VAL [03]", "[ERROR]", "Consultas Médicas de Especialidad, solo deben registrar en Psiquiatría B42 los establecimientos: HBSJO, PRAIS, COSAM RAHUE Y CDR A MAYOR", "[" & (B(42)) & "]")
                        End With
                End Select
        End Select
        '4************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123301, 123309, 123304, 123100
            Case Else 'RESTO DE ESTABLECIMIENTOS
                Select Case B(51)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A07", " [A]", "VAL [04]", "[ERROR]", "Consultas Médicas de Especialidad, solo deben registrar en UAPO B51 los establecimientos: CESFAM: M. LOPETEGUI, PURRANQUE, PABLO ARAYA y ENTRE LAGOS", "[" & (B(51)) & "]")
                        End With
                End Select
        End Select
        '5************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123101
            Case Else
                Select Case B(31)
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A07", " [A]", "VAL [05]", "[ERROR]", "Consultas Médicas de Especialidad, solo deben registrar los establecimientos HBSJO y HPU en las celdas Dermatología B31", "[" & B(31) & "]")
                        End With
                End Select
                Select Case B(46)
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A07", " [A]", "VAL [05]", "[ERROR]", "Consultas Médicas de Especialidad, solo deben registrar los establecimientos HBSJO y HPU en las celdas Máxilo Facial B46", "[" & B(46) & "]")
                        End With
                End Select
                Select Case B(51)
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A07", " [A]", "VAL [05]", "[ERROR]", "Consultas Médicas de Especialidad, solo deben registrar los establecimientos HBSJO y HPU en las celdas Cirugía vascular periférica B51", "[" & B(51) & "]")
                        End With
                End Select
        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM08()
        Dim ii, B(194), C(194), D(194), E(194), F(194), G(194), H(194), I(194), J(194), K(194), L(194), M(194), N(194), O(194), P(194), Q(194), R(194), S(194), T(194), U(194), V(194), W(194), X(194), Y(194), Z(194), AA(194), AB(194), AC(194), AD(194), AE(194), AF(194), AG(194), AH(194), AI(194), AJ(194), AK(194), AL(194), AM(194), AN(194), AO(194), AP(194), AQ(194), AR(194), AS1(194), AT(194), AU(194), AV(194), AW(194), AX(194), AY(194), AZ(194) As Integer
        xlHoja = xlLibro.Worksheets("A08")
        ' SECCIÓN A: ATENCIONES REALIZADAS EN UNIDADES DE URGENCIA DE LA RED
        ' SECCIÓN A.1: ATENCIONES REALIZADAS EN UNIDADES DE EMERGENCIA HOSPITALARIA DE ALTA Y MEDIANA COMPLEJIDAD (UEH)
        For ii = 12 To 14
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        Next
        ' SECCIÓN A.2: ATENCIONES DE URGENCIA REALIZADAS EN SAPU Y SAR												
        'For ii = 19 To 22
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        '' SECCIÓN A.3: ATENCIONES DE URGENCIA REALIZADAS EN ESTABLECIMIENTOS DE BAJA COMPLEJIDAD
        'For ii = 27 To 32
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        '' SECCIÓN A.4: ATENCIONES DE URGENCIA REALIZADAS EN ESTABLECIMIENTOS  ATENCIÓN PRIMARIA NO SAPU
        'For ii = 37 To 42
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        '' SECCIÓN A.5:  CONSULTAS EN SISTEMA DE ATENCIÓN DE URGENCIA EN CENTROS DE SALUD RURAL (SUR) Y POSTAS RURALES
        'For ii = 47 To 52
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        ' SECCIÓN B: CATEGORIZACIÓN DE PACIENTES EN UNIDADES DE URGENCIA, PREVIA A LA ATENCIÓN MÉDICA, EN HOSPITALES DE ALTA Y MEDIANA COMPLEJIDAD
        For ii = 57 To 63
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
        Next
        ' SECCIÓN C: ATENCIONES REALIZADAS POR MÉDICOS ESPECIALISTAS EN LAS UNIDADES DE URGENCIA HOSPITALARIA
        'For ii = 66 To 85
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN D: PACIENTES CON INDICACIÓN DE HOSPITALIZACIÓN EN ESPERA DE CAMAS EN UEH
        'For ii = 90 To 97
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        '' SECCIÓN E: PACIENTES CON INDICACIÓN DE OBSERVACIÓN EN SAR
        'For ii = 100 To 102
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        '' SECCIÓN F: PACIENTES FALLECIDOS EN ESPERA DE ATENCIÓN MÉDICA
        'For ii = 107 To 109
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        'Next
        '' SECCIÓN G: ATENCIONES MÉDICAS ASOCIADAS A  VIOLENCIA
        'For ii = 113 To 114
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        'Next
        ' SECCIÓN H: ATENCIONES  POR ANTICONCEPCIÓN DE EMERGENCIA 
        For ii = 118 To 119
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
        Next
        '' SECCIÓN I: MOTIVOS DE ATENCIÓN POR EMERGENCIA OBSTÉTRICA AL SERVICIO DE  URGENCIA  (Establecimientos Alta y Mediana Complejidad).
        'For ii = 123 To 135
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCIÓN J: LLAMADOS DE URGENCIA A CENTRO REGULADOR
        'For ii = 138 To 138
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN K: INTERVENCIONES PRE HOSPITALARIAS (SAMU)
        'For ii = 142 To 143
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        'Next
        '' SECCIÓN L: TRASLADOS PRIMARIOS A UNIDADES DE URGENCIA (Desde el lugar del evento a unidad de Emergencia)
        'For ii = 146 To 151
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN M: TRASLADO SECUNDARIO (Desde un Establecimiento a Otro)
        'For ii = 155 To 160
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        'Next
        '' SECCION N: CLASIFICACION DE LAS INTERVENCIONES POR GRANDES GRUPOS DE DIAGNOSTICOS (SAMU)
        'For ii = 165 To 168
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        'Next
        '' SECCION O: ATENCIONES  EN URGENCIA POR VIOLENCIA SEXUAL  
        'For ii = 173 To 178
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        'Next
        ''SECCIÓN P: ATENCIONES MÉDICAS POR VIOLENCIA SEXUAL
        'For ii = 182 To 183
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        ''SECCIÓN Q: ATENCIONES DE URGENCIA ASOCIADAS A LESIONES AUTOINFLIGIDAS
        'For ii = 188 To 188
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        'Next
        ''SECCIÓN R: ATENCIONES POR MORDEDURA EN SERVICIO DE  URGENCIA DE LA RED
        'For ii = 193 To 198
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123101 ' ESTABLECIMIENTOS DE ALTA COMPLEJIDAD Y MEDIANA HBO Y PURRANQUE
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (B(12) + B(13) + B(14))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A08", " [A.1]", "VAL [01]", "[ERROR]", "Las Atenciones Realizadas en UEH de Hosp. de Alta Complejidad, celdas C12 a AA14 corresponde solo a HBSJO y HPU", "[" & (B(12) + B(13) + B(14)) & "]")
                        End With
                End Select
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case B(63)
            Case Is <> B(12)
                With Me.DataGridView1.Rows
                    .Add("A08", " [A.1][B]", "VAL [02]", "[ERROR]", " Las Atenciones Realizadas en UEH de Hosp. de Alta Complejidad, celdas B12 debe ser igual a la sección B, Categorizaciones de Pacientes, Previa a la Atención Medica, celda B63", "[" & B(63) & "-" & B(12) & "]")
                End With
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123100, 123101 ' HBO
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case B(63)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A08", " [B]", "VAL [03]", "[ERROR]", "Categorizaciones de Pacientes, Previa a la Atención Medica, celdas C57 a S63 corresponde solo a HBSJO y HPU", "[" & B(63) & "]")
                        End With
                End Select
        End Select
        '4************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123311, 123304, 123305, 123300, 123301, 123303, 123101, 123102, 123103, 123104, 123105, 123100 ' SUR SAPU HOSPITALES DE URGENCIA
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (C(118) + C(119))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A08", " [H]", "VAL [04]", "[ERROR]", "Atenciones por Anticoncepción de Emergencia, Celdas C118 y C119 corresponde solo SUR, SAPU, Hospitales (Urgencia).", "[" & (C(118) + C(119)) & "]")
                        End With
                End Select
        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM09()
        Dim ii, B(309), C(309), D(309), E(309), F(309), G(309), H(309), I(309), J(309), K(309), L(309), M(309), N(309), O(309), P(309), Q(309), R(309), S(309), T(309), U(309), V(309), W(309), X(309), Y(309), Z(309), AA(309), AB(309), AC(309), AD(309), AE(309), AF(309), AG(309), AH(309), AI(309), AJ(309), AK(309), AL(309), AM(309), AN(309), AO(309), AP(309), AQ(309), AR(309), AS1(309), AT(309), AU(309), AV(309), AW(309), AX(309), AY(309), AZ(309) As Integer
        xlHoja = xlLibro.Worksheets("A09")
        ' SECCIÓN A: CONSULTAS Y CONTROLES ODONTOLÓGICOS
        'For ii = 12 To 15
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        'Next
        '' SECCIÓN B: OTRAS ACTIVIDADES DE ODONTOLOGÍA GENERAL
        'For ii = 19 To 38
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        'Next
        ' SECCIÓN C : INGRESOS Y EGRESOS  EN ESTABLECIMIENTOS APS
        For ii = 41 To 58
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("Ah" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
        Next
        ' SECCIÓN D : INTERCONSULTAS GENERADAS EN ESTABLECIMIENTOS APS
        'For ii = 63 To 73
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN E: CONSULTAS ODONTOLÓGICAS  EN HORARIO CONTINUADO (incluidas en  Secciones A y B. . Excluye extensiones horarias de Sección G )
        'For ii = 77 To 78
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN F:  ACTIVIDADES EN ATENCIÓN DE ESPECIALIDADES
        'For ii = 82 To 111
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        'Next
        '' SECCIÓN F.1:  ACTIVIDADES DE APOYO EN ATENCIÓN DE ESPECIALIDADES
        'For ii = 116 To 122
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        'Next
        ' SECCIÓN G: PROGRAMAS ESPECIALES Y GES (Actividades incluidas en Sección B y F)
        For ii = 149 To 181
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
        Next
        ' SECCIÓN G.1: PROGRAMA SEMBRANDO SONRISAS
        'For ii = 160 To 164
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        'Next
        '' SECCIÓN H: SEDACIÓN Y ANESTESIA
        'For ii = 169 To 172
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        'Next
        '' SECCIÓN I:  CONSULTAS, INGRESOS Y EGRESOS A TRATAMIENTOS EN ESTABLECIMIENTOS DE NIVEL SECUNDARIO Y TERCIARIO
        'For ii = 177 To 230
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        'Next
        ''SECCIÓN J: ACTIVIDADES EFECTUADAS POR TÉCNICO PARAMÉDICO DENTAL Y/O HIGIENISTAS DENTALES
        'For ii = 235 To 240
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        'Next
        '' SECCIÓN K.- GESTIÓN DE AGENDA (UNIDADES DENTALES MOVILES)
        'For ii = 246 To 249
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN L: CONSULTORÍAS DE ESPECIALISTAS OTORGADAS
        'For ii = 252 To 258
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case (U(41) + V(41) + W(41) + X(41) + Y(41) + Z(41) + AA(41) + AB(41) + AC(41) + AD(41))
            Case Is < (U(54) + V(54) + W(54) + X(54) + Y(54) + Z(54) + AA(54) + AB(54) + AC(54) + AD(54))
                With Me.DataGridView1.Rows
                    .Add("A09", " [C]", "VAL [01]", "[ERROR]", "Ingresos y Egresos en Establecimientos APS, Ingresos a tratamientos Odontología General,  U41 a AD41 debe ser Menor o igual a Total índice CEOD O COPD en Pacientes Ingresados Odontología General , Celdas U54  a AD54 para mayores de 12 años.", "[" & (U(41) + V(41) + W(41) + X(41) + Y(41) + Z(41) + AA(41) + AB(41) + AC(41) + AD(41)) & "-" & (U(54) + V(54) + W(54) + X(54) + Y(54) + Z(54) + AA(54) + AB(54) + AC(54) + AD(54)) & "]")
                End With
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case (AA(46) + AB(46))
            Case Is < (AF(46))
                With Me.DataGridView1.Rows
                    .Add("A09", " [C]", "VAL [02]", "[REVISAR]", "Ingresos y Egresos en Establecimientos APS, ingreso de 60 Años celda AF46,  debe estar incluidos en los ingresos de altas Odontológicas en su desagregación de 20 a 60 Años, celdas AA46:AB46", "[" & (AA(46) + AB(46)) & " - " & (AF(46)) & "]")
                End With
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case CodigoEstablec '
            Case 123300, 123301, 123302, 123303, 123306, 123310, 123304
            Case Else
                Select Case (D(149) + D(150) + D(151) + D(152) + D(153) + D(154) + D(155) + D(156) + D(157) + D(158) + D(159) + D(160) + D(161) + D(162) + D(163) + D(164) + D(165) + D(166) + D(167) + D(168) + D(169) + D(170) + D(171) + D(172) + D(173) + D(174) + D(175) + D(176) + D(177) + D(178) + D(179) + D(180) + D(181))
                    Case Is <> 0 ' RESTO DE ESTABLECIMIENTOS
                        With Me.DataGridView1.Rows
                            .Add("A09", " [G]", "VAL [03]", "[ERROR]", "Programas Especiales y GES, celdas D149 a D181 solo deben registrar los Cesfam de Osorno, cesfam Entre lagos", "[" & (D(149) + D(150) + D(151) + D(152) + D(153) + D(154) + D(155) + D(156) + D(157) + D(158) + D(159) + D(160) + D(161) + D(162) + D(163) + D(164) + D(165) + D(166) + D(167) + D(168) + D(169) + D(170) + D(171) + D(172) + D(173) + D(174) + D(175) + D(176) + D(177) + D(178) + D(179) + D(180) + D(181)) & "]")
                        End With
                End Select
        End Select
        '4************************************************************************************************************************************************************************************

        Select Case CodigoEstablec
            Case 123103, 123104, 123105
                Select Case (D(146) + D(147) + D(148) + D(149) + D(150) + D(151) + D(152) + D(153) + D(154) + D(155) + D(156) + D(157) + D(158) + D(159) + D(160) + D(161) + D(162) + D(163) + D(164) + D(165) + D(166) + D(167) + D(168) + D(169) + D(170) + D(171) + D(172) + D(173) + D(174) + D(175) + D(176) + D(177) + D(178) + D(179) + D(180) + D(181))
                    Case Is <> 0 ' RESTO DE ESTABLECIMIENTOS
                        With Me.DataGridView1.Rows
                            .Add("A09", " [G]", "VAL [04]", "[ERROR]", " Programas Especiales y GES, Programas: Odontológicos adultos de 60 años, Estrategia Estudiantes 4º Medio y Mejoramiento del Acceso Estrategia Adulto Mayor, celdas D172 a D81 solo deben registrar los Hospitales con APS (HPO- HSMJ- HQUI)", "[" & (D(146) + D(147) + D(148) + D(149) + D(150) + D(151) + D(152) + D(153) + D(154) + D(155) + D(156) + D(157) + D(158) + D(159) + D(160) + D(161) + D(162) + D(163) + D(164) + D(165) + D(166) + D(167) + D(168) + D(169) + D(170) + D(171) + D(172) + D(173) + D(174) + D(175) + D(176) + D(177) + D(178) + D(179) + D(180) + D(181)) & "]")
                        End With
                End Select
        End Select


        xlHoja = Nothing
    End Sub 'OK
    Sub REM11()
        Dim ii, B(264), C(264), D(264), E(264), F(264), G(264), H(264), I(264), J(264), K(264), L(264), M(264), N(264), O(264), P(264), Q(264), R(264), S(264), T(264), U(264), V(264), W(264), X(264), Y(264), Z(264), AA(264), AB(264), AC(264), AD(264), AE(264), AF(264), AG(264), AH(264), AI(264), AJ(264), AK(264), AL(264), AM(264), AN(264), AO(264), AP(264), AQ(264), AR(264), AS1(264), AT(264), AU(264), AV(264), AW(264), AX(264), AY(264), AZ(264) As Integer
        xlHoja = xlLibro.Worksheets("A11")
        ' SECCION A: EXÁMENES DE SÍFILIS
        ' SECCIÓN A.1: EXÁMENES DE SÍFILIS POR GRUPO DE USUARIOS (Uso exclusivo de establecimientos con Laboratorio que procesan)
        For ii = 12 To 29
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
            AV(ii) = xlHoja.Range("AV" & ii & "").Value
            AW(ii) = xlHoja.Range("AW" & ii & "").Value
        Next
        ' SECCIÓN A.2: EXÁMENES DE SÍFILIS POR GRUPO DE USUARIOS (Uso exclusivo de establecimientos que Compran Servicio)
        'For ii = 34 To 51
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        '    AW(ii) = xlHoja.Range("AW" & ii & "").Value
        'Next
        '' SECCIÓN A.3: EXAMEN RPR POR GRUPO DE USUARIOS (Uso exclusivo de establecimientos con Laboratorio que procesan)
        'For ii = 56 To 73
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        '    AW(ii) = xlHoja.Range("AW" & ii & "").Value
        'Next
        '' SECCIÓN A.4: EXAMEN RPR POR GRUPO DE USUARIOS (Uso exclusivo de establecimientos que Compran Servicio)
        'For ii = 78 To 95
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        '    AW(ii) = xlHoja.Range("AW" & ii & "").Value
        'Next
        '' SECCIÓN A.5: EXAMEN MHA-TP POR GRUPO DE USUARIOS (Uso exclusivo de establecimientos con Laboratorio que procesan)
        'For ii = 100 To 117
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        '    AW(ii) = xlHoja.Range("AW" & ii & "").Value
        'Next
        '' SECCIÓN A.6: EXAMEN MHA-TP POR GRUPO DE USUARIOS (Uso exclusivo de establecimientos que Compran Servicio)
        'For ii = 122 To 139
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        '    AW(ii) = xlHoja.Range("AW" & ii & "").Value
        'Next
        ' SECCIÓN B.1: EXÁMENES SEGÚN GRUPOS DE USUARIOS POR CONDICIÓN DE HEPATITIS B, HEPATITIS C, CHAGAS, HTLV 1 Y SIFILIS (Uso exclusivo de establecimientos con Laboratorio que procesan)																
        For ii = 143 To 147
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
        Next
        ' SECCIÓN B.2: EXÁMENES SEGÚN GRUPOS DE USUARIOS POR CONDICIÓN DE HEPATITIS B, HEPATITIS C, CHAGAS, HTLV 1 Y SIFILIS (Uso exclusivo de establecimientos que Compran Servicio)																	
        For ii = 151 To 155
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
        Next
        ' SECCIÓN C.1: EXÁMENES  DE  VIH POR GRUPOS DE USUARIOS (Uso exclusivo de establecimientos con Laboratorio que procesan)
        For ii = 160 To 182
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
            AV(ii) = xlHoja.Range("AV" & ii & "").Value
            AW(ii) = xlHoja.Range("AW" & ii & "").Value
            AX(ii) = xlHoja.Range("AX" & ii & "").Value
        Next
        'SECCIÓN C.2: EXÁMENES  DE  VIH POR GRUPOS DE USUARIOS (Uso exclusivo de establecimientos que Compran Servicio)
        'For ii = 187 To 209
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        '    AW(ii) = xlHoja.Range("AW" & ii & "").Value
        '    AX(ii) = xlHoja.Range("AX" & ii & "").Value
        'Next
        '' SECCIÓN D: DETECCIÓN ENFERMEDAD DE CHAGAS EN GESTANTES Y RECIÉN NACIDOS SEGÚN RESULTADO DE EXÁMENES DE LABORATORIO
        'For ii = 212 To 215
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        '' SECCIÓN E: EXÁMENES DE GONORREA POR GRUPOS DE USUARIOS 
        'For ii = 220 To 226
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        'Next
        ''SECCIÓN F: EXÁMENES DE CHLAMYIDIA TRACHOMATIS POR GRUPOS DE USUARIOS 
        'For ii = 231 To 237
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        'Next
        ''
        'For ii = 242 To 264
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        '    AU(ii) = xlHoja.Range("AU" & ii & "").Value
        '    AV(ii) = xlHoja.Range("AV" & ii & "").Value
        '    AW(ii) = xlHoja.Range("AW" & ii & "").Value
        '    AX(ii) = xlHoja.Range("AX" & ii & "").Value
        'Next


        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123101 ' HBO Y HPU
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (B(12) + B(13) + B(14) + B(15) + B(16) + B(17) + B(18) + B(19) + B(20) + B(21) + B(22) + B(23) + B(24) + B(25) + B(26) + B(27) + B(28) + B(29) + C(12) + C(13) + C(14) + C(15) + C(16) + C(17) + C(18) + C(19) + C(20) + C(21) + C(22) + C(23) + C(24) + C(25) + C(26) + C(27) + C(28) + C(29))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A11", " [A.1]", "VAL [01]", "[ERROR]", "Examen VDRL Por Grupo de Usuarios, celdas B12 a C29, esta sección solo le corresponde solo a HBSJO  y HPU", "[" & (B(12) + B(13) + B(14) + B(15) + B(16) + B(17) + B(18) + B(19) + B(20) + B(21) + B(22) + B(23) + B(24) + B(25) + B(26) + B(27) + B(28) + B(29) + C(12) + C(13) + C(14) + C(15) + C(16) + C(17) + C(18) + C(19) + C(20) + C(21) + C(22) + C(23) + C(24) + C(25) + C(26) + C(27) + C(28) + C(29)) & "]")
                        End With
                End Select
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO 
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (C(143) + C(144) + C(145) + C(146) + C(147) + D(143) + D(144) + D(145) + D(146) + D(147) + E(143) + E(144) + E(145) + E(146) + E(147) + F(143) + F(144) + F(145) + F(146) + F(147) + G(143) + G(144) + G(145) + G(146) + G(147) + H(143) + H(144) + H(145) + H(146) + H(147) + I(143) + I(144) + I(145) + I(146) + I(147) + J(143) + J(144) + J(145) + J(146) + J(147) + K(143) + K(144) + K(145) + K(146) + K(147) + L(143) + L(144) + L(145) + L(146) + L(147) + M(143) + M(144) + M(145) + M(146) + M(147) + N(143) + N(144) + N(145) + N(146) + N(147) + O(143) + O(144) + O(145) + O(146) + O(147) + P(143) + P(144) + P(145) + P(146) + P(147))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A11", " [B.1]", "VAL [02]", "[ERROR]", "Exámenes Según grupos de usuarios por condición de Hepatitis B, C, Chagas, HTLV1 y Sífilis, Uso Exclusivo Lab. Que Procesan, celdas C143 a P147 le corresponde solo a HBSJO", "[" & (C(143) + C(144) + C(145) + C(146) + C(147) + D(143) + D(144) + D(145) + D(146) + D(147) + E(143) + E(144) + E(145) + E(146) + E(147) + F(143) + F(144) + F(145) + F(146) + F(147) + G(143) + G(144) + G(145) + G(146) + G(147) + H(143) + H(144) + H(145) + H(146) + H(147) + I(143) + I(144) + I(145) + I(146) + I(147) + J(143) + J(144) + J(145) + J(146) + J(147) + K(143) + K(144) + K(145) + K(146) + K(147) + L(143) + L(144) + L(145) + L(146) + L(147) + M(143) + M(144) + M(145) + M(146) + M(147) + N(143) + N(144) + N(145) + N(146) + N(147) + O(143) + O(144) + O(145) + O(146) + O(147) + P(143) + P(144) + P(145) + P(146) + P(147)) & "]")
                        End With
                End Select
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO 
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (C(151) + C(152) + C(153) + C(154) + C(155) + D(151) + D(152) + D(153) + D(154) + D(155) + E(151) + E(152) + E(153) + E(154) + E(155) + F(151) + F(152) + F(153) + F(154) + F(155) + G(151) + G(152) + G(153) + G(154) + G(155) + H(151) + H(152) + H(153) + H(154) + H(155) + I(151) + I(152) + I(153) + I(154) + I(155) + J(151) + J(152) + J(153) + J(154) + J(155) + K(151) + K(152) + K(153) + K(154) + K(155) + L(151) + L(152) + L(153) + L(154) + L(155) + M(151) + M(152) + M(153) + M(154) + M(155) + N(151) + N(152) + N(153) + N(154) + N(155) + O(151) + O(152) + O(153) + O(154) + O(155) + P(151) + P(152) + P(153) + P(154) + P(155))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A11", " [B.2]", "VAL [03]", "[ERROR]", "", "[" & (C(151) + C(152) + C(153) + C(154) + C(155) + D(151) + D(152) + D(153) + D(154) + D(155) + E(151) + E(152) + E(153) + E(154) + E(155) + F(151) + F(152) + F(153) + F(154) + F(155) + G(151) + G(152) + G(153) + G(154) + G(155) + H(151) + H(152) + H(153) + H(154) + H(155) + I(151) + I(152) + I(153) + I(154) + I(155) + J(151) + J(152) + J(153) + J(154) + J(155) + K(151) + K(152) + K(153) + K(154) + K(155) + L(151) + L(152) + L(153) + L(154) + L(155) + M(151) + M(152) + M(153) + M(154) + M(155) + N(151) + N(152) + N(153) + N(154) + N(155) + O(151) + O(152) + O(153) + O(154) + O(155) + P(151) + P(152) + P(153) + P(154) + P(155)) & "]")
                        End With
                End Select
        End Select
        '4************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO 
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (C(160) + D(160) + C(161) + D(161) + C(162) + D(162) + C(163) + D(163) + C(164) + D(164) + C(165) + D(165) + C(166) + D(166) + C(167) + D(167) + C(168) + D(168) + C(169) + D(169) + C(170) + D(170) + C(171) + D(171) + C(172) + D(172) + C(173) + D(173) + C(174) + D(174) + C(175) + D(175) + C(176) + D(176) + C(177) + D(177) + C(178) + D(178) + C(179) + D(179) + C(180) + D(180) + C(181) + D(181) + C(182) + D(182))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A11", " [C.1]", "VAL [04]", "[ERROR]", "Exámenes de VIH por Grupos de Usuarios, celda C160 a D182, esta sección solo le corresponde solo a HBSJO.", "[" & (C(160) + D(160) + C(161) + D(161) + C(162) + D(162) + C(163) + D(163) + C(164) + D(164) + C(165) + D(165) + C(166) + D(166) + C(167) + D(167) + C(168) + D(168) + C(169) + D(169) + C(170) + D(170) + C(171) + D(171) + C(172) + D(172) + C(173) + D(173) + C(174) + D(174) + C(175) + D(175) + C(176) + D(176) + C(177) + D(177) + C(178) + D(178) + C(179) + D(179) + C(180) + D(180) + C(181) + D(181) + C(182) + D(182)) & "]")
                        End With
                End Select
        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM19a()
        Dim ii, B(162), C(162), D(162), E(162), F(162), G(162), H(162), I(162), J(162), K(162), L(162), M(162), N(162), O(162), P(162), Q(162), R(162), S(162), T(162), U(162), V(162), W(162), X(162), Y(162), Z(162), AA(162), AB(162), AC(162), AD(162), AE(162), AF(162), AG(162), AH(162), AI(162), AJ(162), AK(162), AL(162), AM(162), AN(162), AO(162), AP(162), AQ(162), AR(162), AS1(162), AT(162), AU(162), AV(162), AW(162), AX(162), AY(162), AZ(162) As Integer
        xlHoja = xlLibro.Worksheets("A19a")
        ' SECCIÓN A: CONSEJERÍAS	
        ' SECCIÓN A.1: CONSEJERÍAS INDIVIDUALES
        'For ii = 14 To 88
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        'Next
        '' SECCIÓN A.2: CONSEJERÍAS INDIVIDUALES POR VIH (NO INCLUIDAS EN LA SECCION A.1)
        'For ii = 93 To 104
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        'Next
        '' SECCIÓN A.3: CONSEJERÍAS FAMILIARES
        'For ii = 107 To 115
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        ' SECCIÓN B: ACTIVIDADES DE PROMOCIÓN
        ' SECCIÓN B.1: ACTIVIDADES DE PROMOCIÓN SEGÚN ESTRATEGIAS Y CONDICIONANTES ABORDADAS Y NÚMERO DE PARTICIPANTES
        For ii = 128 To 143
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
        Next
        ' SECCIÓN B.2: TALLERES GRUPALES DE VIDA SANA SEGÚN TIPO, POR ESPACIOS DE ACCIÓN
        'For ii = 127 To 142
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        'Next
        '' S
        'For ii = 145 To 148
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next
        ''SECCIÓN B.3: ACTIVIDADES DE GESTIÓN SEGÚN TIPO, POR ESPACIOS DE ACCIÓN
        'For ii = 151 To 156
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next
        ''SECCIÓN B.4: TALLERES GRUPALES SEGÚN TEMATICA Y NUMERO DE PARTICIPANTES EN PROGRAMA ESPACIOS AMIGABLES 
        'For ii = 159 To 161
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case (C(128) + C(129) + C(130) + C(131) + C(132) + C(133) + C(134) + C(135) + C(136) + C(137) + C(138) + C(139) + C(140) + C(141) + C(142) + C(143))
            Case Is <> 0
                If (L(128) + L(129) + L(130) + L(131) + L(132) + L(133) + L(134) + L(135) + L(136) + L(137) + L(138) + L(139) + L(140) + L(141) + L(142) + L(143)) = 0 Then
                    With Me.DataGridView1.Rows
                        .Add("A19A ", " [B.1]", "VAL [01]", "[ERROR]", "Actividades de Promoción según estrategias y Condicionantes Abordadas y  Nº de Participantes, Si Existen registros TOTAL ACTIVIDADES, celdas C128:C143 se debe registrar el TOTAL PARTICIPANTES  L128:L143.", "[" & (C(128) + C(129) + C(130) + C(131) + C(132) + C(133) + C(134) + C(135) + C(136) + C(137) + C(138) + C(139) + C(140) + C(141) + C(142) + C(143)) & " - " & (L(128) + L(129) + L(130) + L(131) + L(132) + L(133) + L(134) + L(135) + L(136) + L(137) + L(138) + L(139) + L(140) + L(141) + L(142) + L(143)) & "]")
                    End With
                End If
        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM19b()
        Dim ii, B(49), C(49), D(49), E(49), F(49), G(49), H(49), I(49), J(49), K(49), L(49), M(49), N(49), O(49), P(49), Q(49), R(49), S(49), T(49), U(49), V(49), W(49), X(49), Y(49), Z(49), AA(49), AB(49), AC(49), AD(49), AE(49), AF(49), AG(49), AH(49), AI(49), AJ(49), AK(49), AL(49), AM(49), AN(49), AO(49), AP(49), AQ(49), AR(49), AS1(49), AT(49), AU(49), AV(49), AW(49), AX(49), AY(49), AZ(49) As Integer
        xlHoja = xlLibro.Worksheets("A19b")
        ' SECCIÓN A: ATENCIÓN OFICINAS DE INFORMACIONES (SISTEMA INTEGRAL DE ATENCIÓN A USUARIOS)
        For ii = 11 To 29
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
        Next
        ' SECCIÓN B: ACTIVIDADES POR ESTRATEGIA/LÍNEA DE ACCIÓN O ESPACIO / INSTANCIA DE PARTICIPACIÓN SOCIAL
        'For ii = 33 To 45
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        'Next
        '' SECCIÓN C: REUNIONES DE ADULTO MAYOR
        'For ii = 48 To 49
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case B(11)
            Case Is <> 0
                Select Case (E(11) + F(11) + G(11) + H(11) + I(11))
                    Case Is = 0
                        With Me.DataGridView1.Rows
                            .Add("A19B", " [A]", "VAL [01]", "[REVISAR]", "Atención Oficinas de Informaciones, Si existen reclamos en, celda B11 debiera haber una respuestas a ese reclamos (dentro/fuera de los plazos legales o pendientes), celdas E11 a I11.", "[" & B(11) & "-" & (E(11) + F(11) + G(11) + H(11) + I(11)) & "]")
                        End With
                End Select
        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM21()
        Dim ii, B(85), C(85), D(85), E(85), F(85), G(85), H(85), I(85), J(85), K(85), L(85), M(85), N(85), O(85), P(85), Q(85), R(85), S(85), T(85), U(85), V(85), W(85), X(85), Y(85), Z(85), AA(85), AB(85), AC(85), AD(85), AE(85), AF(85), AG(85), AH(85), AI(85), AJ(85), AK(85), AL(85), AM(85), AN(85), AO(85), AP(85), AQ(85), AR(85), AS1(85), AT(85), AU(85), AV(85), AW(85), AX(85), AY(85), AZ(85) As Integer
        xlHoja = xlLibro.Worksheets("A21")
        ' SECCIÓN A:  CAPACIDAD INSTALADA Y UTILIZACIÓN DE LOS QUIRÓFANOS
        For ii = 12 To 16
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        '' SECCIÓN B:  PROCEDIMIENTOS COMPLEJOS AMBULATORIOS
        'For ii = 19 To 23
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        ''SECCIÓN C:  HOSPITALIZACIÓN DOMICILIARIA
        ''SECCIÓN C.1:  PERSONAS ATENDIDAS EN EL PROGRAMA
        'For ii = 27 To 32
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        ''SECCIÓN C.2:  VISITAS REALIZADAS 
        'For ii = 35 To 35
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN D:  HOSPITAL AMIGO
        '' SECCIÓN D.1: ACOMPAÑAMIENTO A HOSPITALIZADOS
        'For ii = 39 To 42
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN D.2: INFORMACIÓN A FAMILIARES DE PACIENTES EGRESADOS 
        'For ii = 46 To 47
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN E: APOYO PSICOSOCIAL EN NIÑOS (AS) HOSPITALIZADOS				
        'For ii = 50 To 54
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        'Next
        '' SECCIÓN F: GESTIÓN DE PROCESOS DE PACIENTES QUIRÚRGICOS CON CIRUGÍA ELECTIVA						
        'For ii = 59 To 71
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        'Next
        '' SECCIÓN G: CAUSAS DE SUSPENSIÓN DE CIRUGIAS ELECTIVAS						
        'For ii = 75 To 83
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1 ****************************************************************************************************************************************

        xlHoja = Nothing
    End Sub 'OK
    Sub REM23()
        Dim ii, B(169), C(169), D(169), E(169), F(169), G(169), H(169), I(169), J(169), K(169), L(169), M(169), N(169), O(169), P(169), Q(169), R(169), S(169), T(169), U(169), V(169), W(169), X(169), Y(169), Z(169), AA(169), AB(169), AC(169), AD(169), AE(169), AF(169), AG(169), AH(169), AI(169), AJ(169), AK(169), AL(169), AM(169), AN(169), AO(169), AP(169), AQ(169), AR(169), AS1(169), AT(169), AU(169), AV(169), AW(169), AX(169), AY(169), AZ(169) As Integer
        xlHoja = xlLibro.Worksheets("A23")
        ' SECCIÓN A: INGRESOS AGUDOS SEGÚN DIAGNOSTICO
        For ii = 13 To 24
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
        Next
        ' SECCIÓN B: INGRESO CRÓNICO SEGÚN DIAGNÓSTICO (SOLO MÉDICO)
        For ii = 30 To 43
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
        Next
        ' SECCION C: EGRESOS
        'For ii = 49 To 56
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        'Next
        ' SECCIÓN D: CONSULTAS DE MORBILIDAD POR ENFERMEDADES RESPIRATORIAS EN SALAS IRA, ERA Y MIXTA
        For ii = 62 To 62
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
        Next
        ' SECCIÓN E: CONTROLES REALIZADOS
        For ii = 67 To 70
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
        Next
        ' SECCIÓN F: CONSULTAS ATENCIONES AGUDAS
        For ii = 75 To 77
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
        Next
        ' SECCIÓN G: INASISTENTES A CONTROL DE CRÓNICOS
        'For ii = 82 To 88
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        'Next
        '' SECCIÓN H: INASISTENTES A CITACIÓN AGENDADA
        'For ii = 92 To 95
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN I: PROCEDIMIENTOS REALIZADOS
        'For ii = 100 To 107
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        'Next
        '' SECCIÓN J: DERIVACIÓN DE PACIENTES SEGÚN DESTINO
        'For ii = 110 To 112
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN K: RECEPCIÓN DE PACIENTES SEGÚN ORIGEN
        'For ii = 115 To 119
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN L: HOSPITALIZACION ABREVIADA / INTERVENCION EN CRISIS RESPIRATORIA
        'For ii = 124 To 130
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        'Next
        '' SECCIÓN M: EDUCACIÓN EN SALAS
        '' SECCIÓN M.1: EDUCACIÓN INDIVIDUAL EN SALA IRA ERA
        'For ii = 134 To 140
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN M.2: EDUCACIÓN GRUPAL EN SALA (AGENDADA Y PROGRAMADA)
        'For ii = 143 To 152
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        '' SECCIÓN N: VISITAS DOMICILIARIAS REALIZADAS POR EQUIPO IRA-ERA A FAMILIAS
        'For ii = 157 To 160
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        'Next
        '' SECCIÓN O: PROGRAMA DE REHABILITACION PULMONAR 
        'For ii = 164 To 164
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN P: APLICACIÓN Y RESULTADO DE ENCUESTA CALIDAD DE VIDA
        'For ii = 168 To 169
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO
                Select Case (C(24) + AL(24) + AM(24))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [A]", "VAL [01]", "[ERROR]", "HOSPITAL PURRANQUE, No debe registrar Ingresos Agudos Según Diagnostico celdas C24 a AM24", "[" & (C(24) + AL(24) + AM(24)) & "]")
                        End With
                End Select
            Case 123101 ' HOSPITAL PURRANQUE
                Select Case (C(24) + AL(24) + AM(24))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [A]", "VAL [01]", "[ERROR]", "HOSPITAL PURRANQUE, No debe registrar Ingresos Agudos Según Diagnostico celdas C24 a AM24", "[" & (C(24) + AL(24) + AM(24)) & "]")
                        End With
                End Select
            Case 123102 ' HOSPITAL RIO NEGRO
                Select Case (C(24) + AL(24) + AM(24))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [A]", "VAL [01]", "[ERROR]", "HOSPITAL RIO NEGRO, No debe registrar Ingresos Agudos Según Diagnostico celdas C24 a AM24", "[" & (C(24) + AL(24) + AM(24)) & "]")
                        End With
                End Select
            Case Else ' Resto Establecimientos
                Select Case (C(24) + AL(24) + AM(24))
                    Case Is = 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [A]", "VAL [01]", "[ERROR]", "Ingresos Agudos Según Diagnostico: Deben registrar todos los Establecimientos, con excepción el HBSJO, HPU y HRN, celdas C24 a AM24", "[" & (C(24) + AL(24) + AM(24)) & "]")
                        End With
                End Select
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO
                Select Case (C(43) + AL(43))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [B]", "VAL [02]", "[ERROR]", "HBSJO, No debe registrar Ingreso Crónico Según Diagnostico, celdas C43 a AL43", "[" & (C(43) + AL(43)) & "]")
                        End With
                End Select
            Case 123101 ' HPU
                Select Case (C(43) + AL(43))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [B]", "VAL [02]", "[ERROR]", "HPU, No debe registrar Ingreso Crónico Según Diagnostico, celdas C43 a AL43", "[" & (C(43) + AL(43)) & "]")
                        End With
                End Select
            Case 123102 ' HRN
                Select Case (C(43) + AL(43))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [B]", "VAL [02]", "[ERROR]", "HRN, No debe registrar Ingreso Crónico Según Diagnostico, celdas C43 a AL43", "[" & (C(43) + AL(43)) & "]")
                        End With
                End Select
            Case Else
                Select Case (C(43) + AL(43))
                    Case Is = 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [B]", "VAL [02]", "[ERROR]", "Ingreso Crónico Según Diagnostico: Deben registrar todos los Establecimientos, con excepción el HBSJO, HPU y HRN, celdas C43 a AL43", "[" & (C(43) + AL(43)) & "]")
                        End With
                End Select
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO
                Select Case C(62)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [D]", "VAL [03]", "[ERROR]", "HBSJO No Debe Registrar, celda C62, Consultas de Morbilidad por Enfermedades Respiratorias en Salas, IRA, ERA y Mixta", "[" & C(62) & "]")
                        End With
                End Select
            Case 123101 ' HPU
                Select Case C(62)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [D]", "VAL [03]", "[ERROR]", "HPU No Debe Registrar, celda C62, Consultas de Morbilidad por Enfermedades Respiratorias en Salas, IRA, ERA y Mixta", "[" & C(62) & "]")
                        End With
                End Select
            Case 123102 ' HRN
                Select Case C(62)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [D]", "VAL [03]", "[ERROR]", "HRN No Debe Registrar, celda C62, Consultas de Morbilidad por Enfermedades Respiratorias en Salas, IRA, ERA y Mixta", "[" & C(62) & "]")
                        End With
                End Select
            Case Else
                Select Case C(62)
                    Case Is = 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [D]", "VAL [03]", "[ERROR]", " Consultas de Morbilidad por Enfermedades Respiratorias en Salas, IRA, ERA y Mixta, corresponde ingresar todos los Establecimientos, con excepción de HBSJO, HPU y HRN, celda C62", "[" & C(62) & "]")
                        End With
                End Select
        End Select
        '4************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 'hbo
                Select Case (C(70))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [E]", "VAL [04]", "[ERROR]", "HBSJO NO debe Registrar Controles Realizados,de las celdas C67 a C70", "[" & C(70) & "]")
                        End With
                End Select
            Case 123101 ' hpu
                Select Case (C(70))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [E]", "VAL [04]", "[ERROR]", "HPU NO debe Registrar Controles Realizados,de las celdas C67 a C70", "[" & C(70) & "]")
                        End With
                End Select
            Case 123102 'hrn
                Select Case (C(70))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [E]", "VAL [04]", "[ERROR]", "HRN NO debe Registrar Controles Realizados,de las celdas C67 a C70", "[" & C(70) & "]")
                        End With
                End Select
            Case Else
                Select Case (C(70))
                    Case Is = 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [E]", "VAL [04]", "[ERROR]", "Controles Realizados, corresponde ingresar todos los Establecimientos, con excepción de HBSJO, HPU y HRN, celdas C67 a C70", "[" & C(70) & "]")
                        End With
                End Select
        End Select
        '5************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBSJO
                Select Case (C(77))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [F]", "VAL [05]", "[ERROR]", "HBSJO NO debe Registrar, Seguimiento de Atenciones Realizadas en Agudos,celdas C75 a C77", "[" & C(77) & "]")
                        End With
                End Select
            Case 123101 ' HPU
                Select Case (C(77))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [F]", "VAL [05]", "[ERROR]", "HPU NO debe Registrar, Seguimiento de Atenciones Realizadas en Agudos,celdas C75 a C77", "[" & C(77) & "]")
                        End With
                End Select
            Case 123102 ' HRN
                Select Case (C(77))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [F]", "VAL [05]", "[ERROR]", "HRN NO debe Registrar, Seguimiento de Atenciones Realizadas en Agudos,celdas C75 a C77", "[" & C(77) & "]")
                        End With
                End Select
            Case Else
                Select Case (C(77))
                    Case Is = 0
                        With Me.DataGridView1.Rows
                            .Add("A23", " [F]", "VAL [05]", "[ERROR]", "Seguimiento de Atenciones Realizadas en Agudos, corresponde ingresar todos los Establecimientos, con excepción de HBSJO, HPU , HRN celdas C75 a C77", "[" & C(77) & "]")
                        End With
                End Select

        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM24()
        Dim ii, B(69), C(69), D(69), E(69), F(69), G(69), H(69), I(69), J(69), K(69), L(69), M(69), N(69), O(69), P(69), Q(69), R(69), S(69), T(69), U(69), V(69), W(69), X(69), Y(69), Z(69), AA(69), AB(69), AC(69), AD(69), AE(69), AF(69), AG(69), AH(69), AI(69), AJ(69), AK(69), AL(69), AM(69), AN(69), AO(69), AP(69), AQ(69), AR(69), AS1(69), AT(69), AU(69), AV(69), AW(69), AX(69), AY(69), AZ(69) As Integer
        xlHoja = xlLibro.Worksheets("A24")
        ' SECCIÓN A: INFORMACIÓN DE PARTOS Y ABORTOS ATENDIDOS 
        For ii = 11 To 20
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
        Next
        ' SECCIÓN A.1: 
        'For ii = 25 To 28
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        'Next
        '' SECCIÓN B:
        'For ii = 31 To 32
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN C.1
        'For ii = 37 To 38
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        'Next
        '' SECCIÓN C.2: 
        'For ii = 42 To 43
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCIÓN C3
        'For ii = 46 To 46
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN D: 
        'For ii = 50 To 51
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        ''SECCIÓN E:
        'For ii = 54 To 55
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN F:  TIPOS DE LACTANCIA EN NIÑOS Y NIÑAS AL EGRESO DE LA HOSPITALIZACIÓN
        'For ii = 61 To 64
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        'Next
        'For ii = 67 To 69
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************


        xlHoja = Nothing
    End Sub 'OK
    Sub REM25()
        Dim ii, B(150), C(150), D(150), E(150), F(150), G(150), H(150), I(150), J(150), K(150), L(150), M(150), N(150), O(150), P(150), Q(150), R(150), S(150), T(150), U(150), V(150), W(150), X(150), Y(150), Z(150), AA(150), AB(150), AC(150), AD(150), AE(150), AF(150), AG(150), AH(150), AI(150), AJ(150), AK(150), AL(150), AM(150), AN(150), AO(150), AP(150), AQ(150), AR(150), AS1(150), AT(150), AU(150), AV(150), AW(150), AX(150), AY(150), AZ(150) As Integer
        xlHoja = xlLibro.Worksheets("A25")
        ' SECCIÓN A.1: POBLACIÓN DONANTE (CS-UMT-BS)
        For ii = 12 To 17
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
        Next
        ' SECCIÓN A.2:  TIPO DE DONANTES RECHAZADOS
        'For ii = 22 To 28
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        'Next
        '' SECCIÓN A.3: REACCIONES ADVERSAS A LA DONACIÓN (CS - UMT - BS)									
        'For ii = 31 To 43
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN B: INGRESO UNIDADES DE SANGRE A PRODUCIÓN  (CS-BS)						
        'For ii = 47 To 49
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCIÓN C: PRODUCCIÓN DE COMPONENTES SANGUÍNEOS (CS-BS)
        'For ii = 53 To 62
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN C.1: COMPONENTES SANGUÍNEOS ELIMINADOS (CS-BS )								
        'For ii = 66 To 69
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next
        '' SECCIÓN C.2 .: COMPONENTES SANGUÍNEOS ELIMINADOS O DEVUELTOS AL CENTRO DE SANGRE ( UMT)											
        'For ii = 74 To 79
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next
        '' SECCIÓN C.3 : COMPONENTES SANGUÍNEOS TRANSFORMACIONES (CS-BS-UMT)							
        'For ii = 83 To 88
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next
        '' SECCIÓN C.4: COMPONENTES SANGUÍNEOS DISTRIBUÍBLES (CS)								
        'For ii = 93 To 98
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        '' SECCIÓN C.5: SATISFACCION STOCK (7 DÍAS) CS				
        'For ii = 102 To 106
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        ''SECCIÓN C.6: SATISFACCION STOCK CRITICO (3 DÍAS) UMT			
        'For ii = 109 To 113
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCIÓN D: COMPONENTES SANGUINEOS DISTRIBUIDOS (CS) O TRANSFERIDOS (BS Y UMT)
        'For ii = 116 To 121
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN D.1: TRANSFUSIONES (UMT - BS )
        'For ii = 127 To 133
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        ''SECCIÓN E: DEMANDA GLÓBULOS ROJOS PARA TRANSFUSIÓN (UMT-BS)
        'For ii = 136 To 137
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        ''SECCIÓN F: REACCIONES ADVERSAS POR ACTO* TRANSFUSIONAL  (UMT-BS)
        'For ii = 141 To 150
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next

        xlHoja = Nothing
    End Sub 'OK
    Sub REM26()
        Dim ii, B(85), C(85), D(85), E(85), F(85), G(85), H(85), I(85), J(85), K(85), L(85), M(85), N(85), O(85), P(85), Q(85), R(85), S(85), T(85), U(85), V(85), W(85), X(85), Y(85), Z(85), AA(85), AB(85), AC(85), AD(85), AE(85), AF(85), AG(85), AH(85), AI(85), AJ(85), AK(85), AL(85), AM(85), AN(85), AO(85), AP(85), AQ(85), AR(85), AS1(85), AT(85), AU(85), AV(85), AW(85), AX(85), AY(85), AZ(85) As Integer
        xlHoja = xlLibro.Worksheets("A26")
        ' SECCIÓN A: VISITAS DOMICILIARIAS INTEGRALES A FAMILIAS (ESTABLECIMIENTOS APS)
        For ii = 10 To 36
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
        Next
        'SECCIÓN B: OTRAS VISITAS INTEGRALES
        'For ii = 39 To 50
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        'Next
        '' SECCIÓN C:  VISITAS CON FINES DE TRATAMIENTOS Y/O PROCEDIMIENTOS EN  DOMICILIO
        'For ii = 54 To 63
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        ''SECCIÓN D: RESCATE DE PACIENTES INASISTENTES
        'For ii = 68 To 73
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        'Next
        '' SECCIÓN E: OTRAS VISITAS PROGRAMA DE ACOMPAÑAMIENTO PSICOSOCIAL EN ATENCIÓN PRIMARIA
        'For ii = 78 To 79
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        '' SECCIÓN F: APOYO TELEFÓNICO DEL PROGRAMA DE ACOMPAÑAMIENTO PSICOSOCIAL EN APS
        'For ii = 83 To 83
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        'For ii = 87 To 87
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        'Next
        xlHoja = Nothing
    End Sub 'OK
    Sub REM27()
        Dim ii, B(133), C(133), D(133), E(133), F(133), G(133), H(133), I(133), J(133), K(133), L(133), M(133), N(133), O(133), P(133), Q(133), R(133), S(133), T(133), U(133), V(133), W(133), X(133), Y(133), Z(133), AA(133), AB(133), AC(133), AD(133), AE(133), AF(133), AG(133), AH(133), AI(133), AJ(133), AK(133), AL(133), AM(133), AN(133), AO(133), AP(133), AQ(133), AR(133), AS1(133), AT(133), AU(133), AV(133), AW(133), AX(133), AY(133), AZ(133) As Integer
        xlHoja = xlLibro.Worksheets("A27")
        ' SECCIÓN A: PERSONAS QUE INGRESAN A EDUCACIÓN GRUPAL SEGÚN ÁREAS TEMÁTICAS Y EDAD
        For ii = 11 To 42
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
        Next
        ' SECCIÓN B: ACTIVIDADES DE EDUCACIÓN PARA LA SALUD SEGÚN PERSONAL QUE LAS REALIZA (SESIONES)
        For ii = 45 To 76
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
        Next
        ' SECCIÓN C: ACTIVIDAD FÍSICA GRUPAL PARA PROGRAMA SALUD CARDIOVASCULAR (SESIONES)
        'For ii = 77 To 81
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next
        '' SECCIÓN D: EDUCACIÓN GRUPAL A GESTANTES DE ALTO RIESGO OBSTÉTRICO (Nivel Secundario)
        'For ii = 84 To 86
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        '' SECCIÓN E: TALLERES PROGRAMA "MÁS A.M AUTOVALENTES"
        'For ii = 90 To 92
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        'Next
        '' SECCIÓN F: TALLERES PROGRAMA VIDA SANA
        'For ii = 96 To 98
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        'Next
        '' SECCIÓN G: INTERVENCIONES POR PATRÓN DE CONSUMO ALCOHOL y OTRAS SUSTANCIAS (PROGRAMA DIR EX PROGRAMA VIDA SANA ALCOHOL)
        'For ii = 102 To 110
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        'Next
        ''SECCIÓN H: PERSONAS QUE INGRESAN A TALLERES PARA PADRES DEL PROGRAMA DE APOYO A LA SALUD MENTAL INFANTIL (PASMI)
        'For ii = 113 To 113
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next
        ''SECCIÓN I: ORGANIZACIONES SOCIALES DE LA RED DEL PROGRAMA MÁS ADULTOS MAYORES AUTOVALENTES
        'For ii = 117 To 118
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        'Next
        ''SEECCIÓN J: SERVICIOS DE LA RED DEL PROGRAMA MÁS ADULTOS MAYORES AUTOVALENTES
        'For ii = 122 To 123
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        'Next
        ''SECCIÓN K: TALLER GRUPALES DE LACTANCIA MATERNA
        'For ii = 127 To 127
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        'Next
        ''SECCIÓN L: TALLERES GRUPALES  DE LACTANCIA MATERNA EN ATENCIÓN PRIMARIA 
        'For ii = 131 To 131
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        Select Case CodigoEstablec ' VIENE DEL REM A03
            Case 123300, 123301, 123302, 123303, 123306, 123310, 123404, 123425, 123700, 123701, 123311, 123312, 123307, 123411, 123412, 123413, 123414, 123415, 123416, 123417, 123419, 123420
                If A03(6) = 0 And (D(104) + D(105) + D(106) + D(107) + D(108) + D(109) + D(110) + D(111) + D(112) > 0) Then
                    With Me.DataGridView1.Rows
                        .Add("A03", " [D.1][G]", "VAL [18]", "[ERROR]", "Aplicación de Tamizaje para evaluar el nivel de riesgo de consumo de alcohol tabaco y otras sustancias, si existe registro en Resultados de Evaluación, celda C102:C104, debe registrar en Sección G, REM27, Numero de Intervenciones Celdas D104:D112. Validación solo para Establecimientos de Comunas de: Osorno, S.J.Costa (B.Mansa - Puacho) y Purranque", "[" & A03(6) & "-" & D(104) + D(105) + D(106) + D(107) + D(108) + D(109) + D(110) + D(111) + D(112) & "]")
                    End With
                End If
            Case Else
        End Select
        '1*************************************************************************************************************************************************************************************
        Select Case D(42)
            Case Is <> 0
                If D(76) = 0 Then
                    With Me.DataGridView1.Rows
                        .Add("A27", " [A][B]", "VAL [01]", "[REVISAR]", "Personas que Ingresan a Educación Grupal según Áreas Temáticas y Edad, celda D42, Si existen datos en el total de la sección deben existir datos en Sección B: Actividades de Educación  para la Salud personal Según personal que las Realiza, Celda D76.", "[" & D(42) & "-" & D(76) & "]")
                    End With
                End If
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case D(22)
            Case Is <> 0
                If (Y(22) + Z(22)) = 0 Then
                    With Me.DataGridView1.Rows
                        .Add("A27", " [A]", "VAL [02]", "[REVISAR]", "Personas que Ingresan a Educación Grupal según Áreas Temáticas y Edad, Si existe información en celda D22 entonces debe existir información en Gestantes celdas Y22 a Z22", "[" & D(21) & "-" & (Y(21) + Z(21)) & "]")
                    End With
                End If
        End Select

        xlHoja = Nothing
    End Sub 'OK 
    Sub REM28()
        Dim ii, B(323), C(323), D(323), E(323), F(323), G(323), H(323), I(323), J(323), K(323), L(323), M(323), N(323), O(323), P(323), Q(323), R(323), S(323), T(323), U(323), V(323), W(323), X(323), Y(323), Z(323), AA(323), AB(323), AC(323), AD(323), AE(323), AF(323), AG(323), AH(323), AI(323), AJ(323), AK(323), AL(323), AM(323), AN(323), AO(323), AP(323), AQ(323), AR(323), AS1(323), AT(323), AU(323), AV(323), AW(323), AX(323), AY(323), AZ(323) As Integer
        xlHoja = xlLibro.Worksheets("A28")
        ' A. NIVEL PRIMARIO 
        ' SECCIÓN A.1: INGRESOS Y EGRESOS  AL PROGRAMA DE REHABILITACIÓN INTEGRAL
        For ii = 13 To 26
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
        Next
        ' SECCION A.2: INGRESOS POR CONDICIÓN DE SALUD
        For ii = 31 To 56
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
        Next
        ' SECCIÓN A.3: EVALUACIÓN INICIAL
        For ii = 61 To 66
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        ' SECCIÓN A.4: EVALUACIÓN INTERMEDIA
        'For ii = 69 To 73
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        '' SECCIÓN A.5: SESIONES DE REHABILITACIÓN
        'For ii = 76 To 80
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        ''SECCIÓN A.6: PROCEDIMIENTOS Y ACTIVIDADES
        'For ii = 83 To 112
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN A.7: CONSEJERÍA INDIVIDUAL AGENDADA
        'For ii = 115 To 120
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN A.8: CONSEJERÍA FAMILIAR AGENDADA
        'For ii = 123 To 128
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        'Next
        '' SECCIÓN A.9: VISITAS DOMICILIARIAS INTEGRALES
        'For ii = 132 To 133
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        'Next
        '' SECCIÓN A.10: NÚMERO DE PERSONAS QUE INGRESAN Y NÚMERO DE SESIONES DE EDUCACIÓN GRUPAL
        'For ii = 137 To 138
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        'Next
        '' SECCIÓN A.11: PERSONAS QUE LOGRAN PARTICIPACION EN COMUNIDAD 
        'For ii = 142 To 144
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        'Next
        '' SECCIÓN A.12: ACTIVIDADES Y PARTICIPACION
        'For ii = 148 To 164
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        'Next
        '' SECCIÓN B: NIVEL HOSPITALARIO
        '' SECCIÓN B.1: INGRESOS Y EGRESOS  AL PROGRAMA DE REHABILITACIÓN INTEGRAL
        For ii = 166 To 198
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        Next
        '' SECCIÓN B.2: EVALUACIÓN INICIAL
        'For ii = 205 To 211
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        'Next
        '' SECCIÓN B.3: EVALUACION INTERMEDIA
        'For ii = 214 To 219
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCIÓN B.4: SESIONES DE REHABILITACIÓN
        'For ii = 222 To 226
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCION B.5: DERIVACIONES Y CONTINUIDAD EN LOS CUIDADOS
        'For ii = 229 To 231
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCION B.6: PROCEDIMIENTOS Y OTRAS ACTIVIDADES
        'For ii = 234 To 269
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        'Next
        '' SECCION C: : NÚMERO DE AYUDAS TÉCNICAS ENTREGADAS POR TIPO
        'For ii = 274 To 274
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        'Next
        'For ii = 276 To 293
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        'Next
        '' SECCION C.1: : PERSONAS CON CONDICIÓN DE SALUD QUE RECIBE  AYUDAS TÉCNICAS
        'For ii = 297 To 297
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        'Next
        'For ii = 299 To 319
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case B(13)
            Case Is <> D(31)
                With Me.DataGridView1.Rows
                    .Add("A28", " [A.1][A.2]", "VAL [01]", "[ERROR]", "Ingresos y Egresos al programa de Rehabilitación Integral, los  Ingresos al programa, celdas B13, debe ser igual a la suma total de Sección A.2: Ingresos por condición de Salud, celda D31.", "[" & B(13) & "-" & D(31) & "]")
                End With
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case B(13)
            Case Is < B(14)
                With Me.DataGridView1.Rows
                    .Add("A28", " [A.1]", "VAL [02]", "[ERROR]", "Ingresos y Egresos al programa de Rehabilitación Integral, celda B13 debe ser mayor o igual a Ingresos celda B14", "[" & B(13) & "-" & B(14) & "]")
                End With
        End Select
        '3************************************************************************************************************************************************************************************
        Select Case B(14)
            Case Is < B(15)
                With Me.DataGridView1.Rows
                    .Add("A28", " [A.1]", "VAL [03]", "[ERROR]", "Ingresos y Egresos al programa de Rehabilitación Integral, celda B14 debe ser mayor o igual a Ingresos celda B15", "[" & B(14) & "-" & B(15) & "]")
                End With
        End Select
        '4***********************************************************************************************************************************************************************************
        Select Case D(31)
            Case Is > B(66)
                With Me.DataGridView1.Rows
                    .Add("A28", " [A.2][A.3]", "VAL [04]", "[ERROR]", "Ingresos por Condición de Salud, Total Ingreso de Personas, celda D31 debe ser MENOR o IGUAL al total  de Sección A.3: Evaluación Inicial, celda B66", "[" & D(31) & "-" & B(66) & "]")
                End With
        End Select
        '5************************************************************************************************************************************************************************************
        Select Case D(31)
            Case Is > (D(32) + D(33) + D(34) + D(35) + D(36) + D(37) + D(38) + D(39) + D(40) + D(41) + D(42) + D(43) + D(44) + D(45) + D(46) + D(47) + D(48) + D(49) + D(50) + D(51) + D(52) + D(53) + D(54) + D(55) + D(56))
                With Me.DataGridView1.Rows
                    .Add("A28", " [A.2]", "VAL [05]", "[ERROR]", "Ingresos por Condición de Salud, Celda D31, Total Ingresos Personas debe ser menor o igual a la suma de las desagregaciones por condición Física, celdas D32 a D56.", "[" & D(31) & " - " & (D(32) + D(33) + D(34) + D(35) + D(36) + D(37) + D(38) + D(39) + D(40) + D(41) + D(42) + D(43) + D(44) + D(45) + D(46) + D(47) + D(48) + D(49) + D(50) + D(51) + D(52) + D(53) + D(54) + D(55) + D(56)) & "]")
                End With
        End Select
        '6************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123301, 123310, 123307, 123306, 123309, 123302, 123304  ' CESFAM Lopetegui, V Centenario, Pampa Alegre, Purranque, Pablo  Araya, CESFAM Ovejeria y Cesfam Puyehue.
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (AQ(31) + AQ(32) + AQ(33) + AQ(34) + AQ(35) + AQ(36) + AQ(37) + AQ(38) + AQ(39) + AQ(40) + AQ(41) + AQ(42) + AQ(43) + AQ(44) + AQ(45) + AQ(46) + AQ(47) + AQ(48) + AQ(49) + AQ(50) + AQ(51) + AQ(52) + AQ(53) + AQ(54) + AQ(55) + AQ(56))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A28", " [A.2]", "VAL [06]", "[ERROR]", "Ingresos por Condición de Salud, celda AQ31:AQ56,  Tipo de Estrategia debe Corresponder Sólo a CESFAM Lopetegui, V Centenario, Pampa Alegre, Purranque, Pablo  Araya, CESFAM Ovejeria y Cesfam Puyehue", "[" & (AQ(31) + AQ(32) + AQ(33) + AQ(34) + AQ(35) + AQ(36) + AQ(37) + AQ(38) + AQ(39) + AQ(40) + AQ(41) + AQ(42) + AQ(43) + AQ(44) + AQ(45) + AQ(46) + AQ(47) + AQ(48) + AQ(49) + AQ(50) + AQ(51) + AQ(52) + AQ(53) + AQ(54) + AQ(55) + AQ(56)) & "]")
                        End With
                End Select
        End Select
        '7************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123300, 123303, 123101, 123102, 123422, 123423, 123424, 123426, 123427, 123428, 123104, 123311, 123312, 123105, 123305, 123103 ' CESFAM Jauregui, R. Alto, Bahía Mansa, Puaucho, S. Pablo, Puyehue, DSM Pto.  Octay, Hospitales; Purranque, R. Negro, P. Octay, M.S. Juan, M. Quilacahuín
            Case Else ' RESTO ESTABLECIMIENTOS 
                Select Case (AR(31) + AR(32) + AR(33) + AR(34) + AR(35) + AR(36) + AR(37) + AR(38) + AR(39) + AR(40) + AR(41) + AR(42) + AR(43) + AR(44) + AR(45) + AR(46) + AR(47) + AR(48) + AR(49) + AR(50) + AR(51) + AR(52) + AR(53) + AR(54) + AR(55) + AR(56))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A28", " [A.2]", "VAL [07]", "[ERROR]", "Ingresos por Condición de Salud, celda AR31:AR56, Tipo de Estrategia  debe Corresponder Sólo a CESFAM Jauregui, R. Alto, Bahía Mansa, Puaucho, S. Pablo, Puyehue, DSM Pto.  Octay, Hospitales; Purranque, R. Negro, P. Octay, M.S. Juan, M. Quilacahuín", "[" & (AR(31) + AR(32) + AR(33) + AR(34) + AR(35) + AR(36) + AR(37) + AR(38) + AR(39) + AR(40) + AR(41) + AR(42) + AR(43) + AR(44) + AR(45) + AR(46) + AR(47) + AR(48) + AR(49) + AR(50) + AR(51) + AR(52) + AR(53) + AR(54) + AR(55) + AR(56)) & "]")
                        End With
                End Select
        End Select
        '8************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123309, 123410, 123434, 123709, 123422, 123423, 123424, 123426, 123427, 123428, 123311, 123430, 123431, 123312, 123402, 123305, 123432, 123436, 123304, 123406, 123407, 123408, 123705 ' EQUIPO RURAL RIO NEGRO - POSTAS PUERTO OCTAY - COMUNA SAN JUAN DE LA COSTA-COMUNA SAN PABLO - COMUNA PUYEHUE
            Case Else ' RESTO DE ESTABLECIMIENTOS
                Select Case (AS1(31) + AS1(32) + AS1(33) + AS1(34) + AS1(35) + AS1(36) + AS1(37) + AS1(38) + AS1(39) + AS1(40) + AS1(41) + AS1(42) + AS1(43) + AS1(44) + AS1(45) + AS1(46) + AS1(47) + AS1(48) + AS1(49) + AS1(50) + AS1(51) + AS1(52) + AS1(53) + AS1(54) + AS1(55) + AS1(56))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A28", " [A.2]", "VAL [08]", "[ERROR]", "Ingresos por Condición de Salud, celda AS31:AS56, Tipo de Estrategia debe Corresponder Sólo a  Equipo Rural Móvil (ERM) para las camunas de R. Negro, Pto. Octay, S.J. Costa.", "[" & (AS1(31) + AS1(32) + AS1(33) + AS1(34) + AS1(35) + AS1(36) + AS1(37) + AS1(38) + AS1(39) + AS1(40) + AS1(41) + AS1(42) + AS1(43) + AS1(44) + AS1(45) + AS1(46) + AS1(47) + AS1(48) + AS1(49) + AS1(50) + AS1(51) + AS1(52) + AS1(53) + AS1(54) + AS1(55) + AS1(56)) & "]")
                        End With
                End Select
        End Select
        '9************************************************************************************************************************************************************************************
        Select Case B(166)
            Case Is > (B(167) + B(168) + B(169) + B(170) + B(171) + B(172) + B(173) + B(174) + B(175) + B(176) + B(177) + B(178) + B(179) + B(180) + B(181) + B(182) + B(183) + B(184) + B(185) + B(186) + B(187) + B(188) + B(189) + B(190) + B(191) + B(192))
                With Me.DataGridView1.Rows
                    .Add("A28", " [B.1]", "VAL [09]", "[ERROR]", " Ingresos y Egresos al Programa de Rehabilitación Integral, el Total de Ingresos B166 debe ser menor o igual a la sumatoria de ingresos a rehabilitación, celdas B167 a B192", "[" & B(166) & " - " & (B(167) + B(168) + B(169) + B(170) + B(171) + B(172) + B(173) + B(174) + B(175) + B(176) + B(177) + B(178) + B(179) + B(180) + B(181) + B(182) + B(183) + B(184) + B(185) + B(186) + B(187) + B(188) + B(189) + B(190) + B(191) + B(192)) & "]")
                End With
        End Select
        '10************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100, 123101, 123102, 123103, 123104, 123105 ' HOSPITALES
            Case Else ' RESTO DE ESTABLECIMIENTOS
                Select Case (AQ(166) + AQ(167) + AQ(168) + AQ(169) + AQ(170) + AQ(171) + AQ(172) + AQ(173) + AQ(174) + AQ(175) + AQ(176) + AQ(177) + AQ(178) + AQ(179) + AQ(180) + AQ(181) + AQ(182) + AQ(183) + AQ(184) + AQ(185) + AQ(186) + AQ(187) + AQ(188) + AQ(189) + AQ(190) + AQ(191) + AQ(192))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A28", " [B.1]", "VAL [10]", "[ERROR]", "Ingresos y Egresos al Programa de Rehabilitación Integral, Los Hospitales Río Negro y Hospital Purranque deben registrar solo en Tipo de atención Abierta, celdas AQ166:AQ192.", "[" & (AQ(166) + AQ(167) + AQ(168) + AQ(169) + AQ(170) + AQ(171) + AQ(172) + AQ(173) + AQ(174) + AQ(175) + AQ(176) + AQ(177) + AQ(178) + AQ(179) + AQ(180) + AQ(181) + AQ(182) + AQ(183) + AQ(184) + AQ(185) + AQ(186) + AQ(187) + AQ(188) + AQ(189) + AQ(190) + AQ(191) + AQ(192)) & "]")
                        End With
                End Select
        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM29()
        Dim ii, B(94), C(94), D(94), E(94), F(94), G(94), H(94), I(94), J(94), K(94), L(94), M(94), N(94), O(94), P(94), Q(94), R(94), S(94), T(94), U(94), V(94), W(94), X(94), Y(94), Z(94), AA(94), AB(94), AC(94), AD(94), AE(94), AF(94), AG(94), AH(94), AI(94), AJ(94), AK(94), AL(94), AM(94), AN(94), AO(94), AP(94), AQ(94), AR(94), AS1(94), AT(94), AU(94), AV(94), AW(94), AX(94), AY(94), AZ(94) As Integer
        xlHoja = xlLibro.Worksheets("A29")
        ' SECCIÓN A: PROGRAMA DE RESOLUTIVIDAD ATENCIÓN PRIMARIA DE SALUD
        For ii = 11 To 26
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        ' SECCIÓN B: PROCEDIMIENTOS DE IMÁGENES DIAGNOSTICAS Y PROGRAMA DE RESOLUTIVIDAD EN APS  
        For ii = 30 To 59
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
        Next
        '' SECCIÓN B.1: PROGRAMA DE IMÁGENES DIAGNOSTICAS EN ATENCION PRIMARIA
        'For ii = 60 To 79
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123103, 123301, 123307, 123304, 123305, 123312, 123309, 123422, 123423, 123424, 123426, 123427, 123428

            Case Else
                Select Case (C(11) + C(12) + C(13) + C(14) + C(15) + C(16) + C(17) + C(18) + C(19) + C(20) + C(21) + C(22) + C(23) + C(24) + C(25) + C(26))
                    Case Is > 0
                        With Me.DataGridView1.Rows
                            .Add("A29", "[A]", "VAL [01]", "[REVISAR]", "Programa de Resolutividad Atención Primaria de Salud, celdas C11:C26, solo pueden tener registro los establecimientos de Hospital Puerto Octay y CESFAMS: Lopetegui, Purranque, Puyehue, San Pablo y Puaucho", "[" & (C(11) + C(12) + C(13) + C(14) + C(15) + C(16) + C(17) + C(18) + C(19) + C(20) + C(21) + C(22) + C(23) + C(24) + C(25) + C(26)) & "]")
                        End With
                End Select
        End Select
        '2************************************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123300, 123301, 123303, 123306, 123310 ' CESFAM OSORNO
            Case Else
                Select Case (N(12) + O(12))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A29", "[A]", "VAL [02]", "[REVISAR]", "Programa de Resolutividad Atención Primaria de Salud, Consulta por Oftalmología, celdas N12:O12, solo pueden tener registro los CESFAM de Osorno", "[" & (N(12) + O(12)) & "]")
                        End With
                End Select
        End Select
        '4***************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123301, 123306, 123302, 123305, 123312, 123304, 123307, 123309 ' Cesfam Lopetegui -Cesfam Pampa Alegre- Cesfam(Ovejería)-Cesfam San Pablo - Cesfam(Puaucho)-Cesfam Entre Lagos - Cesfam(Purranque) - Cesfam Pablo Araya
            Case Else
                Select Case (C(50) + C(51) + C(52) + C(53))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A29", "[B]", "VAL [03]", "[ERROR]", "Procedimientos de imágenes diagnósticos y programa de Resolutividad en APS, En Cirugía Menor, Celdas C50:C53, deben registrar solo CESFAM: Lopetegui, Pampa Alegre, Ovejería y San Pablo", "[" & (C(50) + C(51) + C(52) + C(53)) & "]")
                        End With
                End Select
        End Select


        xlHoja = Nothing
    End Sub 'OK
    Sub REM30()
        Dim ii, B(135), C(135), D(135), E(135), F(135), G(135), H(135), I(127), J(127), K(127), L(127), M(127), N(127), O(127), P(127), Q(127), R(127), S(127), T(127), U(127), V(127), W(127), X(127), Y(127), Z(127), AA(127), AB(127), AC(127), AD(127), AE(127), AF(127), AG(127), AH(127), AI(127), AJ(127), AK(127), AL(127), AM(127), AN(127), AO(127), AP(127), AQ(127), AR(127), AS1(127), AT(127), AU(127), AV(127), AW(127), AX(127), AY(127), AZ(127) As Integer
        xlHoja = xlLibro.Worksheets("A30")
        ' SECCION A:  CONSULTAS MEDICAS DE ESPECIALIDAD RESUELTAS POR  TELEMEDICINA (TELECONSULTA)
        For ii = 17 To 77
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
        Next
        ' SECCIÓN B : CONSULTAS MÉDICAS DE URGENCIA RESUELTAS POR TELEMEDICINA EN ESTABLECIMIENTOS ATENCIÓN SECUNDARIA DE URGENCIA (TELECONSULTA)
        'For ii = 82 To 84
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        'Next
        ' SECCIÓN C : INFORMES POR TELEMEDICINA EN ESTABLECIMIENTOS ATENCIÓN PRIMARIA Y SECUNDARIA 
        For ii = 90 To 103
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
        Next
        ' SECCION D: TELECONSULTA AMBULATORIA EN ESPECIALIDAD ODONTOLOGICA
        'For ii = 106 To 117
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        'Next
        '' SECCION E: TELEPROCEDIMIENTOS EN ATENCION SECUNDARIA Y TERCIARIA
        'For ii = 123 To 127
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        'Next


        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1*******************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO
            Case Else ' RESTO DE ESTABLECIMIENTOS
                Select Case (B(19) + C(19) + D(19) + E(19) + F(19) + G(19) + H(19) + I(19))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A30", "[A]", "VAL [01]", "[ERROR]", "Teleconsulta Médicas de especialidad, Consultas Ambulatorias y Hospitalizadas, Especialidades de Medicina Interna, Celdas B19:I19, debe registrar solo HBSJO", "[" & (B(19) + C(19) + D(19) + E(19) + F(19) + G(19) + H(19) + I(19)) & "]")
                        End With
                End Select
        End Select
        '2*******************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123101, 123104, 123102, 123100, 123304 'HOSPITALES
            Case Else ' RESTO DE ESTABLECIMIENTOS
                Select Case (B(38) + C(38) + D(38) + E(38) + F(38) + G(38) + H(38) + I(38))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A30", "[A]", "VAL [02]", "[ERROR]", "Teleconsulta Médicas de especialidad, Especialidades de Dermatología, Consultas Ambulatorias, Celda B38:I38, deben registrar solo establecimientos Hospital Purranque, Hospital Río Negro, Hospital Misión San Juan, HBSJO y Cesfam Puyehue", "[" & (B(38) + C(38) + D(38) + E(38) + F(38) + G(38) + H(38) + I(38)) & "]")
                        End With
                End Select
        End Select
        '3*******************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO
            Case Else ' RESTO DE ESTABLECIMIENTOS
                Select Case (B(49) + C(49) + D(49) + E(49) + F(49) + G(49) + H(49) + I(49))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A30", "[A]", "VAL [03]", "[ERROR]", "Teleconsulta Médicas de especialidad, Consultas Ambulatorias, Especialidades de Psiquiatría Adultos, Celdas B49:I49, debe registrar solo HBSJO.", "[" & (B(49) + C(49) + D(49) + E(49) + F(49) + G(49) + H(49) + I(49)) & "]")
                        End With
                End Select
        End Select
        '4*******************************************************************************************************************************
        Select Case (B(19) + C(19) + D(19) + E(19) + F(19) + G(19) + H(19) + I(19))
            Case Is <> N(19)
                With Me.DataGridView1.Rows
                    .Add("A30", "[A]", "VAL [04]", "[ERROR]", "Teleconsulta Médicas de especialidad, Especialidades de Medicina Interna, celdas B19:I19, debe ser igual a Modalidad Institucional, Celda N19", "[" & (B(19) + C(19) + D(19) + E(19) + F(19) + G(19) + H(19) + I(19)) & " - " & N(19) & "]")
                End With
        End Select
        '5*******************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123101, 123102, 123103, 123104, 123105 ' HOSPITALES
            Case Else ' RESTO DE ESTABLECIMIENTOS
                Select Case (X(19) + Y(19) + Z(19) + AA(19))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("A30", "[A]", "VAL [05]", "Teleconsulta Médicas de especialidad, (Incluye consultas ambulatorias y de Hospitalizados), Especialidades de Medicina Interna, Celda X19:AA19, deben registrar establecimientos HPU, HRN, HPO, HMSJ, HPSQ", "[" & (X(19) + Y(19) + Z(19) + AA(19)) & "]")
                        End With
                End Select
        End Select
        '6*******************************************************************************************************************************
        Select Case (B(90) + C(90) + D(90) + E(90))
            Case Is <> (G(90) + H(90))
                With Me.DataGridView1.Rows
                    .Add("A30", "[C]", "VAL [06]", "[REVISAR]", "Teleinformes en Establecimientos de Atención Primaria, Segunda Y Terciaria, celda B90:E90, debe ser iguales a registros en Modalidad de Compra de Servicio, Celda G90:H90", "[" & (B(90) + C(90) + D(90) + E(90)) & "-" & (G(90) + H(90)) & "]")
                End With
        End Select

        xlHoja = Nothing
    End Sub 'OK
    Sub REM31()
        Dim ii, B(93), C(93), D(93), E(93), F(93), G(93), H(93), I(93), J(93), K(93), L(93), M(93), N(66), O(93), P(93), Q(93), R(93), S(93), T(93), U(93), V(93), W(93), X(93), Y(93), Z(93), AA(93), AB(93), AC(93), AD(93), AE(93), AF(93), AG(93), AH(93), AI(93), AJ(93), AK(93), AL(93), AM(93), AN(93), AO(93), AP(93), AQ(93), AR(93), AS1(93), AT(93), AU(93), AV(93), AW(93), AX(93), AY(93), AZ(93) As Integer
        xlHoja = xlLibro.Worksheets("A31")
        ' SECCIÓN A: TIPOS DE TERAPIAS ENTREGADAS
        For ii = 13 To 29
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
        Next

        '' SECCIÓN B: PERSONAS QUE RECIBEN TERAPIAS
        'For ii = 34 To 36
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        'Next

        '' SECCIÓN C: PROFESIONAL QUE ENTREGA LAS  TERAPIAS
        'For ii = 41 To 54
        '    B(ii) = xlHoja.Range("B" & ii & "").Value
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        'Next

        ' SECCIÓN D: MARCO DE ATENCION DE TERAPIAS COMPLEMENTARIAS
        'For ii = 59 To 74
        '    C(ii) = xlHoja.Range("C" & ii & "").Value
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        'Next

        'SECCIÓN E: TIPOS DE TERAPIAS Y PRACTICAS DE BIENESTAR  ENTREGADAS EN ATENCION GRUPAL Y COMUNITARIA
        'For ii = 80 To 92
        '    'C(ii) = xlHoja.Range("C" & ii & "").Value
        '    'D(ii) = xlHoja.Range("D" & ii & "").Value
        '    'E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        'NexT

        xlHoja = Nothing
    End Sub 'OK
    Private Sub BtnSalir_Click(sender As System.Object, e As System.EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub
    Function CargaSerieA()
        Select Case MsgBox("Desea efectuar validación Serie A", MsgBoxStyle.YesNo, "Carga de validación")
            Case MsgBoxResult.Yes
                ValidarNombre = Microsoft.VisualBasic.Left(NombreArchivo, 6)
                If IsNumeric(ValidarNombre) Then
                    CodigoEstablec = CInt(ValidarNombre)

                    Select Case CodigoEstablec ' todos los establecimientos del servicio
                        Case 123000, 123010, 123011, 123012, 123030, 123100, 123101, 123102, 123103, 123104, 123105, 123300, 123301, 123302, 123303, 123304, 123305, 123306, 123307, 123309, 123310, 123311, 123312, 123402, 123404, 123406, 123407, 123408, 123410, 123411, 123412, 123413, 123414, 123415, 123416, 123417, 123419, 123420, 123422, 123423, 123424, 123425, 123426, 123427, 123428, 123430, 123431, 123432, 123434, 123435, 123436, 123437, 123700, 123701, 123705, 123709, 123800, 200209, 123207, 200445, 200455, 200490, 200248, 200477, 200539
                            'lo que valida
                            Try
                                xlExcel = New Excel.Application
                                xlLibro = xlExcel.Workbooks.Open(OpenFileDialog1.FileName)
                                'HOJA NOMBRE
                                NOMBRE()
                                CONTROL_A()

                                If (ValidaAno = "2020") Then
                                    CargaLabel()
                                    COMPLETITUD_SERIEA()
                                    MsgBox("Proceso Finalizado con Éxito", MsgBoxStyle.Information, "Fin proceso")

                                    Me.LBerrores.Text = Me.DataGridView1.RowCount - 1
                                    Me.LblHojaControl.Text = ValidaHojaControl

                                    xlLibro.Saved = True
                                    xlLibro.Close()
                                    xlExcel.Quit()

                                Else ' cuando no corresponde al año

                                    LimpiarLabel()

                                    MsgBox("El archivo seleccionado no corresponde al año en curso", MsgBoxStyle.Exclamation)

                                    xlLibro.Saved = True
                                    xlExcel.ActiveWorkbook.Close()
                                    xlExcel.Quit()
                                End If


                            Catch ex As Exception
                                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al leer archivo")
                                Return False
                            End Try


                        Case Else
                            MsgBox("El archivo seleccionado no corresponde a un establecimiento del servicio de salud osorno", MsgBoxStyle.Exclamation)
                    End Select

                End If
        End Select

        Return True
    End Function
    Function CargaSerieP()
        Select Case MsgBox("Desea efectuar validacion serie P", MsgBoxStyle.YesNo, "Carga de validación")

            Case MsgBoxResult.Yes
                ValidarNombre = Microsoft.VisualBasic.Left(NombreArchivo, 6)
                If IsNumeric(ValidarNombre) Then
                    CodigoEstablec = CInt(ValidarNombre)

                    Select Case CodigoEstablec ' todos los establecimientos del servicio
                        Case 123000, 123010, 123011, 123012, 123030, 123100, 123101, 123102, 123103, 123104, 123105, 123300, 123301, 123302, 123303, 123304, 123305, 123306, 123307, 123309, 123310, 123311, 123312, 123402, 123404, 123406, 123407, 123408, 123410, 123411, 123412, 123413, 123414, 123415, 123416, 123417, 123419, 123420, 123422, 123423, 123424, 123425, 123426, 123427, 123428, 123430, 123431, 123432, 123434, 123435, 123436, 123437, 123700, 123701, 123705, 123709, 123800, 200209, 123207, 200445, 200455, 200490, 200248, 200477
                            'lo que valida

                            Try
                                xlExcel = New Excel.Application
                                xlLibro = xlExcel.Workbooks.Open(OpenFileDialog1.FileName)
                                'HOJA NOMBRE
                                NOMBRE()
                                CONTROL_P()

                                If (ValidaAno = "2020") Then
                                    CargaLabel()
                                    COMPLETITUD_SERIEP()
                                    MsgBox("Proceso Finalizado con Éxito", MsgBoxStyle.Information, "Fin proceso")

                                    Me.LBerrores.Text = Me.DataGridView1.RowCount - 1
                                    Me.LblHojaControl.Text = ValidaHojaControl

                                    xlLibro.Saved = True
                                    xlExcel.ActiveWorkbook.Close()
                                    xlExcel.Quit()

                                Else ' cuando no corresponde al año

                                    LimpiarLabel()

                                    MsgBox("El archivo seleccionado no corresponde al año en curso", MsgBoxStyle.Exclamation)

                                    xlLibro.Saved = True
                                    xlExcel.ActiveWorkbook.Close()
                                    xlExcel.Quit()
                                End If

                            Catch ex As Exception

                            End Try
                        Case Else
                            MsgBox("El archivo seleccionado no corresponde a un establecimiento del servicio de salud osorno", MsgBoxStyle.Exclamation)
                    End Select

                End If
        End Select

        Return True
    End Function
    Function CargaSerieBM()
        Select Case MsgBox("Desea efectuar validacion serie BM", MsgBoxStyle.YesNo, "Carga de validación")

            Case MsgBoxResult.Yes
                ValidarNombre = Microsoft.VisualBasic.Left(NombreArchivo, 6)
                If IsNumeric(ValidarNombre) Then
                    CodigoEstablec = CInt(ValidarNombre)

                    Select Case CodigoEstablec ' todos los establecimientos del servicio
                        Case 123010, 123011, 123012, 123030, 123100, 123101, 123102, 123103, 123104, 123105, 123300, 123301, 123302, 123303, 123304, 123305, 123306, 123307, 123309, 123310, 123311, 123312, 123402, 123404, 123406, 123407, 123408, 123410, 123411, 123412, 123413, 123414, 123415, 123416, 123417, 123419, 123420, 123422, 123423, 123424, 123425, 123426, 123427, 123428, 123430, 123431, 123432, 123434, 123435, 123436, 123437, 123700, 123701, 123705, 123709, 123800, 200209, 123207
                            'lo que valida

                            Try
                                xlExcel = New Excel.Application
                                xlLibro = xlExcel.Workbooks.Open(OpenFileDialog1.FileName)
                                'HOJA NOMBRE
                                NOMBRE()
                                CONTROL_BM()

                                If (ValidaAno = "2020") Then
                                    CargaLabel()
                                    COMPLETITUD_SERIEBM()
                                    MsgBox("Proceso Finalizado con Éxito", MsgBoxStyle.Information, "Fin proceso")

                                    Me.LBerrores.Text = Me.DataGridView1.RowCount - 1
                                    Me.LblHojaControl.Text = ValidaHojaControl

                                    xlLibro.Saved = True
                                    xlExcel.ActiveWorkbook.Close()
                                    xlExcel.Quit()

                                Else ' cuando no corresponde al año

                                    LimpiarLabel()

                                    MsgBox("El archivo seleccionado no corresponde al año en curso", MsgBoxStyle.Exclamation)

                                    xlLibro.Saved = True
                                    xlExcel.ActiveWorkbook.Close()
                                    xlExcel.Quit()
                                End If

                            Catch ex As Exception

                            End Try
                        Case Else
                            MsgBox("El archivo seleccionado no corresponde a un establecimiento del servicio de salud osorno", MsgBoxStyle.Exclamation)
                    End Select

                End If
        End Select

        Return True
    End Function
    Function CargaSerieD()
        Select Case MsgBox("Desea efectuar validación serie D", MsgBoxStyle.YesNo, "Carga de validación")
            Case MsgBoxResult.Yes
                ValidarNombre = Microsoft.VisualBasic.Left(NombreArchivo, 6)
                If IsNumeric(ValidarNombre) Then
                    CodigoEstablec = CInt(ValidarNombre)

                    Select Case CodigoEstablec ' todos los establecimientos del servicio
                        Case 123010, 123011, 123012, 123030, 123100, 123101, 123102, 123103, 123104, 123105, 123300, 123301, 123302, 123303, 123304, 123305, 123306, 123307, 123309, 123310, 123311, 123312, 123402, 123404, 123406, 123407, 123408, 123410, 123411, 123412, 123413, 123414, 123415, 123416, 123417, 123419, 123420, 123422, 123423, 123424, 123425, 123426, 123427, 123428, 123430, 123431, 123432, 123434, 123435, 123436, 123437, 123700, 123701, 123705, 123709, 123800, 200209, 123207
                            'lo que valida
                            Try
                                xlExcel = New Excel.Application
                                xlLibro = xlExcel.Workbooks.Open(OpenFileDialog1.FileName)
                                'HOJA NOMBRE
                                NOMBRE()
                                CONTROL_D()

                                If (ValidaAno = "2020") Then
                                    CargaLabel()
                                    D15()
                                    D16()
                                    MsgBox("Proceso Finalizado con Éxito", MsgBoxStyle.Information, "Fin proceso")

                                    Me.LBerrores.Text = Me.DataGridView1.RowCount - 1
                                    Me.LblHojaControl.Text = ValidaHojaControl

                                    xlLibro.Saved = True
                                    xlExcel.ActiveWorkbook.Close()
                                    xlExcel.Quit()

                                Else ' cuando no corresponde al año

                                    LimpiarLabel()

                                    MsgBox("El archivo seleccionado no corresponde al año en curso", MsgBoxStyle.Exclamation)

                                    xlLibro.Saved = True
                                    xlExcel.ActiveWorkbook.Close()
                                    xlExcel.Quit()
                                End If

                                'xlLibro.Saved = True
                                'xlExcel.ActiveWorkbook.Close()
                                'xlExcel.Quit()
                            Catch ex As Exception
                                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al leer archivo")
                                Return False
                            End Try
                        Case Else
                            MsgBox("El archivo seleccionado no corresponde a un establecimiento del servicio de salud osorno", MsgBoxStyle.Exclamation)
                    End Select

                End If
        End Select

        Return True
    End Function
    Sub COMPLETITUD_SERIEA()
        Select Case CodigoEstablec
            Case 123000 ' EJEMPLO REVISION 
                ProgressBar1.Maximum = 22
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM11()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM21()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM24()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM25()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM29()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM30()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM31()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123010 'SERVICIO SALUD
                ProgressBar1.Maximum = 6
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123011 'PRAIS
                ProgressBar1.Maximum = 3
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123012 'CLINICA DENTAL MOVIL 
                ProgressBar1.Maximum = 1
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123030 'DEPTO. ATENCION INTEGRAL FUNCIONARIO
                ProgressBar1.Maximum = 6
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123100 'HOSPITAL BASE OSORNO
                ProgressBar1.Maximum = 15
                Me.REM01()
                Me.REM02()
                Me.REM03()
                Me.REM04() ' original
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM11()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM21()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM25()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123101 'HOSPITAL PURRANQUE
                ProgressBar1.Maximum = 13
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM11()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM21()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM24()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM25()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123102 'HOSPITAL RIO NEGRO
                ProgressBar1.Maximum = 10
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)

                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)

                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)

                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM21()
                ProgressBar1.Increment(1)

                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM24()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)

                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)

                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)

                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123103 'HOSPITAL PUERTO OCTAY
                ProgressBar1.Maximum = 16
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM21()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123104 'HOSPITAL MISION SAN JUAN DE LA COSTA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123105 'HOSPITAL DEL PERPETUO SOCORRO DE QUILACAHUIN
                ProgressBar1.Maximum = 16
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM21()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123207 'CENTRO DE REHABILITACION MINUSVALIDO DE PURRANQUE
                ProgressBar1.Maximum = 1
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123300 'CESFAM DR. PEDRO JAUREGUI
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123301 'CESFAM DR. MARCELO LOPETEGUI
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123302 'CESFAM OVEJERIA
                ProgressBar1.Maximum = 14
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123303 'CESFAM RAHUE ALTO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123304 'CESFAM ENTRE LAGOS
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123305 'CESFAM SAN PABLO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123306 'CESFAM PAMPA ALEGRE
                ProgressBar1.Maximum = 14
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123307 'CESFAM PURRANQUE
                ProgressBar1.Maximum = 14
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123309 'CESFAM PRACTICANTE PABLO ARAYA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123310 'CESFAM QUINTO CENTENARIO
                ProgressBar1.Maximum = 14
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123311 'CESFAM BAHIA MANSA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123312 'CESFAM PUAUCHO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123402 'POSTA SALUD RURAL CUINCO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123404 'PSR PICHI DAMAS
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123406 'PSR PUYEHUE
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123407 'PSR DESAGUE RUPANCO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123408 'PSR ÑADI PICHI DAMAS
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123410 'PSR TRES ESTEROS
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123411 'PSR CORTE ALTO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123412 'PSR CRUCERO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123413 'PSR COLIGUAL
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123414 'PSR HUEYUSCA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123415 'PSR CONCORDIA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123416 'PSR COLONIA PONCE
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123417 'PSR LA NARANJA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123419 'PSR SAN PEDRO DE PURRANQUE
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123420 'PSR COLLIHUINCO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123422 'PSR RUPANCO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123423 'PSR CASCADAS
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123424 'PSR PIEDRAS NEGRAS
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123425 'PSR CANCURA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123426 'PSR PELLINADA  Me.REM01()
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123427 'PSR LA CALO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123428 'PSR COIHUECO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123430 'PSR PURREHUIN
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123431 'PSR ALEUCAPI
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123432 'PSR LA POZA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123434 'PSR HUILMA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123435 'PSR PUCOPIO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123436 'PSR CHANCO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123437 'PSR CURRIMAHUIDA
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123700 'CECOSF MURRINUMO
                ProgressBar1.Maximum = 12
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123701 'CECOSF MANUEL RODRIGUEZ
                ProgressBar1.Maximum = 12
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123705 'CECOF EL ENCANTO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 123709 'CECOF RIACHUELO
                ProgressBar1.Maximum = 15
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM08()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM23()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM28()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123800 'SAPU  DENTAL DR. P. JAUREGUI
                ProgressBar1.Maximum = 1
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 200209 ' COSAM RAHUE 
                ProgressBar1.Maximum = 4
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 200455
                ProgressBar1.Maximum = 13
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM31()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
            Case 200445
                ProgressBar1.Maximum = 13
                Me.REM01()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM02()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM03()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM04()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM05()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM06()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM07()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM09()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19a()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM19b()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM26()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM27()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.REM31()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        End Select

    End Sub
    Sub COMPLETITUD_SERIEP()
        Select Case CodigoEstablec
            Case 123000 'EJEMPLO
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P9()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P11()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)


            Case 123100 'Hospital Base Osorno
                ProgressBar1.Maximum = 5
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P11()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123102 'Hospital Rio Negro
                ProgressBar1.Maximum = 1
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123103 'Hospital Pto. Octay
                ProgressBar1.Maximum = 8
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123104 'Hospital Mision San Juan
                ProgressBar1.Maximum = 8
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123105 'Hospital Socorro de Quilacahuin
                ProgressBar1.Maximum = 8
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123300 'Cesfam Pedro Jauregui
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P10()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123301 'Cesfam Marcelo Lopetegui
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123302 'Cesfam Ovejeria
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123303 'Cesfam Rahue Alto
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P10()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123304 'Cesfam Entre Lagos
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P10()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123305 ' Cesfam San Pablo
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P10()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123306 ' Cesfam Pampa Alegre
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123307 'Cesfam Purranque
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P10()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123309 'Cesfam Pablo Araya
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P2()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P3()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P4()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P5()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P6()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P7()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P8()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P10()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
                Me.P12()
                ProgressBar1.Increment(1)
                LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

            Case 123310 'Cesfam Quinto Centenario
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P10()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123311 'Cesfam Bahia Mansa
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P10()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123312 'Cesfam Puaucho
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P10()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123402 'PSR Cuinco
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123404 'PSR Pichi Damas
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123406 'PSR Puyehue
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123407 'PSR Desague Rupanco
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123408 'PSR Ñadi Pichi Damas
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123410 'PSR Tres Esteros
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123411 'PSR Corte Alto
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123412 'PSR Crucero
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123413 'PSR Coligual
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123414 'PSR Hueyusca
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123415 'PSR Concordia
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123416 'PSR Colonia Ponce
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123417 'PSR La Naranja
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123419 'PSR San Pedro de Purranque
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123420 'PSR Collihuinco
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123422 'PSR Rupanco
                ProgressBar1.Maximum = 10
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P10()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123423 'PSR Cascada
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123424 'PSR Piedras Negras
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123425 'PSR Cancura
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123426 'PSR Pellinada
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123427 'PSR La Calo
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123428 'PSR Coihueco
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123430 'PSR Purrehuin
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123431 'PSR Aleucapi
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123432 'PSR La Poza
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123434 'PSR Huilma
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123435 'PSR Pucopio
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123436 'PSR Chanco
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123437 'PSR Currimahuida
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123700 'Cecosf Murrinumo
                ProgressBar1.Maximum = 8
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123701 'Cecosf Manuel Rodriguez
                ProgressBar1.Maximum = 8
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123705 'Cecosf El Encanto
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 123709 'Cecosf Riachuelo
                ProgressBar1.Maximum = 9
                Me.P1()
                ProgressBar1.Increment(1)
                Me.P2()
                ProgressBar1.Increment(1)
                Me.P3()
                ProgressBar1.Increment(1)
                Me.P4()
                ProgressBar1.Increment(1)
                Me.P5()
                ProgressBar1.Increment(1)
                Me.P6()
                ProgressBar1.Increment(1)
                Me.P7()
                ProgressBar1.Increment(1)
                Me.P8()
                ProgressBar1.Increment(1)
                Me.P12()
                ProgressBar1.Increment(1)

            Case 200209 'Cosam Rahue
                ProgressBar1.Maximum = 1
                Me.P6()
                ProgressBar1.Increment(1)

        End Select
    End Sub
    Sub COMPLETITUD_SERIEBM()

        ProgressBar1.Maximum = 2
        Me.BM18()
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        Me.BM18A()
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)



    End Sub
    Sub NOMBRE()
        xlHoja = xlLibro.Worksheets("NOMBRE")
        ' ==========================================================================================================================
        If (xlHoja.Range("B2").Value = "") Then ' si es igual a 0
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "NOMBRE", "COMUNA", "[ERROR]", "Falta ingresar AÑO en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaComuna = "SIN DATO"
        Else ' tiene datos
            ValidaComuna = xlHoja.Range("B2").Value
        End If
        ' ==========================================================================================================================
        If xlHoja.Range("B3").Value = "" Then
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "NOMBRE", "ESTABLE", "[ERROR]", "Falta ingresar el NOMBRE DE ESTABLECIMIENTO en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaEstable = "SIN DATO"
        Else
            ValidaEstable = xlHoja.Range("B3").Value ' ESTABLECIMIENTO
        End If
        ' ==========================================================================================================================
        If Not (xlHoja.Range("C3").Value = 0) Or (xlHoja.Range("D3").Value = 0) Or (xlHoja.Range("E3").Value = 0) Or (xlHoja.Range("F3").Value = 0) Or (xlHoja.Range("G3").Value = 0) Or (xlHoja.Range("H3").Value = 0) Then

            ValidaCodigo = xlHoja.Range("C3").Value & xlHoja.Range("D3").Value & xlHoja.Range("E3").Value & xlHoja.Range("F3").Value & xlHoja.Range("G3").Value & xlHoja.Range("H3").Value
        Else
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "CODIGO", "ESTABLECIMIENTO", "[ERROR]", "Falta ingresar el CODIGO ESTABLECIMIENTO en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaCodigo = 0

        End If

        ' ==========================================================================================================================
        If xlHoja.Range("B5").Value = "" Then
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "DEPENDENCIA", "ESTABLE", "[ERROR]", "Falta ingresar DEPENDENCIA en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaDependencia = "SIN DATO"
        Else
            ValidaDependencia = xlHoja.Range("B5").Value
        End If
        ' ==========================================================================================================================
        If xlHoja.Range("B6").Value = "" Then
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "MES", "MES", "[ERROR]", "Falta ingresar el MES en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaMes = "SIN DATO"
        Else
            ValidaMes = xlHoja.Range("B6").Value
        End If
        ' ==========================================================================================================================
        If xlHoja.Range("B7").Value = 0 Then
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "AÑO", "AÑO", "[ERROR]", "Falta ingresar el AÑO en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaAno = ""
        Else
            ValidaAno = CInt(xlHoja.Range("B7").Value)
            '  ValidaAno = CInt(ValidaAno)

        End If
        ' ==========================================================================================================================
        If xlHoja.Range("A9").Value = "" Then
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "VERSION", "DOCU", "[ERROR]", "Falta ingresar la VERSION DEL DOCUMENTO en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaVersion = "SIN DATOS"
        Else
            ValidaVersion = xlHoja.Range("A9").Value
        End If
        ' ==========================================================================================================================
        If xlHoja.Range("B17").Value = "" Then
            With Me.DataGridView1.Rows
                .Add("NOMBRE", "SERIE", "REM", "[ERROR]", "Falta ingresar la SERIE DEL REM en Hoja NOMBRE", "[ 0 ]")
            End With
            ValidaSerieREM = "SIN DATOS"
        Else
            ValidaSerieREM = xlHoja.Range("B17").Value
        End If

        xlHoja = Nothing
    End Sub
    Sub LimpiarLabel()
        Me.DataGridView1.Rows.Clear()
        Me.LBEstable.Text = ""
        Me.LBcodigo.Text = ""
        Me.LBSerie.Text = ""
        Me.LBcomuna.Text = ""
        Me.LBdependencia.Text = ""
        Me.LBmes.Text = ""
        Me.LBversion.Text = ""
        Me.LBaño.Text = ""
        Me.LblHojaControl.Text = ""
        Me.LBerrores.Text = ""
        BtnExportar.UseVisualStyleBackColor = True
        Me.LBLprogreso.Text = "0 %"
        ProgressBar1.Value = 0

    End Sub
    Sub CargaLabel()
        Me.LBversion.Text = ValidaVersion.ToUpper
        Me.LBSerie.Text = ValidaSerieREM
        Me.LBEstable.Text = ValidaEstable.ToUpper
        Me.LBmes.Text = ValidaMes
        Me.LBaño.Text = ValidaAno
        Me.LBcodigo.Text = ValidaCodigo
        Me.LBdependencia.Text = ValidaDependencia
        Me.LBcomuna.Text = ValidaComuna

        BtnExportar.Enabled = True
        BtnExportar.UseVisualStyleBackColor = False
        BtnSalir.UseVisualStyleBackColor = False
    End Sub
    Sub CONTROL_A()
        xlHoja = xlLibro.Worksheets("Contro")
        ValidaHojaControl = xlHoja.Range("D31").Value
        xlHoja = Nothing
    End Sub
    Sub CONTROL_BM()
        xlHoja = xlLibro.Worksheets("Control")
        ValidaHojaControl = xlHoja.Range("D9").Value
        xlHoja = Nothing
    End Sub
    Sub CONTROL_P()
        xlHoja = xlLibro.Worksheets("Contro")
        ValidaHojaControl = xlHoja.Range("D18").Value
        xlHoja = Nothing
    End Sub
    Sub CONTROL_D()
        xlHoja = xlLibro.Worksheets("CONTROL")
        ValidaHojaControl = xlHoja.Range("E13").Value
        xlHoja = Nothing
    End Sub
    Sub P1()
        Dim ii, B(139), C(139), D(139), E(139), F(139), G(139), H(139), I(139), J(139), K(139), L(139), M(139), N(139), O(139), P(139) As Integer
        xlHoja = xlLibro.Worksheets("P1")
        'SECCION A: POBLACIÓN EN CONTROL SEGÚN MÉTODO DE REGULACIÓN DE FERTILIDAD
        For ii = 11 To 33
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value

        Next
        'SECCION B: GESTANTES EN CONTROL  CON RIESGO PSICOSOCIAL 
        For ii = 37 To 46
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        'SECCION C: GESTANTES EN RIESGO PSICOSOCIAL CON VISITA DOMICILIARIA INTEGRAL REALIZADA EN EL SEMESTRE 
        For ii = 50 To 54
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
        Next
        'SECCION D: GESTANTES Y  MUJERES DE 8° MES POST-PARTO EN CONTROL, SEGÚN ESTADO NUTRICIONAL
        For ii = 58 To 67
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
        Next
        'SECCION E: MUJERES Y GESTANTES EN CONTROL CON CONSULTA NUTRICIONAL 
        For ii = 70 To 74
            B(ii) = xlHoja.Range("B" & ii & "").Value
        Next
        'SECCION F: MUJERES EN CONTROL DE CLIMATERIO 
        For ii = 77 To 81
            B(ii) = xlHoja.Range("B" & ii & "").Value
        Next
        'SECCIÓN G: POBLACIÓN EN CONTROL POR PATOLOGÍAS DE ALTO RIESGO OBSTÉTRICO
        For ii = 86 To 92
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
        Next
        'SECCION H: GESTANTES EN CONTROL  DE ENFERMEDADES TRANSMISIBLES (HEPATITIS B,  CHAGAS)
        For ii = 95 To 104
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next
        'SECCION H: GESTANTES EN CONTROL DE ENFERMEDADES TRANSMISIBLES (HEPATITIS B, CHAGAS)

        For ii = 108 To 117
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next
        'SECCIÓN I: POBLACIÓN EN CONTROL POR PATOLOGÍAS DE ALTO RIESGO OBSTÉTRICO
        For ii = 121 To 139
            B(ii) = xlHoja.Range("B" & ii & "").Value
        Next

        MsgBox("REMP1 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '  1 
        Select Case C(62)
            Case Is <> B(46)
                With Me.DataGridView1.Rows
                    .Add("P1", " [B][D]", "VAL [01]", "[ERROR]", "Gestantes y Mujeres de 8º Mes Post-Parto en Control, celda C62 debe ser igual a Gestantes en Control con Riesgo Psicosocial, celda B46", "[" & C(62) & " - " & B(46) & "]")
                End With
        End Select
        ' 2 VALIDACION LOCAL
        If (B(53) <> 0 And C(53) < 4) Or (B(53) = 0 And C(53) <> 0) Then
            With Me.DataGridView1.Rows
                .Add("P1", " [C]", "VAL [02]", "[REVISAR]", "Gestantes en Riesgo Psicosocial con visita Domiciliaria, si celda B53  tiene información se debe multiplicar por el total de visitas en celda C53", "[" & B(53) & " - " & C(53) & "]")
            End With
        End If

        '3 *********************************************************************************************
        Select Case B(71)  ' 
            Case Is > C(58)
                With Me.DataGridView1.Rows
                    .Add("REM P1", " [E][D]", "VAL [03]", "[ERROR]", "Mujeres y Gestantes en Control con Consulta Nutricional, la celda B71 debe ser menor o igual a la celda C58", "[" & B(71) & " - " & C(58) & "]")
                End With
        End Select
        '4 *********************************************************************************************
        Select Case B(72)  ' 
            Case Is > C(59)
                With Me.DataGridView1.Rows
                    .Add("P1", " [E]", "VAL [04]", "[ERROR]", "Mujeres y Gestantes en Control con Consulta Nutricional, la celda B72 debe ser menor o igual a la celda C59", "[" & B(72) & " - " & C(59) & "]")
                End With
        End Select
        '5 *********************************************************************************************
        Select Case B(70)  ' 
            Case Is > C(61)
                With Me.DataGridView1.Rows
                    .Add("P1", " [E]", "VAL [05]", "[ERROR]", "Mujeres y Gestantes en Control con Consulta Nutricional, la celda B70 debe ser menor o igual a la celda C61", "[" & B(70) & " - " & C(61) & "]")
                End With
        End Select

        '6 *********************************************************************************************
        Select Case B(77) ' 
            Case Is < B(78)
                With Me.DataGridView1.Rows
                    .Add("P1", " [F]", "VAL [06]", "[ERROR]", "Mujeres en Control de Climaterio, celda B77 debe ser mayor o igual a la celda B78", "[" & B(77) & " - " & B(78) & "]")
                End With
        End Select
        '7 *********************************************************************************************
        Select Case B(78) ' 
            Case Is < B(79)
                With Me.DataGridView1.Rows
                    .Add("P1", " [F]", "VAL [07]", "[ERROR]", "Mujeres en Control de Climaterio, celda B78 debe ser mayor o igual a la celda B79", "[" & B(78) & " - " & B(79) & "]")
                End With
        End Select
        '8 *********************************************************************************************
        Select Case B(79) ' 
            Case Is < B(80)
                With Me.DataGridView1.Rows
                    .Add("P1", " [F]", "VAL [08]", "[ERROR]", "Mujeres en Control de Climaterio, celda B79 debe ser mayor o igual a la celda B80", "[" & B(79) & " - " & B(80) & "]")
                End With
        End Select
        '9 *********************************************************************************************
        Select Case CodigoEstablec
            Case 123100 ' HBO
            Case Else ' el resto de establecimiento
                Select Case B(139)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("P1", " [I]", "VAL [09]", "[ERROR]", "Población en Control por Patologías de Alto Riesgo Obstétrico, celda B139 corresponde sólo a Hospital Base Osorno", "[" & B(139) & "]")
                        End With
                End Select
        End Select

        xlHoja = Nothing
    End Sub
    Sub P2()
        Dim ii, B(146), C(146), D(146), E(146), F(146), G(146), H(146), I(146), J(146), K(146), L(146), M(146), N(146), O(146), P(146), Q(146), R(146), S(146), T(146), U(146), V(146), W(146), X(146), Y(146), Z(146), AA(146), AB(146), AC(146), AD(146), AE(146), AF(146), AG(146), AH(146), AI(146), AJ(146), AK(146), AL(146), AM(146), AN(146), AO(146), AP(146), AQ(146) As Integer
        xlHoja = xlLibro.Worksheets("P2")
        'SECCION A: POBLACIÓN EN CONTROL, SEGÚN ESTADO NUTRICIONAL PARA NIÑOS MENOR DE UN MES-59 MESES
        For ii = 11 To 39
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
        Next
        'SECCION A.1: POBLACIÓN EN CONTROL, SEGÚN ESTADO NUTRICIONAL PARA NIÑOS DE 60 MESES-9 AÑOS 11 MESES												
        For ii = 44 To 72
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
        Next
        'SECCION B: POBLACION EN CONTROL SEGÚN RESULTADO DE EVALUACIÓN DEL DESARROLLO PSICOMOTOR
        For ii = 76 To 85
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next
        'SECCION C: POBLACIÓN MENOR DE 1 AÑO EN CONTROL, SEGÚN SCORE RIESGO EN IRA Y VISITA DOMICILIARIA INTEGRAL EN EL SEMESTRE
        For ii = 89 To 92
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
        Next
        'SECCION D: POBLACIÓN EN CONTROL EN EL SEMESTRE CON CONSULTA NUTRICIONAL, SEGÚN ESTRATEGIA
        For ii = 96 To 97
            C(ii) = xlHoja.Range("C" & ii & "").Value
        Next
        'SECCION E: POBLACIÓN INASISTENTE A CONTROL DEL NIÑO SANO (AL CORTE)
        For ii = 101 To 110
            C(ii) = xlHoja.Range("C" & ii & "").Value
        Next
        'SECCION F: POBLACIÓN INFANTIL SEGÚN DIAGNÓSTICO DE PRESIÓN ARTERIAL (Incluida en sección A y A1)
        For ii = 114 To 118
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        'SECCION G: POBLACIÓN INFANTIL EUTRÓFICA, SEGÚN RIESGO DE MALNUTRICIÓN POR EXCESO (Incluida en sección A y A1)
        For ii = 122 To 124
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
        Next
        'SECCION H: POBLACIÓN EN CONTROL, SEGÚN EVALUACIÓN DE RIESGO ODONTOLÓGICO
        For ii = 129 To 146
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        Next

        MsgBox("REMP2 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        ' 1 
        Select Case (C(11) - (F(11) + G(11)))
            Case Is <> C(39)
                With Me.DataGridView1.Rows
                    .Add("P2", " [A]", "VAL [01]", "[ERROR]", "Población en Control, la suma de la celda C11 menos la suma de las celdas F11 y G11 debe ser igual a la celda C39", "[" & (C(11) - (F(11) + G(11))) & " - " & C(39) & "]")
                End With
        End Select
        '2
        Select Case C(19)
            Case Is <> (H(33) + I(33) + J(33) + K(33) + L(33) + M(33) + N(33) + O(33) + P(33) + Q(33) + R(33) + S(33) + T(33) + U(33) + V(33) + W(33) + X(33) + Y(33) + Z(33) + AA(33) + AB(33) + AC(33) + AD(33) + AE(33) + AF(33) + AG(33))
                With Me.DataGridView1.Rows
                    .Add("P2", " [A]", "VAL [02]", "[ERROR]", "Población en Control, celda C19 debe ser igual a la suma de las celdas H33 a AG33", "[" & C(19) & " - " & (H(33) + I(33) + J(33) + K(33) + L(33) + M(33) + N(33) + O(33) + P(33) + Q(33) + R(33) + S(33) + T(33) + U(33) + V(33) + W(33) + X(33) + Y(33) + Z(33) + AA(33) + AB(33) + AC(33) + AD(33) + AE(33) + AF(33) + AG(33)) & "]")
                End With
        End Select
        '3 
        Select Case C(18)
            Case Is <> (H(34) + I(34) + J(34) + K(34) + L(34) + M(34) + N(34) + O(34) + P(34) + Q(34) + R(34) + S(34) + T(34) + U(34) + V(34) + W(34) + X(34) + Y(34) + Z(34) + AA(34) + AB(34) + AC(34) + AD(34) + AE(34) + AF(34) + AG(34))
                With Me.DataGridView1.Rows
                    .Add("P2", " [A]", "VAL [03]", "[ERROR]", "Población en Control, celda C18 debe ser igual a la suma de H34 a AG34", "[" & C(15) & " - " & (H(34) + I(34) + J(34) + K(34) + L(34) + M(34) + N(34) + O(34) + P(34) + Q(34) + R(34) + S(34) + T(34) + U(34) + V(34) + W(34) + X(34) + Y(34) + Z(34) + AA(34) + AB(34) + AC(34) + AD(34) + AE(34) + AF(34) + AG(34)) & "]")
                End With
        End Select
        '4 
        Select Case C(32)
            Case Is <> (((H(16) + I(16) + J(16) + K(16) + L(16) + M(16) + N(16) + O(16) + P(16) + Q(16) + R(16) + S(16) + T(16) + U(16)) + (V(22) + W(22) + X(22) + Y(22) + Z(22) + AA(22) + AB(22) + AC(22) + AD(22) + AE(22) + AF(22) + AG(22))) - C(38))
                With Me.DataGridView1.Rows
                    .Add("P2", " [A]", "VAL [04]", "[ERROR]", "Población en Control, celda C32 debe ser igual a la suma de las celdas H16 a U16 más la suma  de las celdas V22 a AG22 menos C38", "[" & C(32) & " - " & (((H(16) + I(16) + J(16) + K(16) + L(16) + M(16) + N(16) + O(16) + P(16) + Q(16) + R(16) + S(16) + T(16) + U(16)) + (V(22) + W(22) + X(22) + Y(22) + Z(22) + AA(22) + AB(22) + AC(22) + AD(22) + AE(22) + AF(22) + AG(22))) - C(38)) & "]")
                End With
        End Select



        xlHoja = Nothing
    End Sub
    Sub P3()
        Dim ii, B(59), C(59), D(59), E(59), F(59), G(59), H(59), I(59), J(59), K(59), L(59), M(59), N(59), O(59), P(59), Q(59), R(59), S(59), T(59), U(59), V(59), W(59), X(59), Y(59), Z(59), AA(59), AB(59), AC(59), AD(59), AE(59), AF(59), AG(59), AH(59), AI(59), AJ(59), AK(59), AL(59), AM(59), AN(59), AO(59), AP(59), AQ(59), AR(59) As Integer
        xlHoja = xlLibro.Worksheets("P3")
        'SECCION A: EXISTENCIA DE POBLACIÓN EN CONTROL
        For ii = 12 To 38
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
        Next
        'SECCION B: CUIDADORES DE PACIENTES CON DEPENDENCIA SEVERA																																									
        For ii = 43 To 43
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        'SECCION C: POBLACIÓN EN CONTROL EN PROGRAMA DE REHABILITACIÓN PULMONAR EN SALA IRA-ERA																																						
        For ii = 48 To 48
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
        Next
        'SECCION D: NIVEL DE CONTROL DE POBLACION RESPIRATORIA CRONICA																																						
        For ii = 53 To 59
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
        Next
        MsgBox("REMP3 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1 ***********************************************************
        Select Case C(34)
            Case Is > (C(32) + C(33))
                With Me.DataGridView1.Rows
                    .Add("P3", " [A]", "VAL [01]", "[ERROR]", "Existencia de Población en Control, celda C34 debe ser menor o igual a la suma de las celdas C32 y C33", "[" & C(34) & " - " & (C(33) + C(33)) & "]")
                End With
        End Select
        '2 ***********************************************************
        Select Case C(36)
            Case Is > C(34)
                With Me.DataGridView1.Rows
                    .Add("P3", " [A]", "VAL [02]", "[ERROR]", "Existencia de Población en Control, celda C36 debe ser menor o igual a celda C34", "[" & C(36) & " - " & C(34) & "]")
                End With
        End Select
        '3 ***********************************************************
        Select Case (C(15) + C(16) + C(17))
            Case Is <> (C(53) + C(54) + C(55) + C(56))
                With Me.DataGridView1.Rows
                    .Add("P3", " [A][D]", "VAL [05]", "[ERROR]", "Existencia de Población en Control, suma de celdas C15 a C17 debe ser igual a la suma de Control de Población Respiratoria Crónica de las celdas C53 a C56", "[" & (C(15) + C(16) + C(17)) & " - " & (C(53) + C(54) + C(55) + C(56)) & "]")
                End With
        End Select

        '4 ***********************************************************
        Select Case (C(18) + C(19))
            Case Is <> (C(57) + C(58) + C(59))
                With Me.DataGridView1.Rows
                    .Add("P3", " [A][D]", "VAL [06]", "[ERROR]", "Existencia de Población en Control, suma de celdas C18 a C19 debe ser igual a la suma de Control de Población Respiratoria Crónica de las celdas C57 a C59", "[" & (C(18) + C(19)) & " - " & (C(57) + C(58) + C(59)) & "]")
                End With
        End Select


        xlHoja = Nothing
    End Sub
    Sub P4()
        Dim ii, B(87), C(87), D(87), E(87), F(87), G(87), H(87), I(87), J(87), K(87), L(87), M(87), N(87), O(87), P(87), Q(87), R(87), S(87), T(87), U(87), V(87), W(87), X(87), Y(87), Z(87), AA(87), AB(87), AC(87), AD(87), AE(87), AF(87), AG(87), AH(87), AI(87), AJ(87), AK(87), AL(87), AM(87), AN(87), AO(87), AP(87), AQ(87), AR(87) As Integer
        xlHoja = xlLibro.Worksheets("P4")
        ' SECCIÓN A: PROGRAMA SALUD CARDIOVASCULAR (PSCV)
        For ii = 12 To 28
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
        Next
        ' SECCIÓN B: METAS DE COMPENSACIÓN	
        For ii = 33 To 40
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
        Next
        ' SECCIÓN C: VARIABLES DE SEGUIMIENTO DEL PSCV AL CORTE
        For ii = 45 To 64
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        For ii = 66 To 67
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        For ii = 69 To 74
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        'SECCIÓN D: POBLACION EN CONTROL (EMPA VIGENTE) SEGÚN COMPENSACION Y ESTADO NUTRICIONAL 
        For ii = 81 To 87
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
        Next
        MsgBox("REMP4 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1 **********************************************
        Select Case C(12)
            Case Is > (C(16) + C(17) + C(18))
                With Me.DataGridView1.Rows
                    .Add("P4", " [A]", "VAL [01]", "[ERROR]", "Programa Salud Cardiovascular, celda C12 debe ser menor o igual a la suma de las celdas C16 a C18", "[" & C(12) & " - " & (C(16) + C(17) + C(18)) & "]")
                End With
        End Select
        ' 2 **********************************************
        Select Case C(12)
            Case Is < C(28)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A]", "VAL [02]", "[ERROR]", "Programa Salud Cardiovascular, celda C12 debe ser mayor o igual a C28", "[" & C(12) & " - " & C(28) & "]")
                End With
        End Select
        '3 **********************************************
        Select Case AH(28)
            Case Is > AH(12)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A]", "VAL [03]", "[REVISAR]", "Programa Salud Cardiovascular, celda AH28  debe ser menor o igual a AH12", "[" & AH(28) & " - " & AH(12) & "]")
                End With
        End Select
        '4 **********************************************
        Select Case AI(28)
            Case Is > AI(12)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A]", "VAL [04]", "[REVISAR]", "Programa Salud Cardiovascular, celda AI28 debe ser menor o igual a AI12", "[" & AI(28) & " - " & AI(12) & "]")
                End With
        End Select
        '5 **********************************************
        Select Case AJ(28)
            Case Is > AJ(12)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A]", "VAL [05]", "[REVISAR]", "Programa Salud Cardiovascular, celda AJ28 debe ser menor o igual a AJ12", "[" & AJ(28) & " - " & AJ(12) & "]")
                End With
        End Select
        '6 **********************************************
        Select Case AK(28)
            Case Is > AK(12)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A]", "VAL [06]", "[REVISAR]", "Programa Salud Cardiovascular, celda AK28 debe ser menor o igual a AK12", "[" & AK(28) & " - " & AK(12) & "]")
                End With
        End Select
        '7 **********************************************
        Select Case C(33)
            Case Is > C(16)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A][B]", "VAL [07]", "[ERROR]", "Prog. Salud CV, celda C33 debe ser menor o igual a Metas, celda C16", "[" & C(33) & " - " & C(16) & "]")
                End With
        End Select
        '8 **********************************************
        Select Case C(38)
            Case Is > C(18)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A][B]", "VAL [08]", "[ERROR]", "Prog. Salud CV, celda C38 debe ser menor o igual a Metas, celda C15", "[" & C(38) & " - " & C(18) & "]")
                End With
        End Select
        '9 **********************************************
        Select Case C(35)
            Case Is > C(17)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A][B]", "VAL [09]", "[ERROR]", "Prog. Salud CV, celda C35 debe ser menor o igual a Metas, celda C17", "[" & C(35) & " - " & C(17) & "]")
                End With
        End Select
        '10 **********************************************
        Select Case C(37)
            Case Is > C(17)
                With Me.DataGridView1.Rows
                    .Add("P4", " [A][B]", "VAL [10]", "[ERROR]", "Prog. Salud CV, celda C37 debe ser menor o igual a Metas, celda C17", "[" & C(37) & " - " & C(17) & "]")
                End With
        End Select
        '11 **********************************************
        Select Case C(39)
            Case Is > (C(20) + C(21))
                With Me.DataGridView1.Rows
                    .Add("P4", " [A][B]", "VAL [11]", "[ERROR]", "Metas de Compensación, celda C39 debe ser menor o igual a Prog. Salud Cardiovascular, suma de celdas C20 y C21", "[" & C(39) & " - " & (C(20) + C(21)) & "]")
                End With
        End Select
        '12 **********************************************
        Select Case C(40)
            Case Is > (C(20) + C(21))
                With Me.DataGridView1.Rows
                    .Add("P4", " [A][B]", "VAL [12]", "[ERROR]", "Metas de Compensación, celda C40 debe ser menor o igual a Prog. Salud Cardiovascular, suma de celdas C20 y C21", "[" & C(40) & " - " & (C(20) + C(21)) & "]")
                End With
        End Select


        xlHoja = Nothing
    End Sub
    Sub P5()
        Dim ii, B(57), C(57), D(57), E(57), F(57), G(57), H(57), I(57), J(57), K(57), L(57), M(57), N(57), O(57), P(57), Q(57), R(57), S(57), T(57), U(57), V(57), W(57), X(57), Y(57), Z(57), AA(57), AB(57), AC(57), AD(57), AE(57), AF(57), AG(57), AH(57), AI(57), AJ(57), AK(57), AL(57), AM(57), AN(57), AO(57), AP(57), AQ(57), AR(57) As Integer
        xlHoja = xlLibro.Worksheets("P5")
        ' SECCION A:  POBLACIÓN EN CONTROL POR CONDICIÓN DE FUNCIONALIDAD
        For ii = 12 To 21
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
        Next
        ' SECCION A.1: EXISTENCIA DE POBLACIÓN EN CONTROL EN PROGRAMA "MÁS ADULTOS MAYORES AUTOVALENTES" POR CONDICIÓN DE FUNCIONALIDAD
        For ii = 26 To 29
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
        Next
        ' SECCION B: POBLACIÓN BAJO CONTROL POR ESTADO NUTRICIONAL
        For ii = 34 To 38
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        ' SECCION C: ADULTOS MAYORES CON SOSPECHA DE MALTRATO						
        For ii = 43 To 43
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
        Next
        ' SECCION D: ADULTOS MAYORES  EN ACTIVIDAD FÍSICA 							
        For ii = 48 To 48
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
        Next
        ' SECCION E: ADULTOS MAYORES CON RIESGO DE CAÍDAS								
        For ii = 53 To 57
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
        Next
        MsgBox("REMP5 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1 ***************************************
        Select Case D(21)
            Case Is <> C(38)
                With Me.DataGridView1.Rows
                    .Add("P5", " [A][B]", "VAL [01]", "[ERROR]", "Población en Control por Condición de Funcionalidad, celda D21 debe ser igual a Población Bajo Control por Estado Nutricional, celda C38", "[" & D(21) & " - " & C(38) & "]")
                End With
        End Select
        Select Case E(21)
            Case Is <> D(38)
                With Me.DataGridView1.Rows
                    .Add("P5", " [A][B]", "VAL [01]", "[ERROR]", "Población en Control por Condición de Funcionalidad, celda E21 debe ser igual a Población Bajo Control por Estado Nutricional, celda D38", "[" & E(21) & " - " & D(38) & "]")
                End With
        End Select
        Select Case F(21)
            Case Is <> E(38)
                With Me.DataGridView1.Rows
                    .Add("P5", " [A][B]", "VAL [01]", "[ERROR]", "Población en Control por Condición de Funcionalidad, celda F21 debe ser igual a Población Bajo Control por Estado Nutricional, celda E38", "[" & F(21) & " - " & E(38) & "]")
                End With
        End Select
        '2 VALIDACION LOCAL **************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123307, 123300, 123301, 123303, 123306, 123310, 123304 ' Más Adultos Mayores Autovalentes
            Case Else ' resto de estable
                Select Case D(29)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("P5", " [A.1]", "VAL [02]", "[ERROR]", "Existencia de Población en Control en Programa MAS, celda D29 solo debe ingresar establecimientos pertenecientes al Programa Más Adultos Autovalentes", D(29))
                        End With
                End Select
        End Select


        xlHoja = Nothing
    End Sub
    Sub P6()
        Dim ii, B(111), C(111), D(111), E(111), F(111), G(111), H(111), I(111), J(111), K(111), L(111), M(111), N(111), O(111), P(111), Q(111), R(111), S(111), T(111), U(111), V(111), W(111), X(111), Y(111), Z(111), AA(111), AB(111), AC(111), AD(111), AE(111), AF(111), AG(111), AH(111), AI(111), AJ(111), AK(111), AL(111), AM(111), AN(111), AO(111), AP(111), AQ(111), AR(111), AS1(111), AT(111), AU(111) As Integer
        xlHoja = xlLibro.Worksheets("P6")
        'SECCION A.1: POBLACIÓN EN CONTROL EN APS AL CORTE
        For ii = 13 To 13
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
        Next
        For ii = 15 To 51
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
        Next
        ' SECCION A.2: PROGRAMA DE REHABILITACIÓN EN ATENCION PRIMARIA (PERSONAS CON TRANSTORNOS PSIQUIÁTRICO).
        For ii = 56 To 57
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
        Next
        ' SECCION A.3: PROGRAMA DE ACOMPAÑAMIENTO PSICOSOCIAL EN LA ATENCION PRIMARIA
        For ii = 62 To 62
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        ' B. ATENCIÓN DE ESPECIALIDADES
        ' SECCION B.1: POBLACIÓN EN CONTROL EN ESPECIALIDAD AL CORTE"														
        For ii = 67 To 67
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
        Next
        For ii = 69 To 105
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
        Next
        ' SECCION B.2: PROGRAMA DE REHABILITACIÓN EN ESPECIALIDAD (PERSONAS CON TRANSTORNOS PSIQUIÁTRICO)
        For ii = 110 To 111
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
        Next
        MsgBox("REMP6 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1********************************************************************************************************************************************************
        Select Case C(13)
            Case Is < (C(15) + C(16) + C(17))
                With Me.DataGridView1.Rows
                    .Add("P6", " [A.1]", "VAL [01]", "[ERROR]", "Población en control al corte, celdas C13 debe ser mayor o igual a la suma de las celdas C15 a C17", "[" & C(13) & " - " & (C(15) + C(16) + C(17)) & "]")
                End With
        End Select
        '2********************************************************************************************************************************************************
        Select Case C(22)
            Case Is < C(13)
                With Me.DataGridView1.Rows
                    .Add("P6", " [A.1]", "VAL [02]", "[ERROR]", "Población en control al corte, personas con diagnostico de trastorno mentales celda C22 deben estar incluidas en el numero de personas en control celda C13", "[" & C(22) & " - " & C(13) & "]")
                End With
        End Select
        '3********************************************************************************************************************************************************
        Select Case C(22)
            Case Is > (C(23) + C(24) + C(25) + C(26) + C(27) + C(28) + C(29) + C(30) + C(31) + C(32) + C(33) + C(34) + C(35) + C(36) + C(37) + C(38) + C(39) + C(40) + C(41) + C(42) + C(43) + C(44) + C(45) + C(46) + C(47) + C(48) + C(49) + C(50) + C(51))
                With Me.DataGridView1.Rows
                    .Add("P6", " [A.1]", "VAL [03]", "[ERROR]", "Población en control al corte, celda C22 debe ser menor o igual a la suma de las celdas C23:C51", "[" & C(22) & " - " & (C(23) + C(24) + C(25) + C(26) + C(27) + C(28) + C(29) + C(30) + C(31) + C(32) + C(33) + C(34) + C(35) + C(36) + C(37) + C(38) + C(39) + C(40) + C(41) + C(42) + C(43) + C(44) + C(45) + C(46) + C(47) + C(48) + C(49) + C(50) + C(51)) & "]")
                End With
        End Select

        '4********************************************************************************************************************************************************
        Select Case C(67)
            Case Is < (C(69) + C(70) + C(71))
                With Me.DataGridView1.Rows
                    .Add("P6", " [B.1]", "VAL [04]", "[ERROR]", "Población en control al corte, celdas C67 debe ser mayor o igual a la suma de las celdas C69 a C71", "[" & C(67) & " - " & (C(69) + C(70) + C(71)) & "]")
                End With
        End Select
        '5********************************************************************************************************************************************************
        Select Case C(76)
            Case Is < C(67)
                With Me.DataGridView1.Rows
                    .Add("P6", " [B.1]", "VAL [05]", "[ERROR]", "Población en control al corte, personas con diagnostico de trastorno mentales celda C76 deben estar incluidas en el numero de personas en control celda C67", "[" & C(76) & " - " & C(67) & "]")
                End With
        End Select
        '6********************************************************************************************************************************************************
        Select Case C(76)
            Case Is > (C(77) + C(78) + C(79) + C(80) + C(81) + C(82) + C(83) + C(84) + C(85) + C(86) + C(87) + C(88) + C(89) + C(90) + C(91) + C(92) + C(93) + C(94) + C(95) + C(96) + C(97) + C(98) + C(99) + C(100) + C(101) + C(102) + C(103) + C(104) + C(105))
                With Me.DataGridView1.Rows
                    .Add("P6", " [B.1]", "VAL [06]", "[ERROR]", " Población en control al corte, celda C76 debe ser menor o igual a la suma de las celdas C77:C105", "[" & C(76) & " - " & (C(77) + C(78) + C(79) + C(80) + C(81) + C(82) + C(83) + C(84) + C(85) + C(86) + C(87) + C(88) + C(89) + C(90) + C(91) + C(92) + C(93) + C(94) + C(95) + C(96) + C(97) + C(98) + C(99) + C(100) + C(101) + C(102) + C(103) + C(104) + C(105)) & "]")
                End With
        End Select

        xlHoja = Nothing
    End Sub
    Sub P7()
        Dim ii, B(34), C(34), D(34), E(34), F(34), G(34), H(34), I(34), J(34), K(34), L(34) As Integer
        xlHoja = xlLibro.Worksheets("P7")
        ' SECCIÓN A. CLASIFICACIÓN DE LAS FAMILIAS SECTOR URBANO
        For ii = 10 To 14
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
        Next
        ' SECCIÓN A.1 CLASIFICACIÓN DE LAS FAMILIAS SECTOR RURAL
        For ii = 18 To 22
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
        Next
        ' SECCIÓN B. INTERVENCIÓN EN FAMILIAS SECTOR URBANO Y RURAL
        For ii = 25 To 31
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
        Next
        MsgBox("REMP7 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1**************************************************************************************************************************************
        Select Case B(10)
            Case Is < B(11)
                With Me.DataGridView1.Rows
                    .Add("P7", " [A]", "VAL [01]", "[ERROR]", "Clasificación de las familias, celdas B10 debe ser mayor B11", "[" & B(10) & " - " & B(11) & "]")
                End With
        End Select
        '2**************************************************************************************************************************************
        Select Case B(11)
            Case Is < (B(12) + B(13) + B(14))
                With Me.DataGridView1.Rows
                    .Add("P7", " [A]", "VAL [02]", "[ERROR]", "Clasificación de las familias Sector Urbano, celda B12 debe ser mayor a la celdas B12:B14", "[" & B(11) & " - " & (B(12) + B(13) + B(14)) & "]")
                End With
        End Select
        '3**************************************************************************************************************************************
        Select Case B(18)
            Case Is < B(19)
                With Me.DataGridView1.Rows
                    .Add("P7", " [A.1]", "VAL [03]", "[ERROR]", "Clasificación de las familias Sector Rural, celda B18 debe ser mayor a la celda B19", "[" & B(18) & " - " & B(19) & "]")
                End With
        End Select
        '4**************************************************************************************************************************************
        Select Case B(19)
            Case Is < (B(20) + B(21) + B(22))
                With Me.DataGridView1.Rows
                    .Add("P7", " [A.1]", "VAL [04]", "[ERROR]", "Clasificación de las familias Sector Urbano, celda B19 debe ser mayor a la celdas B20:B22", "[" & B(19) & " - " & (B(20) + B(21) + B(22)) & "]")
                End With
        End Select




        xlHoja = Nothing
    End Sub
    Sub P8()
        'Dim ii, C(30), D(30), E(30), F(30), G(30), H(30), I(30), J(30), K(30), L(30), M(30), N(30), O(30), P(30), Q(30), R(30), S(30), T(30), U(30), V(30), W(30), X(30), Y(30), Z(30), AA(30), AB(30), AC(30), AD(30), AE(30), AF(30), AG(30), AH(30), AI(30), AJ(30), AK(30), AL(30), AM(30), AN(30), AO(30), AP(30), AQ(30), AR(30), AS1(30), AT(30) As Integer
        'xlHoja = xlLibro.Worksheets("P8")
        ''SECCION A
        'For ii = 12 To 30
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        '    AA(ii) = xlHoja.Range("AA" & ii & "").Value
        '    AB(ii) = xlHoja.Range("AB" & ii & "").Value
        '    AC(ii) = xlHoja.Range("AC" & ii & "").Value
        '    AD(ii) = xlHoja.Range("AD" & ii & "").Value
        '    AE(ii) = xlHoja.Range("AE" & ii & "").Value
        '    AF(ii) = xlHoja.Range("AF" & ii & "").Value
        '    AG(ii) = xlHoja.Range("AG" & ii & "").Value
        '    AH(ii) = xlHoja.Range("AH" & ii & "").Value
        '    AI(ii) = xlHoja.Range("AI" & ii & "").Value
        '    AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        '    AK(ii) = xlHoja.Range("AK" & ii & "").Value
        '    AL(ii) = xlHoja.Range("AL" & ii & "").Value
        '    AM(ii) = xlHoja.Range("AM" & ii & "").Value
        '    AN(ii) = xlHoja.Range("AN" & ii & "").Value
        '    AO(ii) = xlHoja.Range("AO" & ii & "").Value
        '    AP(ii) = xlHoja.Range("AP" & ii & "").Value
        '    AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
        '    AR(ii) = xlHoja.Range("AR" & ii & "").Value
        '    AS1(ii) = xlHoja.Range("AS" & ii & "").Value
        '    AT(ii) = xlHoja.Range("AT" & ii & "").Value
        'Next
        ''*************************************************************************************************************************************************************************************
        ''********************************************************************** VALIDACIONES *************************************************************************************************
        ''*************************************************************************************************************************************************************************************
        ''1**************************************************************************************************************************************************
        ''Select Case D(12)
        ''    Case Is > (D(13) + D(14) + D(15) + D(16) + D(17) + D(18) + D(19) + D(20) + D(21) + D(22) + D(23) + D(24) + D(25) + D(26) + D(27) + D(28) + D(29) + D(30))
        ''        With Me.DataGridView1.Rows
        ''            .Add("P8", " [A]", "VAL [01]", "[ERROR]", "Existencia de Población en Control, celda D12 debe ser menor o igual a la celda D13 a D30", "[" & D(12) & " - " & (D(13) + D(14) + D(15) + D(16) + D(17) + D(18) + D(19) + D(20) + D(21) + D(22) + D(23) + D(24) + D(25) + D(26) + D(27) + D(28) + D(29) + D(30)) & "]")
        ''        End With
        ''End Select

        'xlHoja = Nothing
    End Sub
    Sub P9()
        Dim ii, B(95), C(95), D(95), E(95), F(95), G(95), H(95), I(95), J(95), K(95), L(95), M(95), N(95), O(95), P(95), Q(95), R(95), S(95), T(95), U(95), V(95), W(95), X(95), Y(95), Z(95), AA(95), AB(95), AC(95), AD(95), AE(95), AF(95), AG(95), AH(95), AI(95), AJ(95), AK(95), AL(95), AM(95), AN(95), AO(95), AP(95), AQ(95), AR(95), AS1(95) As Integer
        xlHoja = xlLibro.Worksheets("P9")
        ' SECCION A: POBLACIÓN EN CONTROL DE SALUD INTEGRAL DE ADOLESCENTES, SEGÚN ESTADO NUTRICIONAL
        For ii = 12 To 40
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        ' SECCION B: POBLACIÓN EN CONTROL SALUD INTEGRAL DE ADOLESCENTES, SEGÚN  EDUCACIÓN Y TRABAJO
        For ii = 46 To 51
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        ' SECCION C: POBLACIÓN EN CONTROL SALUD INTEGRAL DE ADOLESCENTES, SEGÚN  ÁREAS DE RIESGO
        For ii = 57 To 62
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        ' SECCION D:POBLACIÓN EN CONTROL SALUD INTEGRAL DE ADOLESCENTES, SEGÚN AMBITOS GINECO-UROLOGICO/SEXUALIDAD
        For ii = 68 To 77
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
        Next
        ' SECCION E:POBLACIÓN ADOLESCENTE QUE RECIBE CONSEJERÍA
        For ii = 82 To 95
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
        Next
        MsgBox("REMP9 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1***********************************************************************************************************************************************************************************

        xlHoja = Nothing
    End Sub
    Sub P10()
        'Dim ii, D(71), E(71), F(71), G(71), H(71), I(71), J(71), K(71), L(71), M(71), N(71), O(71), P(71), Q(71), R(71), S(71), T(71), U(71), V(71), W(71), X(71), Y(71), Z(71) As Integer
        'xlHoja = xlLibro.Worksheets("P10")
        ''SECCION A
        'For ii = 12 To 24
        '    D(ii) = xlHoja.Range("D" & ii & "").Value
        '    E(ii) = xlHoja.Range("E" & ii & "").Value
        '    F(ii) = xlHoja.Range("F" & ii & "").Value
        '    G(ii) = xlHoja.Range("G" & ii & "").Value
        '    H(ii) = xlHoja.Range("H" & ii & "").Value
        '    I(ii) = xlHoja.Range("I" & ii & "").Value
        '    J(ii) = xlHoja.Range("J" & ii & "").Value
        '    K(ii) = xlHoja.Range("K" & ii & "").Value
        '    L(ii) = xlHoja.Range("L" & ii & "").Value
        '    M(ii) = xlHoja.Range("M" & ii & "").Value
        '    N(ii) = xlHoja.Range("N" & ii & "").Value
        '    O(ii) = xlHoja.Range("O" & ii & "").Value
        '    P(ii) = xlHoja.Range("P" & ii & "").Value
        '    Q(ii) = xlHoja.Range("Q" & ii & "").Value
        '    R(ii) = xlHoja.Range("R" & ii & "").Value
        '    S(ii) = xlHoja.Range("S" & ii & "").Value
        '    T(ii) = xlHoja.Range("T" & ii & "").Value
        '    U(ii) = xlHoja.Range("U" & ii & "").Value
        '    V(ii) = xlHoja.Range("V" & ii & "").Value
        '    W(ii) = xlHoja.Range("W" & ii & "").Value
        '    X(ii) = xlHoja.Range("X" & ii & "").Value
        '    Y(ii) = xlHoja.Range("Y" & ii & "").Value
        '    Z(ii) = xlHoja.Range("Z" & ii & "").Value
        'Next

        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1*****************************************************************************************************************************************************************
        'Select Case D(13)
        '    Case Is < D(16)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [01]", "[ERROR]", "Población en Control, celda D13 debe ser mayor igual a D16", "[" & D(13) & " - " & D(16) & "]")
        '        End With
        'End Select
        ''2*****************************************************************************************************************************************************************
        'Select Case D(13)
        '    Case Is < D(19)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [02]", "[ERROR]", "Población en Control, celda D13 debe ser mayor igual a D19", "[" & D(13) & " - " & D(19) & "]")
        '        End With
        'End Select
        ''3*****************************************************************************************************************************************************************
        'Select Case D(13)
        '    Case Is < D(22)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [03]", "[ERROR]", "Población en Control, celda D13 debe ser mayor igual a D22", "[" & D(13) & " - " & D(22) & "]")
        '        End With
        'End Select
        ''4*****************************************************************************************************************************************************************
        'Select Case D(14)
        '    Case Is < D(17)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [04]", "[ERROR]", "Población en Control, celda D14 debe ser mayor igual a D17", "[" & D(14) & " - " & D(17) & "]")
        '        End With
        'End Select
        ''5*****************************************************************************************************************************************************************
        'Select Case D(14)
        '    Case Is < D(20)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [05]", "[ERROR]", "Población en Control, celda D14 debe ser mayor igual a D20", "[" & D(14) & " - " & D(20) & "]")
        '        End With
        'End Select
        ''6*****************************************************************************************************************************************************************
        'Select Case D(14)
        '    Case Is < D(23)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [06]", "[ERROR]", "Población en Control, celda D14 debe ser mayor igual a D23", "[" & D(14) & " - " & D(23) & "]")
        '        End With
        'End Select
        ''7*****************************************************************************************************************************************************************
        'Select Case D(15)
        '    Case Is < D(18)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [07]", "[ERROR]", "Población en Control, celda D15 debe ser mayor igual a D18", "[" & D(15) & " - " & D(18) & "]")
        '        End With
        'End Select
        ''8*****************************************************************************************************************************************************************
        'Select Case D(15)
        '    Case Is < D(21)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [08]", "[ERROR]", "Población en Control, celda D15 debe ser mayor igual a D21", "[" & D(15) & " - " & D(21) & "]")
        '        End With
        'End Select
        ''9*****************************************************************************************************************************************************************
        'Select Case D(15)
        '    Case Is < D(24)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [09]", "[ERROR]", "Población en Control, celda D15 debe ser mayor igual a D24", "[" & D(15) & " - " & D(24) & "]")
        '        End With
        'End Select
        ''10*****************************************************************************************************************************************************************
        'Select Case D(12)
        '    Case Is < D(29)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [C]", "VAL [10]", "[ERROR]", "Población en Control, celda D12 debe ser mayor igual a D29", "[" & D(12) & " - " & D(29) & "]")
        '        End With
        'End Select
        ''11*****************************************************************************************************************************************************************
        'Select Case D(29)
        '    Case Is < D(31)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [11]", "[ERROR]", "Población en Control, celda D29 debe ser mayor igual a D31", "[" & D(29) & " - " & D(31) & "]")
        '        End With
        'End Select
        ''12*****************************************************************************************************************************************************************
        'Select Case D(31)
        '    Case Is < D(49)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [C]", "VAL [12]", "[ERROR]", "Población en Control, celda D31 debe ser mayor igual a D49", "[" & D(31) & " - " & D(49) & "]")
        '        End With
        'End Select
        ''13*****************************************************************************************************************************************************************
        'Select Case D(49)
        '    Case Is < D(59)
        '        With Me.DataGridView1.Rows
        '            .Add("P10", " [A]", "VAL [13]", "[ERROR]", "Población en Control, celda D49 debe ser mayor igual a D59", "[" & D(49) & " - " & D(59) & "]")
        '        End With
        'End Select
        ''14*****************************************************************************************************************************************************************
        'Select Case (D(30) + D(31) + D(32) + D(33) + D(34) + D(35) + D(36) + D(37) + D(38) + D(39) + D(40) + D(41) + D(42) + D(43) + D(44) + D(45) + D(46) + D(47))
        '    Case Is > 0
        '        If D(31) > 0 Then
        '            With Me.DataGridView1.Rows
        '                .Add("P10", " [A]", "VAL [14]", "[ERROR]", "Población en Control, la suma de las celdas D32 a D47 es distinto de cero, entonces la celda D31 debe contener dato", "[" & (D(30) + D(31) + D(32) + D(33) + D(34) + D(35) + D(36) + D(37) + D(38) + D(39) + D(40) + D(41) + D(42) + D(43) + D(44) + D(45) + D(46) + D(47)) & " - " & D(31) & "]")
        '            End With
        '        End If
        'End Select
        ''15*****************************************************************************************************************************************************************
        'Select Case (D(50) + D(51) + D(52) + D(53) + D(54) + D(55) + D(56) + D(57))
        '    Case Is > 0
        '        If D(49) > 0 Then
        '            With Me.DataGridView1.Rows
        '                .Add("P10", " [A]", "VAL [15]", "[ERROR]", "Población en Control, la suma de las celdas D50 a D57 es distinto de cero, entonces la celda D49 debe contener dato", "[" & (D(50) + D(51) + D(52) + D(53) + D(54) + D(55) + D(56) + D(57)) & " - " & D(49) & "]")
        '            End With
        '        End If
        'End Select
        ''16*****************************************************************************************************************************************************************
        'Select Case (D(60) + D(61) + D(62) + D(63) + D(64) + D(65) + D(66) + D(67) + D(68) + D(69) + D(70) + D(71))
        '    Case Is > 0
        '        If D(59) > 0 Then
        '            With Me.DataGridView1.Rows
        '                .Add("P10", " [A]", "VAL [16]", "[ERROR]", "Población en Control, la suma de las celdas D60 a D71 es distinto de cero, entonces la celda D59 debe contener dato", "[" & (D(60) + D(61) + D(62) + D(63) + D(64) + D(65) + D(66) + D(67) + D(68) + D(69) + D(70) + D(71)) & " - " & D(59) & "]")
        '            End With
        '        End If
        'End Select


        'xlHoja = Nothing
    End Sub
    Sub P11()
        Dim ii, B(20), C(20), D(20), E(20), F(20), G(20), H(20), I(20), J(20), K(20), L(20), M(20), N(20), O(20), P(20), Q(20), R(20), S(20), T(20), U(20), V(20), W(20), X(20), Y(20), Z(20), AA(20), AB(20), AC(20), AD(20), AE(20), AF(20), AG(20), AH(20), AI(20), AJ(20), AK(20), AL(20), AM(20), AN(20), AO(20), AP(20), AQ(20), AR(20), AS1(20), AT(20), AU(20), AV(20), AW(20) As Integer
        xlHoja = xlLibro.Worksheets("P11")
        ' SECCION A:  POBLACIÓN EN CONTROL DEL PROGRAMA DE VIH/SIDA (Uso exclusivo Centros de Atención VIH-SIDA)
        For ii = 12 To 15
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
            AK(ii) = xlHoja.Range("AK" & ii & "").Value
            AL(ii) = xlHoja.Range("AL" & ii & "").Value
            AM(ii) = xlHoja.Range("AM" & ii & "").Value
            AN(ii) = xlHoja.Range("AN" & ii & "").Value
            AO(ii) = xlHoja.Range("AO" & ii & "").Value
            AP(ii) = xlHoja.Range("AP" & ii & "").Value
            AQ(ii) = xlHoja.Range("AQ" & ii & "").Value
            AR(ii) = xlHoja.Range("AR" & ii & "").Value
            AS1(ii) = xlHoja.Range("AS" & ii & "").Value
            AT(ii) = xlHoja.Range("AT" & ii & "").Value
            AU(ii) = xlHoja.Range("AU" & ii & "").Value
            AV(ii) = xlHoja.Range("AV" & ii & "").Value
            AW(ii) = xlHoja.Range("AW" & ii & "").Value
        Next
        ' SECCION B:  POBLACIÓN EN CONTROL  POR COMERCIO SEXUAL (Uso exclusivo de Unidades Control Comercio Sexual)
        For ii = 20 To 20
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
            Q(ii) = xlHoja.Range("Q" & ii & "").Value
            R(ii) = xlHoja.Range("R" & ii & "").Value
            S(ii) = xlHoja.Range("S" & ii & "").Value
            T(ii) = xlHoja.Range("T" & ii & "").Value
            U(ii) = xlHoja.Range("U" & ii & "").Value
            V(ii) = xlHoja.Range("V" & ii & "").Value
            W(ii) = xlHoja.Range("W" & ii & "").Value
            X(ii) = xlHoja.Range("X" & ii & "").Value
            Y(ii) = xlHoja.Range("Y" & ii & "").Value
            Z(ii) = xlHoja.Range("Z" & ii & "").Value
            AA(ii) = xlHoja.Range("AA" & ii & "").Value
            AB(ii) = xlHoja.Range("AB" & ii & "").Value
            AC(ii) = xlHoja.Range("AC" & ii & "").Value
            AD(ii) = xlHoja.Range("AD" & ii & "").Value
            AE(ii) = xlHoja.Range("AE" & ii & "").Value
            AF(ii) = xlHoja.Range("AF" & ii & "").Value
            AG(ii) = xlHoja.Range("AG" & ii & "").Value
            AH(ii) = xlHoja.Range("AH" & ii & "").Value
            AI(ii) = xlHoja.Range("AI" & ii & "").Value
            AJ(ii) = xlHoja.Range("AJ" & ii & "").Value
        Next
        MsgBox("REMP11 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1*****************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100 'HBO
            Case Else ' el resto de establecimiento
                Select Case (B(12) + B(13) + B(14) + B(15))
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("P11", " [A]", "VAL [01]", "[ERROR]", "Población en Control del Programa ITS, celdas D12 a D15 debe registrar sólo el HBO", "[" & (B(12) + B(13) + B(14) + B(15)) & "]")
                        End With
                End Select
        End Select
        '2*****************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100
            Case Else ' el resto de establecimiento
                Select Case C(20)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("P11", " [B]", "VAL [02]", "[ERROR]", "Población en Control por Comercio Sexual, celda B20 debe registrar sólo el HBO", "[" & C(20) & "]")
                        End With
                End Select
        End Select

        xlHoja = Nothing
    End Sub
    Sub P12()
        Dim ii, B(87), C(87), D(87), E(87), F(87), G(87), H(87), I(87), J(87), K(87), L(87), M(87), N(87), O(87), P(87), Q(87), R(87), S(87), T(87), U(87), V(87), W(87), X(87), Y(87), Z(87), AA(87), AB(87), AC(87), AD(87), AE(87), AF(87), AG(87), AH(87), AI(87), AJ(87), AK(87), AL(87), AM(87), AN(87), AO(87), AP(87), AQ(87), AR(87), AS1(87) As Integer
        xlHoja = xlLibro.Worksheets("P12")
        ' SECCION A: PROGRAMA DE CANCER DE CUELLO UTERINO: POBLACIÓN CON PAP VIGENTE
        For ii = 10 To 23
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
        Next
        ' "CIÓN B.1- PROGRAMA DE CANCER DE CUELLO UTERINO: PAP REALIZADOS E INFORMADOS SEGÚN RESULTADOS (Examen realizados en  la red pública)"										
        For ii = 29 To 42
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
        Next
        ' "CIÓN B.2- PROGRAMA DE CANCER DE CUELLO UTERINO: PAP REALIZADOS E INFORMADOS SEGÚN RESULTADOS (Examen realizados en  extrasistema)"										
        For ii = 48 To 61
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
        Next
        ' SECCION C: PROGRAMA DE CANCER DE MAMA: MUJERES CON MAMOGRAFÍA VIGENTE EN LOS ULTIMOS 3 AÑOS. 
        For ii = 65 To 74
            B(ii) = xlHoja.Range("B" & ii & "").Value
        Next
        ' SECCION D: PROGRAMA DE CANCER DE MAMA: NÚMERO DE MUJERES CON EXAMEN FÍSICO DE MAMA (VIGENTE)
        For ii = 78 To 87
            B(ii) = xlHoja.Range("B" & ii & "").Value
        Next
        MsgBox("REMP12 OK")
        '*************************************************************************************************************************************************************************************
        '********************************************************************** VALIDACIONES *************************************************************************************************
        '*************************************************************************************************************************************************************************************
        '1*****************************************************************************************************************************************************************
        Select Case (B(11) + B(12) + B(13) + B(14) + B(15) + B(16) + B(17) + B(18))
            Case Is = 0
                With Me.DataGridView1.Rows
                    .Add("P12", " [A]", "VAL [01]", "[ERROR]", "Población Femenina con PAP vigente, celdas B11:B18 (grupo-etario 25 a 64 años) deben registran todos los establecimientos de la Red", "[" & (B(11) + B(12) + B(13) + B(14) + B(15) + B(16) + B(17) + B(18)) & "]")
                End With
        End Select
        '2 *****************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100
            Case Else ' el resto de establecimiento
                Select Case B(42)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("P12", " [B]", "VAL [02]", "[ERROR]", "PAP Realizados e Informados según Resultados, celda B42 debe registrar sólo HBO", "[" & B(42) & "]")
                        End With
                End Select
        End Select
        '3*****************************************************************************************************************************************************************
        Select Case B(61)
            Case Is = 0
                With Me.DataGridView1.Rows
                    .Add("P12", " [B.2]", "VAL [03]", "[ERROR]", "Programa de cáncer de cuello uterino, celda B61 registran todos los establecimientos de la red (se excluye el HBO)", "[" & B(61) & "]")
                End With
        End Select
        '4 *****************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100
                Select Case B(74)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("P12", " [C]", "VAL [04]", "[ERROR]", "Programa de cáncer de mama: Mujeres con mamografía vigente en los últimos 3 años, celda B74 deben registrar Todos los Establecimientos de la Red excluyendo al HBO.", "[" & B(74) & "]")
                        End With
                End Select
            Case Else 'resto de estables

        End Select
        '5 *****************************************************************************************************************************************************************
        Select Case CodigoEstablec
            Case 123100
                Select Case B(87)
                    Case Is <> 0
                        With Me.DataGridView1.Rows
                            .Add("P12", " [D]", "VAL [05]", "[ERROR]", "Programa de cáncer de mama: numero de mujeres con examen físico de mama (vigente), celda B87 deben registrar Todos los Establecimientos de la Red excluyendo al HBO; DSSO; HPU y HRN", "[" & B(69) & "]")
                        End With
                End Select
            Case Else 'resto de estables

        End Select
        ' **** 
        'Select Case C(23)
        '    Case Is <> 0
        '        With Me.DataGridView1.Rows
        '            .Add("P12", " [A]", "VAL [06]", "[ERROR]", "HOMBRES con PAP Vigente (Menor o igual a 3 años), No puede existir registros ", "[" & C(23) & "]")
        '        End With
        'End Select

        xlHoja = Nothing
    End Sub
    Sub BM18()
        Dim ii, C(118), D(118), E(118), F(118), G(118), H(118), I(118), J(118), K(118) As Integer
        xlHoja = xlLibro.Worksheets("BM18")

        'SECCION A
        For ii = 12 To 29
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next

        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        '         ASIGNACION VARIABLES //////////
        REMBM18(1) = C(26)
        REMBM18(2) = C(96)
        REMBM18(3) = C(99)


        xlHoja = Nothing
    End Sub
    Sub BM18A()
        Dim ii, C(623), D(623), E(623), F(623), G(623), H(623), I(623), j(623) As Integer
        xlHoja = xlLibro.Worksheets("BM18A")
        ProgressBar1.Maximum = 15
        'SECCIÓN A: EXÁMENES DE DIAGNOSTICO
        For ii = 13 To 103
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 105 To 174
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 176 To 181
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        For ii = 183 To 245
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 247 To 329
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 331 To 375
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 377 To 410
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 412 To 494
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        ' EXAMENES RADIOLOGICOS
        For ii = 412 To 494
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        'SECCIÓN B: PROCEDIMIENTOS APOYO CLÍNICO Y TERAPÉUTICO
        For ii = 499 To 558
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        ' SECCIÓN C: INTERVENCIONES QUIRÚRGICAS MENORES
        For ii = 560 To 571
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        ' SECCIÓN D: MISCELÁNEOS
        For ii = 573 To 589
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 594 To 596
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        'SECCIÓN E: OTROS EXÁMENES Y PROCEDIMIENTOS  DE APOYO CLÍNICO Y TERAPEUTICO (SIN CÓDIGO EN ARANCEL)
        For ii = 600 To 616
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)
        For ii = 621 To 623
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
        Next
        ProgressBar1.Increment(1)
        LBLprogreso.Text = FormatPercent(CSng(ProgressBar1.Value) / (ProgressBar1.Maximum), 0)

        ' 1 **********************************************************************************************************************************************************************************
        Select Case H(611)
            Case Is > F(420)
                With DataGridView1.Rows
                    .Add("BM18A", " [E]", "VAL [01]", "[ERROR]", "Exámenes y/o Procedimientos, radiografia de torax por sospecha de neumonia por compra de servicio, celda H611 debe ser igual o menor que radiografia de torax simple, celda F420", "[" & H(611) & "-" & F(420) & "]")
                End With
            Case Else
        End Select
        ' 2 **********************************************************************************************************************************************************************************
        Select Case H(610)
            Case Is > F(464)
                With DataGridView1.Rows
                    .Add("BM18A", " [E]", "VAL [02]", "[ERROR]", "Exámenes y/o Procedimientos, Radiografía de pelvis, cadera o Coxofemoral de Screening a los 3 meses por compra de servicio, celda H610 debe ser igual o menor que radiografia de Pelvis, Cadera o Coxofemoral celda F464", "[" & H(610) & "-" & F(464) & "]")
                End With

        End Select
        ' 3 **********************************************************************************************************************************************************************************
        Select Case CodigoEstablec ' CESFAM
            Case 123300, 123301, 123302, 123303, 123304, 123305, 123306, 123307, 123308, 123309, 123310, 123311, 123312
                Select Case C(609)
                    Case Is = 0
                        With DataGridView1.Rows
                            .Add("BM18A", " [E]", "VAL [03]", "[ERROR]", "Exámenes y/o procedimientos, Hemuglotest instantaneo, Celda C609 debe tener registros", "[" & C(609) & "]")
                        End With
                End Select
            Case Else ' resto de establecimientos
        End Select
        '**********************************************************************************************************************************************************************************
        Select Case C(529)
            Case Is > 0
                With DataGridView1.Rows
                    .Add("BM18A", " [B]", "VAL [04]", "[REVISAR]", "Nota: Procedimiento Apoyo Clínico y Terapéutico, si existe  Electrocardiograma por 'Producción propia', Celda C529, debe existir registro en Serie A, REM30, Sección C, Total de N° de Informes Electrocardiograma por 'Compra de Servicio', Celda G86 ", "[" & C(529) & "]")
                End With
        End Select

        Select Case F(422)
            Case Is > 0
                With DataGridView1.Rows
                    .Add("BM18A", " [B]", "VAL [05]", "[REVISAR]", "Nota: Total Exámenes Radiológicos, Compra de Servicio de Mamografía Bilateral, Celda F422 no debe registrarse en Serie BM. Se registran en Serie A, Rem30, Sección B, Mamografías", "[" & F(422) & "]")
                End With
        End Select

        Select Case F(488)
            Case Is > 0
                With DataGridView1.Rows
                    .Add("BM18A", " [B]", "VAL [06]", "[REVISAR]", "Nota: Total Exámenes Radiológicos, Compra de Servicio de Ecografía Mamaria Bilateral, Celda F488 no debe registrarse en Serie BM. Se registran en Serie A, Rem30, Sección B, Ecotomografía Mamaria", "[" & F(488) & "]")
                End With
        End Select

        Select Case F(479)
            Case Is > 0
                With DataGridView1.Rows
                    .Add("BM18A", " [B]", "VAL [07]", "[REVISAR]", "Nota: Total Exámenes Radiológicos, Compra de Servicio de Ecografía Abdominal, Celda F479 no debe registrarse en Serie BM. Se registran en Serie A, Rem30, Sección B, Ecotomografía Abdominal", "[" & F(479) & "]")
                End With
        End Select


        xlHoja = Nothing
    End Sub
    Sub D15()
        Dim ii, C(108), D(108), E(108), F(108), G(108), H(108), I(108), J(108), K(108), L(108), M(108), N(108), O(108), P(108) As Integer
        xlHoja = xlLibro.Worksheets("D15")
        ' SECCION A
        For ii = 13 To 29
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
        Next
        ' SECCION B
        For ii = 33 To 36
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        ' SECCION C
        For ii = 42 To 47
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
        Next
        ' SECCION D
        For ii = 53 To 61
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
            P(ii) = xlHoja.Range("P" & ii & "").Value
        Next
        ' SECCION E
        For ii = 67 To 83
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
        Next
        ' SECCION F
        For ii = 87 To 90
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        ' SECCION G
        For ii = 96 To 100
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
            O(ii) = xlHoja.Range("O" & ii & "").Value
        Next
        ' SECCION H
        For ii = 104 To 108
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
        Next

        xlHoja = Nothing
    End Sub
    Sub D16()
        Dim ii, B(39), C(39), D(39), E(39), F(39), G(39), H(39), I(39), J(39), K(39), L(39), M(39), N(39) As Integer
        xlHoja = xlLibro.Worksheets("D16")
        ' SECCION A
        For ii = 10 To 15
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
        Next
        ' SECCION B
        For ii = 19 To 24
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next
        ' SECCION C
        For ii = 29 To 30
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
            H(ii) = xlHoja.Range("H" & ii & "").Value
            I(ii) = xlHoja.Range("I" & ii & "").Value
            J(ii) = xlHoja.Range("J" & ii & "").Value
            K(ii) = xlHoja.Range("K" & ii & "").Value
            L(ii) = xlHoja.Range("L" & ii & "").Value
            M(ii) = xlHoja.Range("M" & ii & "").Value
            N(ii) = xlHoja.Range("N" & ii & "").Value
        Next
        'SECCION D
        For ii = 34 To 39
            B(ii) = xlHoja.Range("B" & ii & "").Value
            C(ii) = xlHoja.Range("C" & ii & "").Value
            D(ii) = xlHoja.Range("D" & ii & "").Value
            E(ii) = xlHoja.Range("E" & ii & "").Value
            F(ii) = xlHoja.Range("F" & ii & "").Value
            G(ii) = xlHoja.Range("G" & ii & "").Value
        Next


        xlHoja = Nothing
    End Sub
    ' Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean

    'Creamos las variables
    'Dim exApp As New Microsoft.Office.Interop.Excel.Application
    'Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
    'Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

    'Try
    '    'Añadimos el Libro al programa, y la hoja al libro
    '    exLibro = exApp.Workbooks.Add
    '    exHoja = exLibro.Worksheets.Add()


    '    ' ¿Cuantas columnas y cuantas filas?
    '    Dim NCol As Integer = ElGrid.ColumnCount
    '    Dim NRow As Integer = ElGrid.RowCount

    '    'Aqui recorremos todas las filas, y por cada fila todas las columnas y vamos escribiendo.
    '    For i As Integer = 1 To NCol
    '        'exHoja.Cells.Item(6, i) = ElGrid.Columns(i - 1).Name.ToString
    '        exHoja.Cells.Item(6, i) = ElGrid.Columns(i - 1).HeaderText.ToString ' desde la fila que empieza a escribir los titulos
    '        exHoja.Cells.Item(1, i).HorizontalAlignment = 3
    '    Next

    '    For Fila As Integer = 0 To NRow - 1
    '        For Col As Integer = 0 To NCol - 1
    '            exHoja.Cells.Item(Fila + 7, Col + 1) = ElGrid.Rows(Fila).Cells(Col).Value ' fila 6 donde empieza a escribir las tablas
    '        Next
    '    Next

    '    'Titulo en negrita, Alineado al centro y que el tamaño de la columna se ajuste al texto


    '    exHoja.Range("E4").Font.Bold = 1
    '    exHoja.Range("A1:A5").Font.Bold = 1
    '    exHoja.Range("A6:G6").Font.Bold = 1
    '    exHoja.Name = "ERRORES"
    '    'exHoja.Range("A1:A4").Font.Bold = 1
    '    exHoja.Rows.Item(5).HorizontalAlignment = 3
    '    exHoja.Rows.Item(1).HorizontalAlignment = 1
    '    exHoja.Rows.Item(3).HorizontalAlignment = 1
    '    exHoja.Columns("E").AutoFit()
    '    'exHoja.Columns.AutoFit()
    '    exHoja.Rows.Font.Size = 10
    '    exHoja.Rows.Font.Name = "Calibri"
    '    exHoja.Columns.Interior.ColorIndex = 2

    '    exHoja.Cells.Range("E4").Value = "DETALLE VALIDACIONES TECNICAS " & Me.LBSerie.Text
    '    exHoja.Cells.Range("A1").Value = "SERVICIO SALUD OSORNO"
    '    'exHoja.Cells.Range("A1:F1").MergeCells = True
    '    exHoja.Cells.Range("A2").Value = "ESTABLECIMIENTO : " & Me.LBEstable.Text & " [ " & Me.ValidaCodigo & " ]" ' ESTABLECIMIENTO
    '    ' exHoja.Cells.Range("A2:F2").MergeCells = True
    '    exHoja.Cells.Range("A3").Value = "COMUNA : " & Me.LBcomuna.Text ' COMUNA
    '    exHoja.Cells.Range("A4").Value = "MES : " & Me.LBmes.Text ' MES
    '    exHoja.Cells.Range("G6").Value = "RESPUESTAS"
    '    exHoja.Cells.Range("F4").Value = "TOTAL ERRORES : " & Me.LBerrores.Text ' TOTAL ERRORES

    '    'Aplicación visible
    '    exApp.Application.Visible = True

    '    exHoja = Nothing
    '    exLibro = Nothing
    '    exApp = Nothing

    'Catch ex As Exception
    '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")
    '    Return False
    'End Try

    'Return True

    '  End Function


End Class