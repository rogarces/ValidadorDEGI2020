<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Validador2019
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Validador2019))
        Me.LBLprogreso = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.LblHojaControl = New System.Windows.Forms.Label()
        Me.TIThojacontrol = New System.Windows.Forms.Label()
        Me.LBerrores = New System.Windows.Forms.Label()
        Me.TITerrores = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.LBdependencia = New System.Windows.Forms.Label()
        Me.LBcomuna = New System.Windows.Forms.Label()
        Me.TITcomuna = New System.Windows.Forms.Label()
        Me.TITdependencia = New System.Windows.Forms.Label()
        Me.LBcodigo = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LBSerie = New System.Windows.Forms.Label()
        Me.TITserie = New System.Windows.Forms.Label()
        Me.LBversion = New System.Windows.Forms.Label()
        Me.LBaño = New System.Windows.Forms.Label()
        Me.TITano = New System.Windows.Forms.Label()
        Me.LBmes = New System.Windows.Forms.Label()
        Me.TITmes = New System.Windows.Forms.Label()
        Me.LBEstable = New System.Windows.Forms.Label()
        Me.TITestable = New System.Windows.Forms.Label()
        Me.TITruta = New System.Windows.Forms.Label()
        Me.TxtRuta = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.BtnExportar = New System.Windows.Forms.Button()
        Me.BtnAbrir = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.lblversionPublicacion = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LBLprogreso
        '
        Me.LBLprogreso.AutoSize = True
        Me.LBLprogreso.BackColor = System.Drawing.Color.Transparent
        Me.LBLprogreso.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBLprogreso.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBLprogreso.Location = New System.Drawing.Point(10, 46)
        Me.LBLprogreso.Name = "LBLprogreso"
        Me.LBLprogreso.Size = New System.Drawing.Size(16, 14)
        Me.LBLprogreso.TabIndex = 29
        Me.LBLprogreso.Text = "..."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.BackColor = System.Drawing.SystemColors.Control
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ProgressBar1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.ProgressBar1.Location = New System.Drawing.Point(3, 18)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(325, 25)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar1.TabIndex = 28
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.LblHojaControl)
        Me.GroupBox3.Controls.Add(Me.TIThojacontrol)
        Me.GroupBox3.Controls.Add(Me.LBerrores)
        Me.GroupBox3.Controls.Add(Me.TITerrores)
        Me.GroupBox3.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.GroupBox3.Location = New System.Drawing.Point(15, 169)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(575, 85)
        Me.GroupBox3.TabIndex = 27
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "INFORME DE ERRORES"
        '
        'LblHojaControl
        '
        Me.LblHojaControl.AutoSize = True
        Me.LblHojaControl.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblHojaControl.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LblHojaControl.Location = New System.Drawing.Point(161, 54)
        Me.LblHojaControl.Name = "LblHojaControl"
        Me.LblHojaControl.Size = New System.Drawing.Size(16, 14)
        Me.LblHojaControl.TabIndex = 4
        Me.LblHojaControl.Text = "..."
        '
        'TIThojacontrol
        '
        Me.TIThojacontrol.AutoSize = True
        Me.TIThojacontrol.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TIThojacontrol.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TIThojacontrol.Location = New System.Drawing.Point(14, 53)
        Me.TIThojacontrol.Name = "TIThojacontrol"
        Me.TIThojacontrol.Size = New System.Drawing.Size(143, 14)
        Me.TIThojacontrol.TabIndex = 3
        Me.TIThojacontrol.Text = "TOTAL Errores Hoja Control:"
        '
        'LBerrores
        '
        Me.LBerrores.AutoSize = True
        Me.LBerrores.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBerrores.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBerrores.Location = New System.Drawing.Point(149, 27)
        Me.LBerrores.Name = "LBerrores"
        Me.LBerrores.Size = New System.Drawing.Size(16, 14)
        Me.LBerrores.TabIndex = 1
        Me.LBerrores.Text = "..."
        '
        'TITerrores
        '
        Me.TITerrores.AutoSize = True
        Me.TITerrores.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITerrores.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITerrores.Location = New System.Drawing.Point(14, 26)
        Me.TITerrores.Name = "TITerrores"
        Me.TITerrores.Size = New System.Drawing.Size(128, 14)
        Me.TITerrores.TabIndex = 0
        Me.TITerrores.Text = "TOTAL Errores Validador:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DataGridView1)
        Me.GroupBox2.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 274)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1114, 238)
        Me.GroupBox2.TabIndex = 26
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "DETALLE DE ERRORES "
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column5, Me.Column3, Me.Column4, Me.Column6})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.EnableHeadersVisualStyles = False
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.Control
        Me.DataGridView1.Location = New System.Drawing.Point(3, 18)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1108, 217)
        Me.DataGridView1.TabIndex = 12
        '
        'Column1
        '
        Me.Column1.HeaderText = "REM"
        Me.Column1.Name = "Column1"
        Me.Column1.ToolTipText = "Indica Hoja REM"
        Me.Column1.Width = 50
        '
        'Column2
        '
        Me.Column2.HeaderText = "SECCIÓN"
        Me.Column2.Name = "Column2"
        Me.Column2.ToolTipText = "Corresponde a la Sección de la Hoja Rem"
        Me.Column2.Width = 60
        '
        'Column5
        '
        Me.Column5.HeaderText = "Nº VAL"
        Me.Column5.Name = "Column5"
        Me.Column5.ToolTipText = "Indica Numero Validación segun el manual de validaciones"
        Me.Column5.Width = 55
        '
        'Column3
        '
        Me.Column3.HeaderText = "TIPO"
        Me.Column3.Name = "Column3"
        Me.Column3.ToolTipText = "Indica el tipo de error ya sea : Error o Revisar"
        Me.Column3.Width = 60
        '
        'Column4
        '
        Me.Column4.HeaderText = "DETALLE"
        Me.Column4.Name = "Column4"
        Me.Column4.ToolTipText = "Indica detalle de validaciones Tecnicas segun el manual disponible"
        Me.Column4.Width = 780
        '
        'Column6
        '
        Me.Column6.HeaderText = "V. VARIABLE"
        Me.Column6.Name = "Column6"
        Me.Column6.ToolTipText = "Indiva los Valores de las variables quie se estan comparando"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LBdependencia)
        Me.GroupBox1.Controls.Add(Me.LBcomuna)
        Me.GroupBox1.Controls.Add(Me.TITcomuna)
        Me.GroupBox1.Controls.Add(Me.TITdependencia)
        Me.GroupBox1.Controls.Add(Me.LBcodigo)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.LBSerie)
        Me.GroupBox1.Controls.Add(Me.TITserie)
        Me.GroupBox1.Controls.Add(Me.LBversion)
        Me.GroupBox1.Controls.Add(Me.LBaño)
        Me.GroupBox1.Controls.Add(Me.TITano)
        Me.GroupBox1.Controls.Add(Me.LBmes)
        Me.GroupBox1.Controls.Add(Me.TITmes)
        Me.GroupBox1.Controls.Add(Me.LBEstable)
        Me.GroupBox1.Controls.Add(Me.TITestable)
        Me.GroupBox1.Controls.Add(Me.TITruta)
        Me.GroupBox1.Controls.Add(Me.TxtRuta)
        Me.GroupBox1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.GroupBox1.Location = New System.Drawing.Point(12, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(687, 156)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "DATOS ESTABLECIMIENTOS"
        '
        'LBdependencia
        '
        Me.LBdependencia.AutoSize = True
        Me.LBdependencia.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBdependencia.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBdependencia.Location = New System.Drawing.Point(372, 103)
        Me.LBdependencia.Name = "LBdependencia"
        Me.LBdependencia.Size = New System.Drawing.Size(16, 14)
        Me.LBdependencia.TabIndex = 27
        Me.LBdependencia.Text = "..."
        '
        'LBcomuna
        '
        Me.LBcomuna.AutoSize = True
        Me.LBcomuna.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBcomuna.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBcomuna.Location = New System.Drawing.Point(77, 103)
        Me.LBcomuna.Name = "LBcomuna"
        Me.LBcomuna.Size = New System.Drawing.Size(16, 14)
        Me.LBcomuna.TabIndex = 25
        Me.LBcomuna.Text = "..."
        '
        'TITcomuna
        '
        Me.TITcomuna.AutoSize = True
        Me.TITcomuna.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITcomuna.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITcomuna.Location = New System.Drawing.Point(14, 103)
        Me.TITcomuna.Name = "TITcomuna"
        Me.TITcomuna.Size = New System.Drawing.Size(60, 14)
        Me.TITcomuna.TabIndex = 24
        Me.TITcomuna.Text = "COMUNA :"
        '
        'TITdependencia
        '
        Me.TITdependencia.AutoSize = True
        Me.TITdependencia.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITdependencia.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITdependencia.Location = New System.Drawing.Point(282, 103)
        Me.TITdependencia.Name = "TITdependencia"
        Me.TITdependencia.Size = New System.Drawing.Size(85, 14)
        Me.TITdependencia.TabIndex = 26
        Me.TITdependencia.Text = "DEPENDENCIA :"
        '
        'LBcodigo
        '
        Me.LBcodigo.AutoSize = True
        Me.LBcodigo.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBcodigo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBcodigo.Location = New System.Drawing.Point(168, 76)
        Me.LBcodigo.Name = "LBcodigo"
        Me.LBcodigo.Size = New System.Drawing.Size(16, 14)
        Me.LBcodigo.TabIndex = 23
        Me.LBcodigo.Text = "..."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(14, 76)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(149, 14)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "CODIGO ESTABLECIMIENTO :"
        '
        'LBSerie
        '
        Me.LBSerie.AutoSize = True
        Me.LBSerie.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBSerie.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBSerie.Location = New System.Drawing.Point(457, 76)
        Me.LBSerie.Name = "LBSerie"
        Me.LBSerie.Size = New System.Drawing.Size(19, 14)
        Me.LBSerie.TabIndex = 21
        Me.LBSerie.Text = "...."
        '
        'TITserie
        '
        Me.TITserie.AutoSize = True
        Me.TITserie.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITserie.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITserie.Location = New System.Drawing.Point(415, 76)
        Me.TITserie.Name = "TITserie"
        Me.TITserie.Size = New System.Drawing.Size(38, 14)
        Me.TITserie.TabIndex = 20
        Me.TITserie.Text = "SERIE:"
        '
        'LBversion
        '
        Me.LBversion.AutoSize = True
        Me.LBversion.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBversion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBversion.Location = New System.Drawing.Point(391, 130)
        Me.LBversion.Name = "LBversion"
        Me.LBversion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LBversion.Size = New System.Drawing.Size(16, 14)
        Me.LBversion.TabIndex = 19
        Me.LBversion.Text = "..."
        '
        'LBaño
        '
        Me.LBaño.AutoSize = True
        Me.LBaño.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBaño.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBaño.Location = New System.Drawing.Point(207, 130)
        Me.LBaño.Name = "LBaño"
        Me.LBaño.Size = New System.Drawing.Size(19, 14)
        Me.LBaño.TabIndex = 17
        Me.LBaño.Text = "...."
        '
        'TITano
        '
        Me.TITano.AutoSize = True
        Me.TITano.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITano.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITano.Location = New System.Drawing.Point(168, 130)
        Me.TITano.Name = "TITano"
        Me.TITano.Size = New System.Drawing.Size(36, 14)
        Me.TITano.TabIndex = 16
        Me.TITano.Text = "AÑO :"
        '
        'LBmes
        '
        Me.LBmes.AutoSize = True
        Me.LBmes.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBmes.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBmes.Location = New System.Drawing.Point(51, 130)
        Me.LBmes.Name = "LBmes"
        Me.LBmes.Size = New System.Drawing.Size(16, 14)
        Me.LBmes.TabIndex = 15
        Me.LBmes.Text = "..."
        '
        'TITmes
        '
        Me.TITmes.AutoSize = True
        Me.TITmes.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITmes.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITmes.Location = New System.Drawing.Point(14, 130)
        Me.TITmes.Name = "TITmes"
        Me.TITmes.Size = New System.Drawing.Size(35, 14)
        Me.TITmes.TabIndex = 14
        Me.TITmes.Text = "MES :"
        '
        'LBEstable
        '
        Me.LBEstable.AutoSize = True
        Me.LBEstable.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBEstable.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.LBEstable.Location = New System.Drawing.Point(125, 49)
        Me.LBEstable.Name = "LBEstable"
        Me.LBEstable.Size = New System.Drawing.Size(16, 14)
        Me.LBEstable.TabIndex = 13
        Me.LBEstable.Text = "..."
        '
        'TITestable
        '
        Me.TITestable.AutoSize = True
        Me.TITestable.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITestable.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITestable.Location = New System.Drawing.Point(14, 49)
        Me.TITestable.Name = "TITestable"
        Me.TITestable.Size = New System.Drawing.Size(108, 14)
        Me.TITestable.TabIndex = 12
        Me.TITestable.Text = "ESTABLECIMIENTO : "
        '
        'TITruta
        '
        Me.TITruta.AutoSize = True
        Me.TITruta.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TITruta.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.TITruta.Location = New System.Drawing.Point(14, 22)
        Me.TITruta.Name = "TITruta"
        Me.TITruta.Size = New System.Drawing.Size(37, 14)
        Me.TITruta.TabIndex = 11
        Me.TITruta.Text = "RUTA:"
        '
        'TxtRuta
        '
        Me.TxtRuta.Location = New System.Drawing.Point(64, 19)
        Me.TxtRuta.Name = "TxtRuta"
        Me.TxtRuta.Size = New System.Drawing.Size(599, 22)
        Me.TxtRuta.TabIndex = 10
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'BtnSalir
        '
        Me.BtnSalir.BackColor = System.Drawing.Color.Firebrick
        Me.BtnSalir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnSalir.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSalir.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnSalir.Location = New System.Drawing.Point(964, 12)
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.Size = New System.Drawing.Size(113, 141)
        Me.BtnSalir.TabIndex = 23
        Me.BtnSalir.Text = "SALIR"
        Me.BtnSalir.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.BtnSalir.UseVisualStyleBackColor = False
        '
        'BtnExportar
        '
        Me.BtnExportar.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.BtnExportar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnExportar.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnExportar.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnExportar.Location = New System.Drawing.Point(766, 86)
        Me.BtnExportar.Name = "BtnExportar"
        Me.BtnExportar.Size = New System.Drawing.Size(192, 67)
        Me.BtnExportar.TabIndex = 24
        Me.BtnExportar.Text = "EXPORTAR"
        Me.BtnExportar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BtnExportar.UseVisualStyleBackColor = False
        '
        'BtnAbrir
        '
        Me.BtnAbrir.BackColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.BtnAbrir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.BtnAbrir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAbrir.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAbrir.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnAbrir.Location = New System.Drawing.Point(766, 12)
        Me.BtnAbrir.Name = "BtnAbrir"
        Me.BtnAbrir.Size = New System.Drawing.Size(192, 68)
        Me.BtnAbrir.TabIndex = 1
        Me.BtnAbrir.Text = "ABRIR"
        Me.BtnAbrir.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BtnAbrir.UseVisualStyleBackColor = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.LBLprogreso)
        Me.GroupBox4.Controls.Add(Me.ProgressBar1)
        Me.GroupBox4.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(63, Byte), Integer))
        Me.GroupBox4.Location = New System.Drawing.Point(766, 159)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(331, 68)
        Me.GroupBox4.TabIndex = 30
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "% de Carga"
        '
        'lblversionPublicacion
        '
        Me.lblversionPublicacion.AutoSize = True
        Me.lblversionPublicacion.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblversionPublicacion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.lblversionPublicacion.Location = New System.Drawing.Point(768, 240)
        Me.lblversionPublicacion.Name = "lblversionPublicacion"
        Me.lblversionPublicacion.Size = New System.Drawing.Size(13, 14)
        Me.lblversionPublicacion.TabIndex = 31
        Me.lblversionPublicacion.Text = ".."
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.ValidadorRemDEGI.My.Resources.Resources.degi2019
        Me.PictureBox1.Location = New System.Drawing.Point(595, 170)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(167, 99)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Validador2019
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(1138, 525)
        Me.Controls.Add(Me.lblversionPublicacion)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.BtnSalir)
        Me.Controls.Add(Me.BtnExportar)
        Me.Controls.Add(Me.BtnAbrir)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.SystemColors.Control
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Validador2019"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Validador 2019 Serie A 1.3 (Actualizado)"
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LBLprogreso As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents LblHojaControl As System.Windows.Forms.Label
    Friend WithEvents TIThojacontrol As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents LBerrores As System.Windows.Forms.Label
    Friend WithEvents TITerrores As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LBdependencia As System.Windows.Forms.Label
    Friend WithEvents TITdependencia As System.Windows.Forms.Label
    Friend WithEvents LBcomuna As System.Windows.Forms.Label
    Friend WithEvents TITcomuna As System.Windows.Forms.Label
    Friend WithEvents LBcodigo As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LBSerie As System.Windows.Forms.Label
    Friend WithEvents TITserie As System.Windows.Forms.Label
    Friend WithEvents LBversion As System.Windows.Forms.Label
    Friend WithEvents LBaño As System.Windows.Forms.Label
    Friend WithEvents TITano As System.Windows.Forms.Label
    Friend WithEvents LBmes As System.Windows.Forms.Label
    Friend WithEvents TITmes As System.Windows.Forms.Label
    Friend WithEvents LBEstable As System.Windows.Forms.Label
    Friend WithEvents TITestable As System.Windows.Forms.Label
    Friend WithEvents TITruta As System.Windows.Forms.Label
    Friend WithEvents TxtRuta As System.Windows.Forms.TextBox
    Friend WithEvents BtnSalir As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BtnExportar As System.Windows.Forms.Button
    Friend WithEvents BtnAbrir As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents lblversionPublicacion As System.Windows.Forms.Label
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
End Class
