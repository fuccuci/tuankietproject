Imports System.Data.OleDb
Imports ConvertNumberToChar
Public Class frmThanhtoanDL
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Dim start As Boolean = False
    Dim splitn As New SplitNumbers
    Dim numbers As New ConvertNumbersToString
    Dim rpt As CrystalReport_Receipts
    Dim rptPC As CrystalReport_Expense
    Dim SaveFlag As Boolean = False
    Dim Arrvar(15) As String
    Dim Arrval(15) As String
    Dim Maxnumber As Integer
    Dim MaxDL As Integer
    Dim dsdl As DataSet
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        FillDataSet()

        Try
            strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
            Try
                CboEmploy_code.DataSource = Nothing
                Try
                    mydataset.Tables("Tbl_Employee").Clear()
                Catch ex As Exception
                End Try
                CboEmploy_code.Items.Clear()
                FillCombo(CboEmploy_code, strSQL, "Tbl_Employee", "Employ_Code", "Employ_Name")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Catch ex As Exception
        End Try
        start = True
        If (CboEmploy_code.Items.Count > 0) Then
            CboEmploy_code.SelectedIndex = 0
            txtEmployeeName.Text = CboEmploy_code.SelectedValue
        End If

        strSQL = " SELECT MAX(Ordinal_No)+1  FROM Tbl_Receipts WHERE  MaLoaiThu = 'TM'" & " AND ( MONTH(Receipt_Date) = " & DateTimePickerngay.Value.Month & ") AND  (YEAR(Receipt_Date) = " & DateTimePickerngay.Value.Year() & ")"
        Maxnumber = GetMaxNumber(strSQL)
        txtsoPTDS.Text = Maxnumber


        strSQL = " SELECT MAX(Ordinal_No)+1  FROM Tbl_Expenses WHERE  ( MONTH(Expense_Date) = " & DateTimePickerngay.Value.Month & ") AND  (YEAR(Expense_Date) = " & DateTimePickerngay.Value.Year() & ")"
        txtsophieuhoahong.Text = GetMaxNumber(strSQL)

        txtsophieuthue.Text = Maxnumber + 1

        rpt = New CrystalReport_Receipts
        rptPC = New CrystalReport_Expense
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDownSobanin As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents DateTimePickerchuky As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdIn As System.Windows.Forms.Button
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents cmdLuu As System.Windows.Forms.Button
    Friend WithEvents cmddeletes As System.Windows.Forms.Button
    Friend WithEvents DateTimePickerngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents DataGridListReceipts As System.Windows.Forms.DataGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txttendaily As System.Windows.Forms.TextBox
    Friend WithEvents txtdiachidaily As System.Windows.Forms.TextBox
    Friend WithEvents txtchitietctgochoahong As System.Windows.Forms.TextBox
    Friend WithEvents txttienhoahong As System.Windows.Forms.TextBox
    Friend WithEvents txtchitietctgocthue As System.Windows.Forms.TextBox
    Friend WithEvents txttienthue As System.Windows.Forms.TextBox
    Friend WithEvents txtsophieuthue As System.Windows.Forms.TextBox
    Friend WithEvents txtsophieuhoahong As System.Windows.Forms.TextBox
    Friend WithEvents txtsoPTDS As System.Windows.Forms.TextBox
    Friend WithEvents txtsophaitra As System.Windows.Forms.TextBox
    Friend WithEvents txtTiendoanhso As System.Windows.Forms.TextBox
    Friend WithEvents txtchitietctgds As System.Windows.Forms.TextBox
    Friend WithEvents cboDsDaiLy As System.Windows.Forms.ComboBox
    Friend WithEvents cmbHTthu As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxPTTHUE As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxPCHH As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxPTDS As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.CheckBoxPTTHUE = New System.Windows.Forms.CheckBox
        Me.CheckBoxPCHH = New System.Windows.Forms.CheckBox
        Me.CheckBoxPTDS = New System.Windows.Forms.CheckBox
        Me.txtchitietctgocthue = New System.Windows.Forms.TextBox
        Me.txtsophieuthue = New System.Windows.Forms.TextBox
        Me.txtsophaitra = New System.Windows.Forms.TextBox
        Me.cboDsDaiLy = New System.Windows.Forms.ComboBox
        Me.txtchitietctgds = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.DateTimePickerngay = New System.Windows.Forms.DateTimePicker
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txttendaily = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtdiachidaily = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.DateTimePickerchuky = New System.Windows.Forms.DateTimePicker
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtchitietctgochoahong = New System.Windows.Forms.TextBox
        Me.TextBox9 = New System.Windows.Forms.TextBox
        Me.txttienhoahong = New System.Windows.Forms.TextBox
        Me.TextBox14 = New System.Windows.Forms.TextBox
        Me.txttienthue = New System.Windows.Forms.TextBox
        Me.txtsophieuhoahong = New System.Windows.Forms.TextBox
        Me.txtsoPTDS = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtTiendoanhso = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdIn = New System.Windows.Forms.Button
        Me.cmddong = New System.Windows.Forms.Button
        Me.cmdLuu = New System.Windows.Forms.Button
        Me.cmddeletes = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.NumericUpDownSobanin = New System.Windows.Forms.NumericUpDown
        Me.Label15 = New System.Windows.Forms.Label
        Me.cmbHTthu = New System.Windows.Forms.ComboBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.NumericUpDownSobanin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(482, 10)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(189, 27)
        Me.cbostations.TabIndex = 54
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.txtchitietctgocthue)
        Me.GroupBox1.Controls.Add(Me.txtsophieuthue)
        Me.GroupBox1.Controls.Add(Me.txtsophaitra)
        Me.GroupBox1.Controls.Add(Me.cboDsDaiLy)
        Me.GroupBox1.Controls.Add(Me.txtchitietctgds)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerngay)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txttendaily)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtdiachidaily)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerchuky)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.txtchitietctgochoahong)
        Me.GroupBox1.Controls.Add(Me.TextBox9)
        Me.GroupBox1.Controls.Add(Me.txttienhoahong)
        Me.GroupBox1.Controls.Add(Me.TextBox14)
        Me.GroupBox1.Controls.Add(Me.txttienthue)
        Me.GroupBox1.Controls.Add(Me.txtsophieuhoahong)
        Me.GroupBox1.Controls.Add(Me.txtsoPTDS)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.txtTiendoanhso)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(4, 37)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(668, 323)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CheckBoxPTTHUE)
        Me.GroupBox3.Controls.Add(Me.CheckBoxPCHH)
        Me.GroupBox3.Controls.Add(Me.CheckBoxPTDS)
        Me.GroupBox3.Location = New System.Drawing.Point(212, 224)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(448, 40)
        Me.GroupBox3.TabIndex = 52
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "In phiếu"
        '
        'CheckBoxPTTHUE
        '
        Me.CheckBoxPTTHUE.Checked = True
        Me.CheckBoxPTTHUE.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxPTTHUE.Location = New System.Drawing.Point(307, 12)
        Me.CheckBoxPTTHUE.Name = "CheckBoxPTTHUE"
        Me.CheckBoxPTTHUE.Size = New System.Drawing.Size(136, 24)
        Me.CheckBoxPTTHUE.TabIndex = 58
        Me.CheckBoxPTTHUE.Text = "PT Thuế TNDN"
        '
        'CheckBoxPCHH
        '
        Me.CheckBoxPCHH.Checked = True
        Me.CheckBoxPCHH.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxPCHH.Location = New System.Drawing.Point(208, 12)
        Me.CheckBoxPCHH.Name = "CheckBoxPCHH"
        Me.CheckBoxPCHH.Size = New System.Drawing.Size(80, 24)
        Me.CheckBoxPCHH.TabIndex = 57
        Me.CheckBoxPCHH.Text = "PC HH"
        '
        'CheckBoxPTDS
        '
        Me.CheckBoxPTDS.Checked = True
        Me.CheckBoxPTDS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxPTDS.Location = New System.Drawing.Point(96, 12)
        Me.CheckBoxPTDS.Name = "CheckBoxPTDS"
        Me.CheckBoxPTDS.Size = New System.Drawing.Size(80, 24)
        Me.CheckBoxPTDS.TabIndex = 56
        Me.CheckBoxPTDS.Text = "PT DS"
        '
        'txtchitietctgocthue
        '
        Me.txtchitietctgocthue.BackColor = System.Drawing.Color.White
        Me.txtchitietctgocthue.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtchitietctgocthue.ForeColor = System.Drawing.Color.Blue
        Me.txtchitietctgocthue.Location = New System.Drawing.Point(287, 196)
        Me.txtchitietctgocthue.Name = "txtchitietctgocthue"
        Me.txtchitietctgocthue.Size = New System.Drawing.Size(377, 26)
        Me.txtchitietctgocthue.TabIndex = 17
        Me.txtchitietctgocthue.Text = "1. Phiếu thanh toán ĐL "
        '
        'txtsophieuthue
        '
        Me.txtsophieuthue.BackColor = System.Drawing.Color.White
        Me.txtsophieuthue.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsophieuthue.ForeColor = System.Drawing.Color.Blue
        Me.txtsophieuthue.Location = New System.Drawing.Point(208, 196)
        Me.txtsophieuthue.Name = "txtsophieuthue"
        Me.txtsophieuthue.ReadOnly = True
        Me.txtsophieuthue.Size = New System.Drawing.Size(77, 26)
        Me.txtsophieuthue.TabIndex = 15
        Me.txtsophieuthue.Text = ""
        Me.txtsophieuthue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtsophaitra
        '
        Me.txtsophaitra.AutoSize = False
        Me.txtsophaitra.BackColor = System.Drawing.Color.White
        Me.txtsophaitra.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsophaitra.ForeColor = System.Drawing.Color.Brown
        Me.txtsophaitra.Location = New System.Drawing.Point(95, 231)
        Me.txtsophaitra.Name = "txtsophaitra"
        Me.txtsophaitra.ReadOnly = True
        Me.txtsophaitra.Size = New System.Drawing.Size(112, 32)
        Me.txtsophaitra.TabIndex = 18
        Me.txtsophaitra.Text = ""
        Me.txtsophaitra.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboDsDaiLy
        '
        Me.cboDsDaiLy.Location = New System.Drawing.Point(61, 45)
        Me.cboDsDaiLy.Name = "cboDsDaiLy"
        Me.cboDsDaiLy.Size = New System.Drawing.Size(201, 27)
        Me.cboDsDaiLy.TabIndex = 3
        '
        'txtchitietctgds
        '
        Me.txtchitietctgds.BackColor = System.Drawing.Color.White
        Me.txtchitietctgds.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtchitietctgds.ForeColor = System.Drawing.Color.Blue
        Me.txtchitietctgds.Location = New System.Drawing.Point(287, 140)
        Me.txtchitietctgds.Name = "txtchitietctgds"
        Me.txtchitietctgds.Size = New System.Drawing.Size(377, 26)
        Me.txtchitietctgds.TabIndex = 8
        Me.txtchitietctgds.Text = "1.HĐ ; 1 phiếu thanh toán ĐL ; 1 BKê "
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(6, 104)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(658, 2)
        Me.Label17.TabIndex = 51
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Location = New System.Drawing.Point(264, 14)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(178, 27)
        Me.CboEmploy_code.TabIndex = 2
        '
        'DateTimePickerngay
        '
        Me.DateTimePickerngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngay.Location = New System.Drawing.Point(61, 15)
        Me.DateTimePickerngay.Name = "DateTimePickerngay"
        Me.DateTimePickerngay.Size = New System.Drawing.Size(105, 26)
        Me.DateTimePickerngay.TabIndex = 1
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(442, 14)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(224, 26)
        Me.txtEmployeeName.TabIndex = 50
        Me.txtEmployeeName.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 22)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Ngày"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(168, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(99, 22)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Mã/Tên CTV "
        '
        'txttendaily
        '
        Me.txttendaily.BackColor = System.Drawing.Color.White
        Me.txttendaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendaily.ForeColor = System.Drawing.Color.Blue
        Me.txttendaily.Location = New System.Drawing.Point(264, 45)
        Me.txttendaily.Name = "txttendaily"
        Me.txttendaily.ReadOnly = True
        Me.txttendaily.Size = New System.Drawing.Size(400, 26)
        Me.txttendaily.TabIndex = 4
        Me.txttendaily.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 22)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "Đại lý"
        '
        'txtdiachidaily
        '
        Me.txtdiachidaily.BackColor = System.Drawing.Color.White
        Me.txtdiachidaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdiachidaily.ForeColor = System.Drawing.Color.Blue
        Me.txtdiachidaily.Location = New System.Drawing.Point(264, 74)
        Me.txtdiachidaily.Name = "txtdiachidaily"
        Me.txtdiachidaily.ReadOnly = True
        Me.txtdiachidaily.Size = New System.Drawing.Size(400, 26)
        Me.txtdiachidaily.TabIndex = 5
        Me.txtdiachidaily.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(168, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 22)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Đ/c đại lý"
        '
        'DateTimePickerchuky
        '
        Me.DateTimePickerchuky.CustomFormat = "MM/yyyy"
        Me.DateTimePickerchuky.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerchuky.Location = New System.Drawing.Point(62, 74)
        Me.DateTimePickerchuky.Name = "DateTimePickerchuky"
        Me.DateTimePickerchuky.Size = New System.Drawing.Size(105, 26)
        Me.DateTimePickerchuky.TabIndex = 4
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 76)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(54, 22)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Chu kỳ"
        '
        'txtchitietctgochoahong
        '
        Me.txtchitietctgochoahong.BackColor = System.Drawing.Color.White
        Me.txtchitietctgochoahong.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtchitietctgochoahong.ForeColor = System.Drawing.Color.Blue
        Me.txtchitietctgochoahong.Location = New System.Drawing.Point(287, 168)
        Me.txtchitietctgochoahong.Name = "txtchitietctgochoahong"
        Me.txtchitietctgochoahong.Size = New System.Drawing.Size(377, 26)
        Me.txtchitietctgochoahong.TabIndex = 13
        Me.txtchitietctgochoahong.Text = "1. Phiếu thu số ; 1 phiếu thanh toán ĐL "
        '
        'TextBox9
        '
        Me.TextBox9.BackColor = System.Drawing.Color.White
        Me.TextBox9.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox9.ForeColor = System.Drawing.Color.Blue
        Me.TextBox9.Location = New System.Drawing.Point(256, 546)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(104, 26)
        Me.TextBox9.TabIndex = 50
        Me.TextBox9.Text = ""
        '
        'txttienhoahong
        '
        Me.txttienhoahong.BackColor = System.Drawing.Color.White
        Me.txttienhoahong.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttienhoahong.ForeColor = System.Drawing.Color.Blue
        Me.txttienhoahong.Location = New System.Drawing.Point(95, 168)
        Me.txttienhoahong.Name = "txttienhoahong"
        Me.txttienhoahong.Size = New System.Drawing.Size(112, 26)
        Me.txttienhoahong.TabIndex = 10
        Me.txttienhoahong.Text = ""
        Me.txttienhoahong.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox14
        '
        Me.TextBox14.BackColor = System.Drawing.Color.White
        Me.TextBox14.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox14.ForeColor = System.Drawing.Color.Blue
        Me.TextBox14.Location = New System.Drawing.Point(256, 572)
        Me.TextBox14.Name = "TextBox14"
        Me.TextBox14.Size = New System.Drawing.Size(104, 26)
        Me.TextBox14.TabIndex = 50
        Me.TextBox14.Text = ""
        '
        'txttienthue
        '
        Me.txttienthue.BackColor = System.Drawing.Color.White
        Me.txttienthue.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttienthue.ForeColor = System.Drawing.Color.Blue
        Me.txttienthue.Location = New System.Drawing.Point(95, 196)
        Me.txttienthue.Name = "txttienthue"
        Me.txttienthue.Size = New System.Drawing.Size(112, 26)
        Me.txttienthue.TabIndex = 14
        Me.txttienthue.Text = ""
        Me.txttienthue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtsophieuhoahong
        '
        Me.txtsophieuhoahong.BackColor = System.Drawing.Color.White
        Me.txtsophieuhoahong.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsophieuhoahong.ForeColor = System.Drawing.Color.Blue
        Me.txtsophieuhoahong.Location = New System.Drawing.Point(208, 168)
        Me.txtsophieuhoahong.Name = "txtsophieuhoahong"
        Me.txtsophieuhoahong.Size = New System.Drawing.Size(77, 26)
        Me.txtsophieuhoahong.TabIndex = 11
        Me.txtsophieuhoahong.Text = ""
        Me.txtsophieuhoahong.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtsoPTDS
        '
        Me.txtsoPTDS.BackColor = System.Drawing.Color.White
        Me.txtsoPTDS.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsoPTDS.ForeColor = System.Drawing.Color.Blue
        Me.txtsoPTDS.Location = New System.Drawing.Point(208, 140)
        Me.txtsoPTDS.Name = "txtsoPTDS"
        Me.txtsoPTDS.Size = New System.Drawing.Size(77, 26)
        Me.txtsoPTDS.TabIndex = 7
        Me.txtsoPTDS.Text = ""
        Me.txtsoPTDS.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(4, 236)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(84, 22)
        Me.Label16.TabIndex = 25
        Me.Label16.Text = "Số phải trả"
        '
        'txtTiendoanhso
        '
        Me.txtTiendoanhso.BackColor = System.Drawing.Color.White
        Me.txtTiendoanhso.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTiendoanhso.ForeColor = System.Drawing.Color.Blue
        Me.txtTiendoanhso.Location = New System.Drawing.Point(95, 140)
        Me.txtTiendoanhso.Name = "txtTiendoanhso"
        Me.txtTiendoanhso.Size = New System.Drawing.Size(112, 26)
        Me.txtTiendoanhso.TabIndex = 6
        Me.txtTiendoanhso.Text = ""
        Me.txtTiendoanhso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(4, 199)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(93, 22)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Thuế TNDN"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(4, 172)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(74, 22)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Hoa hồng"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(4, 142)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(75, 22)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Doanh số "
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(127, 118)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(54, 22)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Số tiền"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(212, 118)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(67, 22)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "Số phiếu"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(400, 117)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(154, 22)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Chi tiết chứng từ gốc"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdIn)
        Me.GroupBox2.Controls.Add(Me.cmddong)
        Me.GroupBox2.Controls.Add(Me.cmdLuu)
        Me.GroupBox2.Controls.Add(Me.cmddeletes)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.NumericUpDownSobanin)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 265)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(656, 48)
        Me.GroupBox2.TabIndex = 19
        Me.GroupBox2.TabStop = False
        '
        'cmdIn
        '
        Me.cmdIn.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdIn.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdIn.Location = New System.Drawing.Point(136, 14)
        Me.cmdIn.Name = "cmdIn"
        Me.cmdIn.Size = New System.Drawing.Size(80, 27)
        Me.cmdIn.TabIndex = 21
        Me.cmdIn.Text = "In"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(568, 14)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(80, 27)
        Me.cmddong.TabIndex = 31
        Me.cmddong.Text = "Đóng"
        '
        'cmdLuu
        '
        Me.cmdLuu.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuu.Location = New System.Drawing.Point(224, 14)
        Me.cmdLuu.Name = "cmdLuu"
        Me.cmdLuu.Size = New System.Drawing.Size(80, 27)
        Me.cmdLuu.TabIndex = 20
        Me.cmdLuu.Text = "Lưu"
        '
        'cmddeletes
        '
        Me.cmddeletes.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddeletes.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddeletes.Location = New System.Drawing.Point(40, 14)
        Me.cmddeletes.Name = "cmddeletes"
        Me.cmddeletes.Size = New System.Drawing.Size(80, 27)
        Me.cmddeletes.TabIndex = 23
        Me.cmddeletes.Text = "Tạo mới"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(432, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 22)
        Me.Label9.TabIndex = 52
        Me.Label9.Text = "Số bản in "
        '
        'NumericUpDownSobanin
        '
        Me.NumericUpDownSobanin.Location = New System.Drawing.Point(512, 16)
        Me.NumericUpDownSobanin.Name = "NumericUpDownSobanin"
        Me.NumericUpDownSobanin.Size = New System.Drawing.Size(48, 26)
        Me.NumericUpDownSobanin.TabIndex = 51
        Me.NumericUpDownSobanin.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label15
        '
        Me.Label15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label15.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label15.Font = New System.Drawing.Font("Arial", 22.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label15.Location = New System.Drawing.Point(4, 6)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(668, 34)
        Me.Label15.TabIndex = 56
        Me.Label15.Text = "THANH TOÁN ĐL PSTN  "
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbHTthu
        '
        Me.cmbHTthu.BackColor = System.Drawing.Color.White
        Me.cmbHTthu.Enabled = False
        Me.cmbHTthu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHTthu.Location = New System.Drawing.Point(328, 10)
        Me.cmbHTthu.Name = "cmbHTthu"
        Me.cmbHTthu.Size = New System.Drawing.Size(152, 27)
        Me.cmbHTthu.TabIndex = 57
        '
        'frmThanhtoanDL
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(678, 363)
        Me.Controls.Add(Me.cmbHTthu)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmThanhtoanDL"
        Me.Text = "Thanh toán đại lý"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.NumericUpDownSobanin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox14 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox

    Private Sub cmddong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddong.Click
        Me.Close()
    End Sub

    Public Sub FillDataSet()
        mydataset = New DataSet
        strSQL = "SELECT StationID,Station_Name,Station_Address FROM Tbl_Stations "
        FillCombo(cbostations, strSQL, "Tbl_Stations", "Station_Name", "StationID")

        strSQL = "SELECT MaLoaithu,TenLoaiThu FROM Tbl_LoaiThu"
        FillCombo(cmbHTthu, strSQL, "Tbl_LoaiThu", "TenLoaiThu", "MaLoaithu")

    End Sub

    Private Sub FillCombo(ByRef cbo As ComboBox, ByVal strQuery As String, ByVal strTablename As String, ByVal strdislaymember As String, ByVal strValuemember As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, strTablename)
            cbo.DataSource = mydataset.Tables(strTablename).DefaultView
            cbo.DisplayMember = strdislaymember
            cbo.ValueMember = strValuemember
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CboEmploy_code_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboEmploy_code.SelectedIndexChanged
        If (start) Then
            Try
                txtEmployeeName.Text = CboEmploy_code.SelectedValue
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub cbostations_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbostations.SelectedIndexChanged
        Try
            strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
            Try
                CboEmploy_code.DataSource = Nothing
                Try
                    mydataset.Tables("Tbl_Employee").Clear()
                Catch ex As Exception
                End Try
                CboEmploy_code.Items.Clear()
                FillCombo(CboEmploy_code, strSQL, "Tbl_Employee", "Employ_Code", "Employ_Name")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txttendaily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttendaily.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttendaily.Text = txttendaily.Text.ToUpper
            txtdiachidaily.Focus()
        End If
    End Sub

    Private Sub txtdiachidaily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdiachidaily.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtdiachidaily.Text = txtdiachidaily.Text.ToUpper
            txtTiendoanhso.Focus()
        End If
    End Sub

    Private Sub CboEmploy_code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CboEmploy_code.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cboDsDaiLy.Focus()
        End If
    End Sub

    Private Sub cbostations_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbostations.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerngay.Focus()
        End If
    End Sub


    Private Sub DateTimePickerchuky_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerchuky.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtTiendoanhso.Focus()
        End If
    End Sub

    Private Sub txtTiendoanhso_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTiendoanhso.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttienhoahong.Focus()
        End If
    End Sub


    Private Sub txtslctgocds_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtchitietctgds_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtsophieuhoahong_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsophieuhoahong.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtchitietctgochoahong_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtchitietctgochoahong.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txttienthue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttienthue.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (txtTiendoanhso.Text <> "" And txttienthue.Text <> "" And txttienhoahong.Text <> "") Then
                txtsophaitra.Text = CLng(txtTiendoanhso.Text) + CLng(txttienthue.Text) - CLng(txttienhoahong.Text)
            End If
            cmdLuu.Focus()
        End If
    End Sub

    Private Sub txttienhoahong_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttienhoahong.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttienthue.Focus()
        End If
    End Sub

    Private Sub txtsoPTDS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsoPTDS.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtsophieuthue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsophieuthue.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtsophaitra_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsophaitra.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtsLctgocthue_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If

    End Sub

    Private Sub txtslctgochoahong_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtchitietctgocthue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtchitietctgocthue.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub DateTimePickerngay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerngay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub cmdIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIn.Click
        Dim value
        value = MsgBox("Bạn có thật sự muốn in phiếu thu không?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Nhắc nhở")

        If (Trim$(strPrinterName) = "") Then
            MsgBox("Máy in chưa được thiết lập! Vui lòng chọn máy in rồi tiếp tục")
            Exit Sub
        End If

        If (value = vbYes) Then
            AssignVar()
            Prints()
        End If
    End Sub
    Private Sub prints()
        Try
            ' IN PHIEU THU DOANH SO
            If (CheckBoxPTDS.Checked) Then
                AssignVal(1)
                rpt.PrintToPrinter(NumericUpDownSobanin.Value, True, 1, 1)
            End If
            ' IN PHIEU CHI
            If (CheckBoxPCHH.Checked) Then
                AssignVal(2)
                rptPC.PrintToPrinter(NumericUpDownSobanin.Value, True, 1, 1)
            End If
            ' IN PHIEU THU TNDN
            If (CheckBoxPTTHUE.Checked) Then
                AssignVal(3)
                rpt.PrintToPrinter(NumericUpDownSobanin.Value, True, 1, 1)
            End If
        Catch ex As Exception
            MsgBox("Không tìm thấy máy in " & strPrinterName & ".", MsgBoxStyle.Critical, "Lỗi in")
        End Try
    End Sub
    Private Sub AssignVar()
        rptPC.PrintOptions.PrinterName = strPrinterName
        rpt.PrintOptions.PrinterName = strPrinterName
        ' gan ten cac tham bien vao 
        Arrvar(0) = "txtDonviThu"                   ' Ten don vi
        Arrvar(1) = "txtDiaChiDonVi"                ' Dia chi don  vi
        Arrvar(2) = "txtngay"
        Arrvar(3) = "txtthang"
        Arrvar(4) = "txtnam"
        Arrvar(5) = "txtso"
        Arrvar(6) = "txtQuyen"
        Arrvar(7) = "txttenNguoiNop"
        Arrvar(8) = "txtDiaChiNguoiNop"
        Arrvar(9) = "txtLydonop"
        Arrvar(10) = "txtSotien"
        Arrvar(11) = "txtTienBangChu"
        Arrvar(12) = "txtchungtugoc1"
        Arrvar(13) = "txtchungtugoc2"
        Arrvar(14) = "txttenNguoiNop1"
        Arrvar(15) = "txttenphieu"

        ' gan gia tri cua bien vao
        Arrval(0) = "Đơn vị: " & cbostations.Text
        Arrval(1) = "Địa chỉ: " & GetStringName(cbostations.SelectedValue, "Tbl_Stations")
        Arrval(2) = DateTimePickerngay.Value.Day
        Arrval(3) = DateTimePickerngay.Value.Month
        Arrval(4) = DateTimePickerngay.Value.Year

        Arrval(7) = txttendaily.Text & "(" & cboDsDaiLy.Text & ")"
        Arrval(8) = txtdiachidaily.Text

    End Sub

    Private Sub AssignVal(ByVal index As Integer)
        Dim strMaLoai As String
        'strMaLoai = listitem.SubItems(10).Text
        Arrval(14) = "" 'txtEmployeeName.Text 'Ten nhan vien de trong
        Select Case index
            Case 1
                Arrval(5) = txtsoPTDS.Text  ' So thu tu

                If Trim$(txtsoPTDS.Text) = "" Then
                    Arrval(6) = ""
                Else
                    Dim so As Long
                    so = (CLng(txtsoPTDS.Text) Mod 50)
                    If (so = 0) Then
                        Arrval(6) = CLng(txtsoPTDS.Text) \ 50
                    Else
                        Arrval(6) = (CLng(txtsoPTDS.Text) \ 50) + 1
                    End If
                End If
                Arrval(9) = "Thu tiền cước DS đại lý PSTN. Kỳ cước " & DateTimePickerchuky.Text & ". "   ' Ly do nop

                splitn.strnumbers = CStr(CLng(txtTiendoanhso.Text))
                Arrval(10) = splitn.Splitnumer(",") ' So tien
                numbers.Number = CLng(Arrval(10))

                numbers.Number = CLng(Arrval(10))
                Dim str As String
                str = numbers.NumbersToString
                If (Len(str) > 0) Then
                    str = str.Substring(0, 1).ToUpper + str.Remove(0, 1)
                Else
                    str = "Không"
                End If

                Arrval(11) = str + " đồng." ' tien bang chu
                Arrval(12) = "- 3 Chứng từ gốc" & "  : " & txtchitietctgds.Text & " ."
                Arrval(13) = ""
                Arrval(15) = "PHIẾU THU"
                SetFieldTextOjectReports(rpt, Arrvar, Arrval)
            Case 2
                Arrval(5) = txtsophieuhoahong.Text  ' So thu tu

                If Trim$(txtsophieuhoahong.Text) = "" Then
                    Arrval(6) = ""
                Else
                    Dim so As Long
                    so = (CLng(txtsophieuhoahong.Text) Mod 50)
                    If (so = 0) Then
                        Arrval(6) = CLng(txtsophieuhoahong.Text) \ 50
                    Else
                        Arrval(6) = (CLng(txtsophieuhoahong.Text) \ 50) + 1
                    End If
                End If

                Arrval(9) = "Chi tiền hoa hồng đại lý PSTN. Kỳ cước " & DateTimePickerchuky.Text & ". "   ' Ly do nop
                Arrval(10) = txttienhoahong.Text ' So tien 

                splitn.strnumbers = CStr(CLng(txttienhoahong.Text))
                Arrval(10) = splitn.Splitnumer(",") ' So tien
                numbers.Number = CLng(Arrval(10))

                Dim str As String
                str = numbers.NumbersToString
                If (Len(str) > 0) Then
                    str = str.Substring(0, 1).ToUpper + str.Remove(0, 1)
                Else
                    str = "Không"
                End If

                Arrval(11) = str + " đồng." ' tien bang chu
                Arrval(12) = "- 2 chứng từ gốc" & "  : " & txtchitietctgochoahong.Text & " ."
                Arrval(13) = ""
                'Arrval(15) = "PHIẾU CHI"
                SetFieldTextOjectReports(rptPC, Arrvar, Arrval)
            Case 3

                Arrval(5) = txtsophieuthue.Text  ' So thu tu

                If Trim$(txtsophieuthue.Text) = "" Then
                    Arrval(6) = ""
                Else
                    Dim so As Long
                    so = (CLng(txtsophieuthue.Text) Mod 50)
                    If (so = 0) Then
                        Arrval(6) = CLng(txtsophieuthue.Text) \ 50
                    Else
                        Arrval(6) = (CLng(txtsophieuthue.Text) \ 50) + 1
                    End If
                End If

                Arrval(9) = "Thu tiền thuế TNDN,Đại lý PSTN  . Kỳ cước " & DateTimePickerchuky.Text & ". "   ' Ly do nop
                splitn.strnumbers = CStr(CLng(txttienthue.Text))
                Arrval(10) = splitn.Splitnumer(",") ' So tien
                numbers.Number = CLng(Arrval(10))

                Dim str As String
                str = numbers.NumbersToString
                If (Len(str) > 0) Then
                    str = str.Substring(0, 1).ToUpper + str.Remove(0, 1)
                Else
                    str = "Không"
                End If

                Arrval(11) = str + " đồng." ' tien bang chu
                Arrval(12) = "- 1 chứng từ gốc" & "  : " & txtchitietctgocthue.Text & " ."
                Arrval(13) = ""
                Arrval(15) = "PHIẾU THU"
                SetFieldTextOjectReports(rpt, Arrvar, Arrval)
        End Select

    End Sub

    Private Sub SetFieldTextOjectReports(ByRef rpt As CrystalReport_Receipts, ByVal ArrVar As String(), ByVal ArrVal As String())
        Dim ReportTextObject As CrystalDecisions.CrystalReports.Engine.TextObject
        Dim i As Integer
        Try
            For i = 0 To ArrVar.Length - 1
                ReportTextObject = CType(rpt.ReportDefinition.ReportObjects.Item(ArrVar(i)), CrystalDecisions.CrystalReports.Engine.TextObject)
                ReportTextObject.Text = ArrVal(i)
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SetFieldTextOjectReports(ByRef rpt As CrystalReport_Expense, ByVal ArrVar As String(), ByVal ArrVal As String())
        Dim ReportTextObject As CrystalDecisions.CrystalReports.Engine.TextObject
        Dim i As Integer
        Try
            For i = 0 To ArrVar.Length - 1
                ReportTextObject = CType(rpt.ReportDefinition.ReportObjects.Item(ArrVar(i)), CrystalDecisions.CrystalReports.Engine.TextObject)
                ReportTextObject.Text = ArrVal(i)
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Function GetStringName(ByVal strCode As String, ByVal strtablename As String) As String
        Dim i As Integer
        Dim strresult As String
        For i = 0 To mydataset.Tables(strtablename).Rows.Count - 1
            strresult = mydataset.Tables(strtablename).Rows(i).Item(0)
            If (strresult.Equals(strCode)) Then
                strresult = mydataset.Tables(strtablename).Rows(i).Item(2)
                Exit For
            End If
        Next
        Return strresult
    End Function

    Private Sub cmdLuu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLuu.Click
        If (checkInfo()) Then
            SaveReceipt()
        End If
    End Sub

    Private Function GetMaxNumber(ByVal strQuery As String) As Integer
        Dim results As Integer
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = strQuery
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then

                If Not oleread.IsDBNull(0) Then
                    results = oleread.GetValue(0)
                Else
                    results = 1
                End If

            End If
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :" & ex.ToString)
        End Try
        oledbcon.Close()
        Return results
    End Function

    Private Sub SaveReceipt()
        Try
            'Kiem tra so phieu thu
            If CheckReceiptNo() Then
                Exit Sub
            End If

            Dim soquyen As Long
            'Luu chi tiet phieu thu doanh so
            soquyen = ComputingOrdinal(CLng(txtsoPTDS.Text))

            strSQL = " SELECT MAX(STTDL)+1  FROM Tbl_Receipts "
            MaxDL = GetMaxNumber(strSQL)

            strSQL = " INSERT INTO Tbl_Receipts(Volume,Ordinal_No,Receipt_Date,Service_Code,Descriptions,List_Quantity,List_Detail,Invoice_Quantity,Charge_Cycle,Total_Money,Employ_Code,MaLoaiThu,STTDL,TenDaiLy) " & _
             " VALUES(" & soquyen & _
             "," & Maxnumber & _
             ",'" & DateTimePickerngay.Value.ToShortDateString & _
             "','PSTNDS'" & _
             ",' Thu tiền cước đại lý doanh số PSTN. Kỳ cước " & DateTimePickerchuky.Text & _
             "', 1 " & _
             ",'" & txtchitietctgds.Text & _
             "',1 " & _
             ",'" & DateTimePickerchuky.Text & _
             "'," & CLng(txtTiendoanhso.Text) & _
             ",'" & CboEmploy_code.Text & _
             "','" & CStr(cmbHTthu.SelectedValue) & _
            "'," & MaxDL & _
            ",'" & cboDsDaiLy.Text & "')"

            ExcuxeSQL(strSQL)

            ' Lưu vao so Quy ; Tien Doanh so PSTN
            strSQL = " INSERT INTO Tbl_Receipts_Expenses(Recei_Expen_Date,Recei_No,Descriptions,Recei_Money) " & _
            " VALUES('" & DateTimePickerngay.Value.ToShortDateString & _
            "'," & Maxnumber & _
            ",'Thu tiền cước doanh số DLPSTN:" & cboDsDaiLy.Text & " Kỳ cước " & DateTimePickerchuky.Text & _
            "'," & CLng(txtTiendoanhso.Text) & ")"
            ExcuxeSQL(strSQL)

            ' Phieu Chi  hoa hong
            SaveExpense()

            'Luu tien thue TNDN dai ly PSTN
            soquyen = ComputingOrdinal(CLng(txtsophieuthue.Text))

            strSQL = " INSERT INTO Tbl_Receipts(Volume,Ordinal_No,Receipt_Date,Service_Code,Descriptions,List_Quantity,List_Detail,Invoice_Quantity,Charge_Cycle,Total_Money,Employ_Code,MaLoaiThu,STTDL,TenDaiLy) " & _
             " VALUES(" & soquyen & _
             "," & CLng(txtsophieuthue.Text) & _
             ",'" & DateTimePickerngay.Value.ToShortDateString & _
             "','PSTNTNDN" & _
             "',' Thu tiền cước thuế TNDN DLPSTN:" & cboDsDaiLy.Text & " Kỳ cước " & DateTimePickerchuky.Text & _
             "', 0" & _
             ",'" & txtchitietctgocthue.Text & _
             "', 0 " & _
             ",'" & DateTimePickerchuky.Text & _
             "'," & CLng(txttienthue.Text) & _
             ",'" & CboEmploy_code.Text & _
             "','" & CStr(cmbHTthu.SelectedValue) & _
            "'," & MaxDL & _
            ",'" & cboDsDaiLy.Text & "')"
            ExcuxeSQL(strSQL)

            ' Lưu vao so Quy ; Tien Thue TNDN PSTN Dai Ly
            strSQL = " INSERT INTO Tbl_Receipts_Expenses(Recei_Expen_Date,Recei_No,Descriptions,Recei_Money) " & _
            " VALUES('" & DateTimePickerngay.Value.ToShortDateString & _
            "'," & CLng(txtsophieuthue.Text) & _
            ",'Thu tiền thuế TNDN DLPSTN:" & cboDsDaiLy.Text & " .Kỳ cước " & DateTimePickerchuky.Text & _
            "'," & CLng(txttienthue.Text) & ")"
            ExcuxeSQL(strSQL)

            MsgBox("Đã lưu vào hệ thống")
            cmdLuu.Enabled = False
        Catch ex As Exception
        End Try

    End Sub

    Private Function CheckReceiptNo() As Boolean
        Dim SoPT As Long
        Dim strquery As String
        Dim result As Boolean = False
        SoPT = txtsoPTDS.Text
        strquery = " SELECT Ordinal_No FROM Tbl_Receipts WHERE MONTH(Receipt_Date) =" & DateTimePickerngay.Value.Month & " AND YEAR(Receipt_Date) =" & DateTimePickerngay.Value.Year & " AND Ordinal_No = " & SoPT & " AND MaLoaiThu ='" & CStr(cmbHTthu.SelectedValue) & "'"
        If (CheckReciept_No(strQuery)) Then
            result = True
            MsgBox("Số phiếu thu :" & SoPT & " đã tồn tại. Vui lòng kiểm tra lại.", MsgBoxStyle.Critical)
        End If
        Return result
    End Function

    Private Function ComputingOrdinal(ByVal so As Long) As Long
        Dim result As Long
        Dim sodu As Long
        sodu = CLng(so Mod 50)
        If (sodu = 0) Then
            result = CLng(txtsoPTDS.Text) \ 50
        Else
            result = (CLng(txtsoPTDS.Text) \ 50) + 1
        End If

        Return result
    End Function

    Private Sub SaveExpense()

        Dim soquyen As Long
        'Luu chi tiet phieu chi hoa hong dai ly 
        soquyen = ComputingOrdinal(CLng(txtsophieuhoahong.Text))

        strSQL = " INSERT INTO Tbl_Expenses(Volume,Ordinal_No,Expense_Date,Service_Code,Descriptions,List_Quantity,List_Detail,Charge_Cycle,Total_Money,Employ_Code,List_Detail_Receipt_No,STTDL,Status) " & _
        " VALUES(" & soquyen & _
        "," & CLng(txtsophieuhoahong.Text) & _
        ",'" & DateTimePickerngay.Value.ToShortDateString & _
        "','PSTNHH" & _
        "','" & " Chi trả hoa hồng DLPSTN :" & cboDsDaiLy.Text & ". Kỳ cước " & DateTimePickerchuky.Text & _
        "', 0" & _
        ",'" & txtchitietctgochoahong.Text & _
        "','" & DateTimePickerchuky.Text & _
        "'," & CLng(txttienhoahong.Text) & _
        ",'" & CboEmploy_code.Text & _
        "','" & txtsoPTDS.Text & "," & txtsophieuthue.Text & _
        "'," & MaxDL & " ,True )"
        ExcuxeSQL(strSQL)

        ' Lưu vao so Quy ; Tien Chi Hoa Hong
        strSQL = " INSERT INTO Tbl_Receipts_Expenses(Recei_Expen_Date,Expen_No,Descriptions,Expen_Money) " & _
       " VALUES('" & DateTimePickerngay.Value.ToShortDateString & _
       "'," & CLng(txtsophieuhoahong.Text) & _
       ",'" & " Chi trả hoa hồng DLPSTN :" & cboDsDaiLy.Text & ". Kỳ cước " & DateTimePickerchuky.Text & _
       "'," & CLng(txttienhoahong.Text) & " )"

        ExcuxeSQL(strSQL)

        'Cap nhat lai phieu thu
        '

    End Sub

    Private Sub txtsoPTDS_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsoPTDS.Validated
        If (Trim$(txtsoPTDS.Text <> "")) Then
            txtsophieuthue.Text = CLng(txtsoPTDS.Text) + 1
        Else
            txtsophieuthue.Text = ""
        End If
    End Sub

    Private Sub cmddeletes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddeletes.Click
        cmdLuu.Enabled = True
        txttendaily.Text = vbNullString
        txtdiachidaily.Text = vbNullString
        txtTiendoanhso.Text = vbNullString
        txttienhoahong.Text = vbNullString
        txttienthue.Text = vbNullString
        txtsophaitra.Text = vbNullString

        strSQL = " SELECT MAX(Ordinal_No)+1  FROM Tbl_Receipts WHERE  MaLoaiThu = 'TM'" & " AND ( MONTH(Receipt_Date) = " & DateTimePickerngay.Value.Month & ") AND  (YEAR(Receipt_Date) = " & DateTimePickerngay.Value.Year() & ")"
        Maxnumber = GetMaxNumber(strSQL)
        txtsoPTDS.Text = Maxnumber


        strSQL = " SELECT MAX(Ordinal_No)+1  FROM Tbl_Expenses WHERE  ( MONTH(Expense_Date) = " & DateTimePickerngay.Value.Month & ") AND  (YEAR(Expense_Date) = " & DateTimePickerngay.Value.Year() & ")"
        txtsophieuhoahong.Text = GetMaxNumber(strSQL)
        txtsophieuthue.Text = Maxnumber + 1

    End Sub

    Private Sub txttienthue_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txttienthue.Validated
        If (txtTiendoanhso.Text <> "" And txttienthue.Text <> "" And txttienhoahong.Text <> "") Then
            txtsophaitra.Text = CLng(txtTiendoanhso.Text) + CLng(txttienthue.Text) - CLng(txttienhoahong.Text)
        End If
    End Sub

    Private Sub txtTiendoanhso_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTiendoanhso.Validated
        If (txtTiendoanhso.Text <> "" And IsNumeric(txtTiendoanhso.Text)) Then
            txtchitietctgochoahong.Text = "1. Phiếu thu số " & txtsoPTDS.Text & "; 1 phiếu thanh toán ĐL  "
        Else
            MsgBox("Tiền phải là kiểu số")
            txtTiendoanhso.Focus()
        End If
    End Sub

    Private Sub txtchitietctgds_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtchitietctgds.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttienhoahong.Focus()
        End If
    End Sub

    Private Function checkInfo() As Boolean
        Dim result As Boolean = True

        If (Trim$(CboEmploy_code.Text) = "") Then
            MsgBox("Tên CTV chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            CboEmploy_code.Focus()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txttendaily.Text) = "") Then
            MsgBox("Tên đại lý chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txttendaily.Focus()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txtdiachidaily.Text) = "") Then
            MsgBox("Địa chỉ đại lý chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtdiachidaily.Focus()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txtTiendoanhso.Text) = "") Then
            MsgBox("Tiền doanh số chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTiendoanhso.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtTiendoanhso.Text))) Then
            MsgBox("Tiền doanh số phải là số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTiendoanhso.Focus()
            txtTiendoanhso.SelectAll()
            result = False
            GoTo endFunction
        End If


        If (Trim$(txttienhoahong.Text) = "") Then
            MsgBox("Tiền hoa hồng chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txttienhoahong.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txttienhoahong.Text))) Then
            MsgBox("Tiền hoa hồng phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txttienhoahong.Focus()
            txttienhoahong.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txttienthue.Text) = "") Then
            MsgBox("Tiền thuế TNDN chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txttienthue.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txttienthue.Text))) Then
            MsgBox("Tiền thuế TNDN phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txttienthue.Focus()
            txttienthue.SelectAll()
            result = False
            GoTo endFunction
        End If


        If (Trim$(txtsoPTDS.Text) = "") Then
            MsgBox("Số phiếu thu chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsoPTDS.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtsoPTDS.Text))) Then
            MsgBox("Số phiếu thu phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsoPTDS.Focus()
            txtsoPTDS.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (CLng(Trim$(txtsoPTDS.Text)) < 1) Then
            MsgBox("Số phiếu thu phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsoPTDS.Focus()
            txtsoPTDS.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txtsophieuhoahong.Text) = "") Then
            MsgBox("Số phiếu chi chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsophieuhoahong.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtsophieuhoahong.Text))) Then
            MsgBox("Số phiếu chi phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsophieuhoahong.Focus()
            txtsophieuhoahong.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (CLng(Trim$(txtsophieuhoahong.Text)) < 1) Then
            MsgBox("Số phiếu chi phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsophieuhoahong.Focus()
            txtsophieuhoahong.SelectAll()
            result = False
            GoTo endFunction
        End If
endFunction:
        Return result
    End Function

    Private Sub txttienhoahong_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txttienhoahong.Validated
        If (Trim$(txttienhoahong.Text) <> "" AndAlso IsNumeric(txttienhoahong.Text)) Then
            txttienthue.Text = CLng(txttienhoahong.Text) * 0.05
        Else
            MsgBox("Nhập liệu không hợp lệ!")
        End If

    End Sub

    Private Sub frmThanhtoanDL_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        strSQL = "SELECT * FROM Tbl_DaiLy "
        FillDataSet(strSQL)
        If (cboDsDaiLy.Items.Count > 0) Then
            cboDsDaiLy.SelectedIndex = 0
        End If

        If (cmbHTthu.Items.Count > 0) Then
            If (cmbHTthu.FindString("TIỀN MẶT") > 0) Then
                cmbHTthu.SelectedIndex = cmbHTthu.FindString("TIỀN MẶT")
            End If
        End If
    End Sub

    Private Sub FillDataset(ByVal strQuery As String)
        Try
            dsdl = New DataSet
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(dsdl, "Tbl_DaiLy")

            Dim i As Integer
            For i = 0 To dsdl.Tables("Tbl_DaiLy").Rows.Count - 1
                cboDsDaiLy.Items.Add(dsdl.Tables("Tbl_DaiLy").Rows(i).Item("MaDL"))
            Next
        Catch ex As Exception
            MsgBox("Error Fill DataSet :" & ex.ToString)
        End Try
    End Sub

    Private Sub FillTextBox(ByVal strMa As String)
        Dim i As Integer
        For i = 0 To dsdl.Tables("Tbl_DaiLy").Rows.Count - 1
            If (strMa = dsdl.Tables("Tbl_DaiLy").Rows(i).Item("MaDL")) Then
                txttendaily.Text = dsdl.Tables("Tbl_DaiLy").Rows(i).Item("TenDL")
                txtdiachidaily.Text = dsdl.Tables("Tbl_DaiLy").Rows(i).Item("DiaChi")
                Exit Sub
            End If
        Next
    End Sub

    Private Sub cboDsDaiLy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDsDaiLy.SelectedIndexChanged
        FillTextBox(cboDsDaiLy.Text)
    End Sub

    Private Sub cboDsDaiLy_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboDsDaiLy.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            Dim index As Integer
            index = cboDsDaiLy.FindString(cboDsDaiLy.Text)
            If (index > 0) Then
                cboDsDaiLy.SelectedIndex = index
            End If
            DateTimePickerchuky.Focus()
        End If
    End Sub
End Class
