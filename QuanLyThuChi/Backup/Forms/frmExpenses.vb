Imports ConvertNumberToChar
Imports System.Data.OleDb
Public Class frmExpenses
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private Indexlistview As Integer
    Private strListID As String
    Private strListRecei_No As String
    Dim start As Boolean = False
    Dim splitn As New SplitNumbers
    Dim numbers As New ConvertNumbersToString
    Dim rpt As CrystalReport_Expense
    Dim SaveFlag As Boolean = False
    Dim Arrvar(14) As String
    Dim Arrval(14) As String
    Dim ArrListReceiID() As String
    Dim MaxIDExpense As Long
    Friend WithEvents ThePrintDocument As System.Drawing.Printing.PrintDocument
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        DateTimePickerngay.Value = Now
        DateTimePickerchuky.Value = Now
        SetDetailListView()
        Try
            mydataset = New DataSet
            FillDataSet()
        Catch eLoad As System.Exception
            System.Windows.Forms.MessageBox.Show(eLoad.Message)
        End Try
        start = True

        If (cbostations.Items.Count > 0) Then
            cbostations.SelectedIndex = 0
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
        End If

        If (CboEmploy_code.Items.Count > 0) Then
            CboEmploy_code.SelectedIndex = 0
            txtEmployeeName.Text = CboEmploy_code.SelectedValue
        End If

        If (Cbolydo.Items.Count > 0) Then
            Cbolydo.SelectedIndex = 0
            txtlydo.Text = "Chi nộp " & Cbolydo.SelectedValue
        End If

        
        txtEmployeeName.ReadOnly = True
        rpt = New CrystalReport_Expense
        'dim size as  CrystalDecisions.Shared.PaperSize =     
        'rpt.PrintOptions.PaperSize 
        'rpt.PrintOptions.PaperOrientation = CrystalDecisions.[Shared].PaperOrientation.Portrait


    End Sub

    Public Sub New(ByVal strStation As String, ByVal strService As String, ByVal SLCT As Integer, ByVal strSLCTDetail As String, ByVal SoTien As String, ByVal Chuky As String, ByVal _strListID As String, ByVal _strsobknt As String, ByVal _strListPT As String, ByVal _EmployeeCode As String, ByVal _NgayNop As Date, ByVal Status As Boolean)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call

        DateTimePickerngay.Value = _NgayNop
        'DateTimePickerchuky.Value = Chuky
        txtChuky.Text = Chuky 'DateTimePickerchuky.Value.Month & "/" & DateTimePickerchuky.Value.Year
        txtngaynop.Text = DateTimePickerngay.Value.ToShortDateString
        SetDetailListView()
        Try
            mydataset = New DataSet
            FillDataSet()
        Catch eLoad As System.Exception
            System.Windows.Forms.MessageBox.Show(eLoad.Message)
        End Try
        start = True

        If (cbostations.Items.Count > 0) Then
            cbostations.SelectedIndex = cbostations.FindString(strStation)
            txtTram.Text = strStation
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
        End If

        If (CboEmploy_code.Items.Count > 0) Then
            CboEmploy_code.SelectedIndex = CboEmploy_code.FindString(_EmployeeCode)
            txtEmployeeName.Text = CboEmploy_code.SelectedValue
            txtMaNguoiNop.Text = _EmployeeCode
        End If

        If (Cbolydo.Items.Count > 0) Then
            Cbolydo.SelectedIndex = Cbolydo.FindString(strService)
            txtDichvu.Text = Cbolydo.Text
            txtlydo.Text = "Chi nộp " & Cbolydo.SelectedValue
        End If

        If (Status) Then
            cmdLuu.Enabled = False
        End If

        GetNumber_Vol_Ord()
        Computing_Vol_Ord()
        txtso.Focus()
        txtsotien.Text = SoTien
        txtsoBK.Text = SLCT
        txtchitietbk.Text = strSLCTDetail
        txtSoBKNT.Text = _strsobknt
        strListID = _strListID
        ArrListReceiID = strListID.Split(",")
        txtListDetailReceiNo.Text = _strListPT
        strListRecei_No = _strListPT
        txtEmployeeName.ReadOnly = True
        rpt = New CrystalReport_Expense
        AddSubItemListView()
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
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDownSobanin As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtso As System.Windows.Forms.TextBox
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents txtsoBK As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePickerchuky As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdPreview As System.Windows.Forms.Button
    Friend WithEvents cmdThemPT As System.Windows.Forms.Button
    Friend WithEvents cmdbotPT As System.Windows.Forms.Button
    Friend WithEvents cmdIn As System.Windows.Forms.Button
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents cmdLuu As System.Windows.Forms.Button
    Friend WithEvents cmddeletes As System.Windows.Forms.Button
    Friend WithEvents txtquyen As System.Windows.Forms.TextBox
    Friend WithEvents txtsotien As System.Windows.Forms.TextBox
    Friend WithEvents txtchitietbk As System.Windows.Forms.TextBox
    Friend WithEvents txttongtien As System.Windows.Forms.TextBox
    Friend WithEvents ListViewDetail As System.Windows.Forms.ListView
    Friend WithEvents DateTimePickerngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtlydo As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtListDetailReceiNo As System.Windows.Forms.TextBox
    Friend WithEvents lbltongDV As System.Windows.Forms.Label
    Friend WithEvents txtSoBKNT As System.Windows.Forms.TextBox
    Friend WithEvents txtMaNguoiNop As System.Windows.Forms.TextBox
    Friend WithEvents txtChuky As System.Windows.Forms.TextBox
    Friend WithEvents txtngaynop As System.Windows.Forms.TextBox
    Friend WithEvents txtDichvu As System.Windows.Forms.TextBox
    Friend WithEvents txtTram As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmExpenses))
        Me.Label15 = New System.Windows.Forms.Label
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtDichvu = New System.Windows.Forms.TextBox
        Me.txtngaynop = New System.Windows.Forms.TextBox
        Me.txtChuky = New System.Windows.Forms.TextBox
        Me.txtMaNguoiNop = New System.Windows.Forms.TextBox
        Me.DateTimePickerngay = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.NumericUpDownSobanin = New System.Windows.Forms.NumericUpDown
        Me.Label9 = New System.Windows.Forms.Label
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtso = New System.Windows.Forms.TextBox
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.txtsoBK = New System.Windows.Forms.TextBox
        Me.DateTimePickerchuky = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdPreview = New System.Windows.Forms.Button
        Me.cmdThemPT = New System.Windows.Forms.Button
        Me.cmdbotPT = New System.Windows.Forms.Button
        Me.cmdIn = New System.Windows.Forms.Button
        Me.cmddong = New System.Windows.Forms.Button
        Me.cmdLuu = New System.Windows.Forms.Button
        Me.cmddeletes = New System.Windows.Forms.Button
        Me.txtquyen = New System.Windows.Forms.TextBox
        Me.txtsotien = New System.Windows.Forms.TextBox
        Me.txtchitietbk = New System.Windows.Forms.TextBox
        Me.txtSoBKNT = New System.Windows.Forms.TextBox
        Me.txttongtien = New System.Windows.Forms.TextBox
        Me.ListViewDetail = New System.Windows.Forms.ListView
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtlydo = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.lbltongDV = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtListDetailReceiNo = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtTram = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDownSobanin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label15
        '
        Me.Label15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label15.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label15.Font = New System.Drawing.Font("Arial", 24.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label15.Location = New System.Drawing.Point(4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(715, 46)
        Me.Label15.TabIndex = 52
        Me.Label15.Text = "NHẬP PHIẾU CHI                         "
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(528, 18)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(192, 27)
        Me.cbostations.TabIndex = 53
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtDichvu)
        Me.GroupBox1.Controls.Add(Me.txtngaynop)
        Me.GroupBox1.Controls.Add(Me.txtChuky)
        Me.GroupBox1.Controls.Add(Me.txtMaNguoiNop)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerngay)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.NumericUpDownSobanin)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.txtso)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.txtsoBK)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerchuky)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.txtquyen)
        Me.GroupBox1.Controls.Add(Me.txtsotien)
        Me.GroupBox1.Controls.Add(Me.txtchitietbk)
        Me.GroupBox1.Controls.Add(Me.txtSoBKNT)
        Me.GroupBox1.Controls.Add(Me.txttongtien)
        Me.GroupBox1.Controls.Add(Me.ListViewDetail)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtlydo)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.lbltongDV)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.txtListDetailReceiNo)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(5, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 422)
        Me.GroupBox1.TabIndex = 51
        Me.GroupBox1.TabStop = False
        '
        'txtDichvu
        '
        Me.txtDichvu.BackColor = System.Drawing.Color.White
        Me.txtDichvu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDichvu.ForeColor = System.Drawing.Color.FromArgb(CType(64, Byte), CType(0, Byte), CType(64, Byte))
        Me.txtDichvu.Location = New System.Drawing.Point(77, 56)
        Me.txtDichvu.Name = "txtDichvu"
        Me.txtDichvu.ReadOnly = True
        Me.txtDichvu.Size = New System.Drawing.Size(99, 26)
        Me.txtDichvu.TabIndex = 58
        Me.txtDichvu.Text = ""
        '
        'txtngaynop
        '
        Me.txtngaynop.BackColor = System.Drawing.Color.White
        Me.txtngaynop.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtngaynop.ForeColor = System.Drawing.Color.FromArgb(CType(64, Byte), CType(0, Byte), CType(64, Byte))
        Me.txtngaynop.Location = New System.Drawing.Point(77, 18)
        Me.txtngaynop.Name = "txtngaynop"
        Me.txtngaynop.ReadOnly = True
        Me.txtngaynop.Size = New System.Drawing.Size(99, 26)
        Me.txtngaynop.TabIndex = 57
        Me.txtngaynop.Text = ""
        '
        'txtChuky
        '
        Me.txtChuky.BackColor = System.Drawing.Color.White
        Me.txtChuky.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChuky.ForeColor = System.Drawing.Color.FromArgb(CType(64, Byte), CType(0, Byte), CType(64, Byte))
        Me.txtChuky.Location = New System.Drawing.Point(77, 115)
        Me.txtChuky.Name = "txtChuky"
        Me.txtChuky.ReadOnly = True
        Me.txtChuky.Size = New System.Drawing.Size(99, 26)
        Me.txtChuky.TabIndex = 56
        Me.txtChuky.Text = ""
        '
        'txtMaNguoiNop
        '
        Me.txtMaNguoiNop.BackColor = System.Drawing.Color.White
        Me.txtMaNguoiNop.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaNguoiNop.ForeColor = System.Drawing.Color.Blue
        Me.txtMaNguoiNop.Location = New System.Drawing.Point(259, 18)
        Me.txtMaNguoiNop.Name = "txtMaNguoiNop"
        Me.txtMaNguoiNop.ReadOnly = True
        Me.txtMaNguoiNop.Size = New System.Drawing.Size(188, 26)
        Me.txtMaNguoiNop.TabIndex = 54
        Me.txtMaNguoiNop.Text = ""
        '
        'DateTimePickerngay
        '
        Me.DateTimePickerngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngay.Enabled = False
        Me.DateTimePickerngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngay.Location = New System.Drawing.Point(77, 18)
        Me.DateTimePickerngay.Name = "DateTimePickerngay"
        Me.DateTimePickerngay.Size = New System.Drawing.Size(99, 26)
        Me.DateTimePickerngay.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 22)
        Me.Label3.TabIndex = 53
        Me.Label3.Text = "Ngày/T/N"
        '
        'NumericUpDownSobanin
        '
        Me.NumericUpDownSobanin.Location = New System.Drawing.Point(654, 220)
        Me.NumericUpDownSobanin.Name = "NumericUpDownSobanin"
        Me.NumericUpDownSobanin.Size = New System.Drawing.Size(48, 26)
        Me.NumericUpDownSobanin.TabIndex = 51
        Me.NumericUpDownSobanin.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(584, 222)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 22)
        Me.Label9.TabIndex = 52
        Me.Label9.Text = "Số bản in "
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Enabled = False
        Me.CboEmploy_code.Location = New System.Drawing.Point(260, 18)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(188, 27)
        Me.CboEmploy_code.TabIndex = 2
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label14.Location = New System.Drawing.Point(4, 48)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(700, 2)
        Me.Label14.TabIndex = 34
        '
        'txtso
        '
        Me.txtso.BackColor = System.Drawing.Color.White
        Me.txtso.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtso.ForeColor = System.Drawing.Color.Black
        Me.txtso.Location = New System.Drawing.Point(77, 86)
        Me.txtso.Name = "txtso"
        Me.txtso.Size = New System.Drawing.Size(99, 26)
        Me.txtso.TabIndex = 4
        Me.txtso.Text = ""
        Me.txtso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Cbolydo
        '
        Me.Cbolydo.Enabled = False
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(77, 56)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(99, 27)
        Me.Cbolydo.TabIndex = 3
        '
        'txtsoBK
        '
        Me.txtsoBK.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsoBK.ForeColor = System.Drawing.Color.Black
        Me.txtsoBK.Location = New System.Drawing.Point(77, 144)
        Me.txtsoBK.Name = "txtsoBK"
        Me.txtsoBK.Size = New System.Drawing.Size(99, 26)
        Me.txtsoBK.TabIndex = 8
        Me.txtsoBK.Text = ""
        Me.txtsoBK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DateTimePickerchuky
        '
        Me.DateTimePickerchuky.CustomFormat = "MM/yyyy"
        Me.DateTimePickerchuky.Enabled = False
        Me.DateTimePickerchuky.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerchuky.Location = New System.Drawing.Point(77, 115)
        Me.DateTimePickerchuky.Name = "DateTimePickerchuky"
        Me.DateTimePickerchuky.Size = New System.Drawing.Size(99, 26)
        Me.DateTimePickerchuky.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(176, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 22)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Quyển số"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdPreview)
        Me.GroupBox2.Controls.Add(Me.cmdThemPT)
        Me.GroupBox2.Controls.Add(Me.cmdbotPT)
        Me.GroupBox2.Controls.Add(Me.cmdIn)
        Me.GroupBox2.Controls.Add(Me.cmddong)
        Me.GroupBox2.Controls.Add(Me.cmdLuu)
        Me.GroupBox2.Controls.Add(Me.cmddeletes)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 170)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(696, 48)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        '
        'cmdPreview
        '
        Me.cmdPreview.Location = New System.Drawing.Point(431, 14)
        Me.cmdPreview.Name = "cmdPreview"
        Me.cmdPreview.Size = New System.Drawing.Size(80, 28)
        Me.cmdPreview.TabIndex = 52
        Me.cmdPreview.Text = "Xem"
        '
        'cmdThemPT
        '
        Me.cmdThemPT.Enabled = False
        Me.cmdThemPT.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdThemPT.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdThemPT.Location = New System.Drawing.Point(16, 14)
        Me.cmdThemPT.Name = "cmdThemPT"
        Me.cmdThemPT.Size = New System.Drawing.Size(80, 28)
        Me.cmdThemPT.TabIndex = 12
        Me.cmdThemPT.Text = "Thêm PC"
        '
        'cmdbotPT
        '
        Me.cmdbotPT.Enabled = False
        Me.cmdbotPT.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdbotPT.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdbotPT.Location = New System.Drawing.Point(99, 14)
        Me.cmdbotPT.Name = "cmdbotPT"
        Me.cmdbotPT.Size = New System.Drawing.Size(80, 28)
        Me.cmdbotPT.TabIndex = 23
        Me.cmdbotPT.Text = "Bớt PC"
        '
        'cmdIn
        '
        Me.cmdIn.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdIn.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdIn.Location = New System.Drawing.Point(182, 14)
        Me.cmdIn.Name = "cmdIn"
        Me.cmdIn.Size = New System.Drawing.Size(80, 28)
        Me.cmdIn.TabIndex = 23
        Me.cmdIn.Text = "In"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(608, 14)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(80, 28)
        Me.cmddong.TabIndex = 31
        Me.cmddong.Text = "Đóng"
        '
        'cmdLuu
        '
        Me.cmdLuu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuu.Location = New System.Drawing.Point(265, 14)
        Me.cmdLuu.Name = "cmdLuu"
        Me.cmdLuu.Size = New System.Drawing.Size(80, 28)
        Me.cmdLuu.TabIndex = 23
        Me.cmdLuu.Text = "Lưu"
        '
        'cmddeletes
        '
        Me.cmddeletes.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddeletes.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddeletes.Location = New System.Drawing.Point(348, 14)
        Me.cmddeletes.Name = "cmddeletes"
        Me.cmddeletes.Size = New System.Drawing.Size(80, 28)
        Me.cmddeletes.TabIndex = 23
        Me.cmddeletes.Text = "Xóa DS"
        '
        'txtquyen
        '
        Me.txtquyen.BackColor = System.Drawing.Color.White
        Me.txtquyen.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtquyen.ForeColor = System.Drawing.Color.Black
        Me.txtquyen.Location = New System.Drawing.Point(248, 86)
        Me.txtquyen.Name = "txtquyen"
        Me.txtquyen.ReadOnly = True
        Me.txtquyen.Size = New System.Drawing.Size(160, 26)
        Me.txtquyen.TabIndex = 5
        Me.txtquyen.Text = ""
        Me.txtquyen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtsotien
        '
        Me.txtsotien.BackColor = System.Drawing.Color.White
        Me.txtsotien.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsotien.ForeColor = System.Drawing.Color.Black
        Me.txtsotien.Location = New System.Drawing.Point(248, 115)
        Me.txtsotien.Name = "txtsotien"
        Me.txtsotien.ReadOnly = True
        Me.txtsotien.Size = New System.Drawing.Size(160, 26)
        Me.txtsotien.TabIndex = 7
        Me.txtsotien.Text = ""
        Me.txtsotien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtchitietbk
        '
        Me.txtchitietbk.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtchitietbk.ForeColor = System.Drawing.Color.Black
        Me.txtchitietbk.Location = New System.Drawing.Point(248, 144)
        Me.txtchitietbk.Name = "txtchitietbk"
        Me.txtchitietbk.Size = New System.Drawing.Size(160, 26)
        Me.txtchitietbk.TabIndex = 9
        Me.txtchitietbk.Text = ""
        '
        'txtSoBKNT
        '
        Me.txtSoBKNT.BackColor = System.Drawing.Color.White
        Me.txtSoBKNT.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoBKNT.ForeColor = System.Drawing.Color.Brown
        Me.txtSoBKNT.Location = New System.Drawing.Point(536, 86)
        Me.txtSoBKNT.Name = "txtSoBKNT"
        Me.txtSoBKNT.ReadOnly = True
        Me.txtSoBKNT.Size = New System.Drawing.Size(168, 26)
        Me.txtSoBKNT.TabIndex = 50
        Me.txtSoBKNT.Text = ""
        Me.txtSoBKNT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txttongtien
        '
        Me.txttongtien.BackColor = System.Drawing.Color.White
        Me.txttongtien.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttongtien.ForeColor = System.Drawing.Color.Brown
        Me.txttongtien.Location = New System.Drawing.Point(536, 115)
        Me.txttongtien.Name = "txttongtien"
        Me.txttongtien.ReadOnly = True
        Me.txttongtien.Size = New System.Drawing.Size(168, 26)
        Me.txttongtien.TabIndex = 50
        Me.txttongtien.Text = ""
        Me.txttongtien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ListViewDetail
        '
        Me.ListViewDetail.FullRowSelect = True
        Me.ListViewDetail.GridLines = True
        Me.ListViewDetail.Location = New System.Drawing.Point(8, 247)
        Me.ListViewDetail.Name = "ListViewDetail"
        Me.ListViewDetail.Size = New System.Drawing.Size(696, 169)
        Me.ListViewDetail.TabIndex = 28
        Me.ListViewDetail.View = System.Windows.Forms.View.Details
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(448, 18)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.ReadOnly = True
        Me.txtEmployeeName.Size = New System.Drawing.Size(256, 26)
        Me.txtEmployeeName.TabIndex = 50
        Me.txtEmployeeName.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 22)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Số PC"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(184, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(84, 22)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Người nhận"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 58)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Dịch vụ"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(176, 117)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(52, 22)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Số tiền"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 146)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(78, 22)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "SL CT gốc"
        '
        'txtlydo
        '
        Me.txtlydo.BackColor = System.Drawing.Color.White
        Me.txtlydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtlydo.ForeColor = System.Drawing.Color.Black
        Me.txtlydo.Location = New System.Drawing.Point(184, 56)
        Me.txtlydo.Name = "txtlydo"
        Me.txtlydo.ReadOnly = True
        Me.txtlydo.Size = New System.Drawing.Size(520, 26)
        Me.txtlydo.TabIndex = 50
        Me.txtlydo.Text = ""
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(176, 146)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(80, 22)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "Chi tiết CT"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(5, 117)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(54, 22)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Chu kỳ"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType(((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic) _
                        Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Brown
        Me.Label12.Location = New System.Drawing.Point(9, 222)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(239, 23)
        Me.Label12.TabIndex = 25
        Me.Label12.Text = "Danh sách chi tiết các phiếu chi"
        '
        'lbltongDV
        '
        Me.lbltongDV.AutoSize = True
        Me.lbltongDV.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbltongDV.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lbltongDV.Location = New System.Drawing.Point(408, 88)
        Me.lbltongDV.Name = "lbltongDV"
        Me.lbltongDV.Size = New System.Drawing.Size(136, 22)
        Me.lbltongDV.TabIndex = 25
        Me.lbltongDV.Text = "Số BK nộp tiền NH"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(408, 117)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(69, 22)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Tổng tiền"
        '
        'txtListDetailReceiNo
        '
        Me.txtListDetailReceiNo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtListDetailReceiNo.ForeColor = System.Drawing.Color.Brown
        Me.txtListDetailReceiNo.Location = New System.Drawing.Point(480, 144)
        Me.txtListDetailReceiNo.Name = "txtListDetailReceiNo"
        Me.txtListDetailReceiNo.Size = New System.Drawing.Size(224, 26)
        Me.txtListDetailReceiNo.TabIndex = 50
        Me.txtListDetailReceiNo.Text = ""
        Me.txtListDetailReceiNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(408, 146)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(78, 22)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Chi tiết PT"
        '
        'txtTram
        '
        Me.txtTram.BackColor = System.Drawing.Color.White
        Me.txtTram.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTram.ForeColor = System.Drawing.Color.Blue
        Me.txtTram.Location = New System.Drawing.Point(526, 18)
        Me.txtTram.Name = "txtTram"
        Me.txtTram.ReadOnly = True
        Me.txtTram.Size = New System.Drawing.Size(192, 26)
        Me.txtTram.TabIndex = 54
        Me.txtTram.Text = ""
        '
        'frmExpenses
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 470)
        Me.Controls.Add(Me.txtTram)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmExpenses"
        Me.Text = "Nhập phiếu chi"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.NumericUpDownSobanin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub FillDataSet()

        strSQL = "SELECT StationID,Station_Name,Station_Address FROM Tbl_Stations "
        FillCombo(cbostations, strSQL, "Tbl_Stations", "Station_Name", "StationID")

        strSQL = "SELECT Service_Code,Service_Name FROM Tbl_Services "
        FillCombo(Cbolydo, strSQL, "Tbl_Services", "Service_Code", "Service_Name")

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

    Private Sub SetDetailListView()
        ListViewDetail.Items.Clear()
        ListViewDetail.Columns.Clear()
        ListViewDetail.Columns.Add("STT", 50, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Dịch vụ", 80, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Chu kỳ", 70, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("SL CT", 60, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Chi tiết CT", 140, HorizontalAlignment.Left)
        ListViewDetail.Columns.Add("Tiền", 80, HorizontalAlignment.Right)
        ListViewDetail.Columns.Add("Mô tả", 180, HorizontalAlignment.Left)
        ListViewDetail.Columns.Add("Q Số", 50, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Số", 50, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("DS số PT", 80, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Số BKNT", 50, HorizontalAlignment.Center)

    End Sub

    Private Sub AddSubItemListView()

        ListViewDetail.Items.Add(ListViewDetail.Items.Count + 1)
        ListViewDetail.Items(Indexlistview).SubItems.Add(Cbolydo.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtChuky.Text)  'DateTimePickerchuky.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtsoBK.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtchitietbk.Text)
        splitn.strnumbers = CStr(CLng(txtsotien.Text))
        ListViewDetail.Items(Indexlistview).SubItems.Add(splitn.Splitnumer(","))
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtlydo.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtquyen.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtso.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(strListRecei_No)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtSoBKNT.Text)
        Indexlistview += 1

    End Sub

    Private Sub SaveExpense(ByVal lvitem As ListViewItem)
        strSQL = " INSERT INTO Tbl_Expenses(Volume,Ordinal_No,Expense_Date,Service_Code,Descriptions,List_Quantity,List_Detail,Charge_Cycle,Total_Money,Employ_Code,List_Detail_Receipt_No,Ordinal_No_List) " & _
        " VALUES(" & CInt(Trim$(lvitem.SubItems(7).Text)) & _
        "," & CInt(Trim$(lvitem.SubItems(8).Text)) & _
        ",'" & DateTimePickerngay.Value.ToShortDateString & _
        "','" & lvitem.SubItems(1).Text & _
        "','" & lvitem.SubItems(6).Text & _
        "'," & CInt(Trim$(lvitem.SubItems(3).Text)) & _
        ",'" & lvitem.SubItems(4).Text & _
        "','" & lvitem.SubItems(2).Text & _
        "'," & CLng(lvitem.SubItems(5).Text) & _
        ",'" & CboEmploy_code.Text & _
        "','" & lvitem.SubItems(9).Text & _
        "'," & CInt(lvitem.SubItems(10).Text) & ")"
           ExcuxeSQL(strSQL)
    End Sub

    Private Sub Resort()
        Dim i As Integer
        For i = 0 To ListViewDetail.Items.Count - 1
            ListViewDetail.Items(i).SubItems(0).Text = CStr(i + 1)
        Next
    End Sub

    Private Sub cmdThemPT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdThemPT.Click
        If checkInfo() Then
            AddSubItemListView()
            Cbolydo.Focus()
            txtso.Text = CInt(txtso.Text) + 1
            Computing_Vol_Ord()
            txttongtien.Text = SumColumnItem(5)
            'txtTongtienDv.Text = SumColumnItem(5, Cbolydo.Text)
            DeleteTextBox()
        End If

    End Sub

    Private Sub cmdbotPT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdbotPT.Click
        Dim value
        Dim lvitem As ListViewItem
        Try
            lvitem = ListViewDetail.FocusedItem
            If (lvitem.SubItems.Item(0).Text <> "") Then
                value = MsgBox("Bạn có thật sự muốn xóa dòng này không?", MsgBoxStyle.YesNo)
                If (value = vbYes) Then
                    ListViewDetail.Items.Remove(lvitem)
                    Indexlistview -= 1
                    Resort()
                    txttongtien.Text = SumColumnItem(5)
                    'txtTongtienDv.Text = SumColumnItem(5, Cbolydo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox("Không có dòng nào được chọn để xoá", MsgBoxStyle.Critical, "Lổi xoá.")
        End Try
    End Sub

    Private Sub cmdIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIn.Click
        Dim value
        value = MsgBox("Bạn có thật sự muốn in phiếu chi không?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Nhắc nhở")

        If (Trim$(strPrinterName) = "") Then
            MsgBox("Máy in chưa được thiết lập! Vui lòng chọn máy in rồi tiếp tục")
            Exit Sub
        End If

        If (value = vbYes) Then
            AssignVar()
            Prints()
        End If
    End Sub

    Private Sub cmdLuu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLuu.Click
        Dim value
        If (ListViewDetail.Items.Count < 1) Then
            MsgBox("Không có dữ liệu để ghi.")
            SaveFlag = False
            Exit Sub
        End If

        'If (SaveFlag) Then
        '    value = MsgBox("Danh sách này đã được ghi vào hệ thống! Bạn có muốn ghi tiếp không?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Cảnh báo!")
        '    If (value = vbYes) Then
        '        SaveExpenses()
        '    End If
        '    Exit Sub
        'End If

        value = MsgBox("Bạn có thật sự muốn ghi danh sách vào hệ thống không?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Nhắc nhở")
        If (value = vbYes) Then
            SaveExpenses()
        End If
    End Sub


    Private Sub cmddong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddong.Click
        Me.Close()
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

    Private Sub Cbolydo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbolydo.SelectedIndexChanged
        If (start) Then
            Try
                Dim str As String
                str = Cbolydo.Text
                txtlydo.Text = "Chi nộp " & Cbolydo.SelectedValue
                If (str = "KHAC") Then
                    txtlydo.ReadOnly = False
                Else
                    txtlydo.ReadOnly = True
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub CboEmploy_code_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboEmploy_code.SelectedIndexChanged
        If (start) Then
            Try
                txtEmployeeName.Text = CboEmploy_code.SelectedValue
            Catch ex As Exception
            End Try
        End If
    End Sub


    Private Sub ListViewDetail_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewDetail.SelectedIndexChanged
        Dim lvitem As ListViewItem
        Try
            lvitem = ListViewDetail.FocusedItem
        Catch ex As Exception
            MsgBox("The Item selected Invalid", MsgBoxStyle.Critical, "Delete Error.")
        End Try
    End Sub

    Private Sub DateTimePickerngay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerngay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            CboEmploy_code.Focus()
        End If
    End Sub

    Private Sub CboEmploy_code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CboEmploy_code.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (CboEmploy_code.FindString(CboEmploy_code.Text) > -1) Then
                CboEmploy_code.SelectedIndex = CboEmploy_code.FindString(CboEmploy_code.Text)
                txtEmployeeName.Text = CboEmploy_code.SelectedValue
            Else
                MsgBox("Không tìm thấy Mã CTV tương ứng", MsgBoxStyle.Critical, "Nhập sai")
                txtEmployeeName.Text = ""
                CboEmploy_code.Focus()
                Exit Sub
            End If
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub Cbolydo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Cbolydo.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub DateTimePickerchuky_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerchuky.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtchitietbk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtchitietbk.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtEmployeeName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEmployeeName.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtlydo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtlydo.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtquyen_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtquyen.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtso_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtso.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (Trim$(txtso.Text) = "") Then
                MsgBox("Số thứ tự chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtso.Focus()
                Exit Sub
            End If

            If (Not IsNumeric(Trim$(txtso.Text))) Then
                MsgBox("Số thứ tụ phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtso.Focus()
                txtso.SelectAll()
                Exit Sub
            End If

            If (CLng(Trim$(txtso.Text)) < 1) Then
                MsgBox("Số thứ tụ phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtso.Focus()
                txtso.SelectAll()
                Exit Sub
            End If
            Computing_Vol_Ord()
            DateTimePickerchuky.Focus()
            If (Not cmdLuu.Enabled) Then
                Dim lvitem As ListViewItem
                lvitem = ListViewDetail.Items(0)
                lvitem.SubItems(7).Text = txtquyen.Text
                lvitem.SubItems(8).Text = txtso.Text
            End If

        End If
    End Sub

    Private Sub txtsoBK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsoBK.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtsotien_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsotien.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (Trim$(txtsotien.Text) = "") Then
                MsgBox("Số tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsotien.Focus()
                Exit Sub
            End If

            If (Not IsNumeric(Trim$(txtsotien.Text))) Then
                MsgBox("Số tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsotien.Focus()
                txtsotien.SelectAll()
                Exit Sub
            End If
            splitn.strnumbers = CStr(CLng(txtsotien.Text))
            txtsotien.Text = splitn.Splitnumer(",")
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txttongtien_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttongtien.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtTongtienDv_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSoBKNT.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub cbostations_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbostations.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerngay.Focus()
        End If
    End Sub

    Private Sub DateTimePickerngay_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePickerngay.ValueChanged
        If DateTimePickerngay.Value > Now Then
            DateTimePickerngay.Value = Now
        End If
    End Sub

    Private Sub DateTimePickerchuky_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePickerchuky.ValueChanged
        If DateTimePickerchuky.Value > Now Then
            DateTimePickerchuky.Value = Now
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

    Private Function checkInfo() As Boolean
        Dim result As Boolean = True

        If (Trim$(CboEmploy_code.Text) = "") Then
            MsgBox("Tên người nộp chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            CboEmploy_code.Focus()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txtso.Text) = "") Then
            MsgBox("Số thứ tự chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtso.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtso.Text))) Then
            MsgBox("Số thứ tụ phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtso.Focus()
            txtso.SelectAll()
            result = False
            GoTo endFunction
        End If


        If (CLng(Trim$(txtso.Text)) < 1) Then
            MsgBox("Số thứ tụ phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtso.Focus()
            txtso.SelectAll()
            Exit Function
        End If

        If (Trim$(txtquyen.Text) = "") Then
            MsgBox("Số quyển chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtquyen.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtquyen.Text))) Then
            MsgBox("Số thứ tự phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtquyen.Focus()
            txtquyen.SelectAll()
            result = False
            GoTo endFunction
        End If


        If (CLng(Trim$(txtquyen.Text)) < 1) Then
            MsgBox("Số quyển phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtquyen.Focus()
            txtquyen.SelectAll()
            Exit Function
        End If

        If (Trim$(txtsotien.Text) = "") Then
            MsgBox("Số tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsotien.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtsotien.Text))) Then
            MsgBox("Số tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsotien.Focus()
            txtsotien.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txtsoBK.Text) = "") Then
            MsgBox("Số lượng chứng từ chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsoBK.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtsoBK.Text))) Then
            MsgBox("Số bảng kê phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsoBK.Focus()
            txtsoBK.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txtSoBKNT.Text) = "") Then
            MsgBox("Số bảng kê nộp tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtSoBKNT.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtSoBKNT.Text))) Then
            MsgBox("Số bảng kê nộp tiền là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtSoBKNT.Focus()
            txtSoBKNT.SelectAll()
            result = False
            GoTo endFunction
        End If


        If (CLng(Trim$(txtSoBKNT.Text)) < 1) Then
            MsgBox("Số bảng kê nộp tiền phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtSoBKNT.Focus()
            txtSoBKNT.SelectAll()
            Exit Function
        End If

endFunction:
        Return result
    End Function

    Private Sub GetNumber_Vol_Ord()
        Dim strQuery As String
        strQuery = " SELECT MAX(Ordinal_No)+1  FROM Tbl_Expenses WHERE ( MONTH(Expense_Date) = " & DateTimePickerngay.Value.Month & ") AND  (YEAR(Expense_Date) = " & DateTimePickerngay.Value.Year() & ")"
        txtso.Text = GetMaxNumber(strQuery)
    End Sub
    Private Sub Computing_Vol_Ord()

        Dim so As Long
        so = (CLng(txtso.Text) Mod 50)
        If (so = 0) Then
            txtquyen.Text = CLng(txtso.Text) \ 50
        Else
            txtquyen.Text = (CLng(txtso.Text) \ 50) + 1
        End If

    End Sub

    Private Sub DeleteTextBox()
        txtsotien.Text = vbNullString
        txtsoBK.Text = vbNullString
        txtchitietbk.Text = vbNullString
    End Sub

    Private Function SumColumnItem(ByVal index As Integer) As String
        Dim result As Long = 0
        Dim strvalue As String
        Dim i As Integer
        For i = 0 To ListViewDetail.Items.Count - 1
            result += CLng(ListViewDetail.Items(i).SubItems(index).Text)
        Next
        splitn.strnumbers = CStr(result)
        strvalue = splitn.Splitnumer(",")
        Return strvalue
    End Function

    Private Function SumColumnItem(ByVal index As Integer, ByVal strSer As String) As String
        Dim result As Long = 0
        Dim strvalue As String
        Dim i As Integer
        For i = 0 To ListViewDetail.Items.Count - 1
            If (ListViewDetail.Items(i).SubItems(1).Text = strSer) Then
                result += CLng(ListViewDetail.Items(i).SubItems(index).Text)
            End If
        Next
        splitn.strnumbers = CStr(result)
        strvalue = splitn.Splitnumer(",")
        Return strvalue
    End Function

    Private Sub Cbolydo_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cbolydo.Validated
        'lbltongDV.Text = "TT DV " & Cbolydo.Text
        'txtTongtienDv.Text = SumColumnItem(5, Cbolydo.Text)
    End Sub
    Private Sub SaveExpenses()
        Dim lvitem As ListViewItem
        Dim i As Integer

        Try
            For i = 0 To ListViewDetail.Items.Count - 1
                lvitem = ListViewDetail.Items(i)
                SaveExpense(lvitem) ' Luu vao bang phieu chi
                SaveToExpenses(lvitem) ' Luu vao bang so quy
                UpdateReceipts()
            Next
            MsgBox("Đã lưu vào hệ thống!!")
            cmdLuu.Enabled = False
            'SaveFlag = True
            'SaveFlag = False
            'Indexlistview = 0
            'ListViewDetail.Items.Clear()
            'DeleteTextBox()
            txttongtien.Text = "0"
        Catch ex As Exception
            MsgBox("Không có dòng nào để ghi", MsgBoxStyle.Critical, "Lổi xoá.")
        End Try
    End Sub

    Public Sub UpdateReceipts()
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        Dim valueID As Long
        Dim strQuery As String

        strSQL = " SELECT MAX(ID)  FROM Tbl_Expenses  "
        MaxIDExpense = GetMaxNumber(strSQL)

        Try
            For i = 0 To ArrListReceiID.Length - 1
                strQuery = " UPDATE Tbl_Receipts SET Status = True, Ordinal_Expen_No = " & MaxIDExpense & " WHERE ID = " & CLng(ArrListReceiID(i))
                UpdateReceipt(strQuery)
            Next
        Catch ex As Exception
        End Try

    End Sub

    Public Sub UpdateReceipt(ByVal strQuery As String)
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strQuery
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        oledbcon.Close()
    End Sub

    Private Sub cmddeletes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddeletes.Click
        GetNumber_Vol_Ord()
        Computing_Vol_Ord()
        cmdLuu.Enabled = True
        SaveFlag = False
        Indexlistview = 0
        ListViewDetail.Items.Clear()
        DeleteTextBox()
        txttongtien.Text = "0"
        'txtTongtienDv.Text = "0"
    End Sub
    Private Sub AssignVar()
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
        Arrvar(13) = "txttenNguoiNop1"

        ' gan gia tri cua bien vao
        Arrval(0) = "Đơn vị: " & cbostations.Text
        Arrval(1) = "Địa chỉ: " & GetStringName(cbostations.SelectedValue)
        Arrval(2) = DateTimePickerngay.Value.Day
        Arrval(3) = DateTimePickerngay.Value.Month
        Arrval(4) = DateTimePickerngay.Value.Year

        Arrval(7) = txtEmployeeName.Text
        Arrval(8) = CboEmploy_code.Text

    End Sub
    Private Sub AssignVal(ByVal listitem As ListViewItem)

        Arrval(5) = listitem.SubItems(8).Text ' So thu tu
        Arrval(6) = listitem.SubItems(7).Text ' Quyen so
        Arrval(9) = listitem.SubItems(6).Text & "Kỳ cước " & listitem.SubItems(2).Text & ". "  ' Ly do nop
        Arrval(10) = listitem.SubItems(5).Text ' So tien 
        numbers.Number = CLng(Arrval(10))
        Dim str As String
        str = numbers.NumbersToString
        If (Len(str) > 0) Then
            str = str.Substring(0, 1).ToUpper + str.Remove(0, 1)
        Else
            str = "Không"
        End If
        Arrval(11) = str + " đồng." ' tien bang chu
        Arrval(12) = listitem.SubItems(3).Text & "  chứng từ : " & listitem.SubItems(4).Text & "."  ' Chung tu goc 1
        Arrval(13) = "" 'listitem.SubItems(5).Text & "  hóa đơn." ' Chung tu goc 2
        Arrval(14) = txtEmployeeName.Text
        SetFieldTextOjectReports(rpt, Arrvar, Arrval)

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

    Private Function GetStringName(ByVal strCode) As String
        Dim i As Integer
        Dim strresult As String
        For i = 0 To mydataset.Tables(0).Rows.Count - 1
            strresult = mydataset.Tables(0).Rows(i).Item(0)
            If (strresult.Equals(strCode)) Then
                strresult = mydataset.Tables(0).Rows(i).Item(2)
                Exit For
            End If
        Next
        Return strresult
    End Function

    Private Sub Prints()
        Dim i As Integer = 0
        Dim listitem As ListViewItem
        Try
            For i = 0 To ListViewDetail.Items.Count - 1
                listitem = ListViewDetail.Items(i)
                AssignVal(listitem)
                rpt.PrintToPrinter(NumericUpDownSobanin.Value, True, 1, 1)
            Next
            MsgBox("Đã in xong : " & i & " phiếu chi.")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click
        If (ListViewDetail.Items.Count < 1) Then
            MsgBox("Không có phiếu chi để xem!")
            Exit Sub
        End If
        rpt = New CrystalReport_Expense
        AssignVar()
        Dim frm As New frmPreview
        Dim listitem As ListViewItem
        listitem = ListViewDetail.Items(0)
        AssignVal(listitem)
        frm.CrystalReportViewerReceipts.ReportSource = rpt
        frm.ShowDialog()
    End Sub

    Private Sub cmdSetprinting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PrintDialog1.Document = ThePrintDocument
        PrintDialog1.ShowDialog()
        rpt.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName
    End Sub
    'luu vao so quy
    Private Sub SaveToExpenses(ByVal lvitem As ListViewItem)

        strSQL = " INSERT INTO Tbl_Receipts_Expenses(Recei_Expen_Date,Expen_No,Descriptions,Expen_Money) " & _
        " VALUES('" & DateTimePickerngay.Value.ToShortDateString & _
        "'," & CInt(Trim$(lvitem.SubItems(8).Text)) & _
        ",'" & lvitem.SubItems(6).Text & " Kỳ cước " & lvitem.SubItems(2).Text & _
        "'," & CLng(lvitem.SubItems(5).Text) & ")"
        ExcuxeSQL(strSQL)

    End Sub

    Private Sub txtso_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtso.Enter
        GetNumber_Vol_Ord()
        Computing_Vol_Ord()
    End Sub

End Class
