Imports ConvertNumberToChar
Imports System.Data.OleDb
Public Class frmPhieuThu
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private Indexlistview As Integer
    Dim start As Boolean = False
    Dim splitn As New SplitNumbers
    Dim numbers As New ConvertNumbersToString
    Dim rpt As CrystalReport_Receipts
    Dim SaveFlag As Boolean = False
    Dim Arrvar(15) As String
    Dim Arrval(15) As String
    Friend WithEvents ThePrintDocument As System.Drawing.Printing.PrintDocument
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtso As System.Windows.Forms.TextBox
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents txtsoBK As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePickerchuky As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdThemPT As System.Windows.Forms.Button
    Friend WithEvents cmdbotPT As System.Windows.Forms.Button
    Friend WithEvents cmdIn As System.Windows.Forms.Button
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents cmdLuu As System.Windows.Forms.Button
    Friend WithEvents txtquyen As System.Windows.Forms.TextBox
    Friend WithEvents txtsotien As System.Windows.Forms.TextBox
    Friend WithEvents txtchitietbk As System.Windows.Forms.TextBox
    Friend WithEvents txtsohd As System.Windows.Forms.TextBox
    Friend WithEvents txtTongtienDv As System.Windows.Forms.TextBox
    Friend WithEvents txttongtien As System.Windows.Forms.TextBox
    Friend WithEvents ListViewDetail As System.Windows.Forms.ListView
    Friend WithEvents DateTimePickerngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtlydo As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents lbltongDV As System.Windows.Forms.Label
    Friend WithEvents cmddeletes As System.Windows.Forms.Button
    Friend WithEvents cmdPreview As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDownSobanin As System.Windows.Forms.NumericUpDown
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cmbHTthu As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cboAccounts As System.Windows.Forms.ComboBox
    Friend WithEvents txtbankname As System.Windows.Forms.TextBox
    Friend WithEvents txtsoGNT As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePickerngaynop As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtslunc As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPhieuThu))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtslunc = New System.Windows.Forms.TextBox()
        Me.NumericUpDownSobanin = New System.Windows.Forms.NumericUpDown()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtso = New System.Windows.Forms.TextBox()
        Me.Cbolydo = New System.Windows.Forms.ComboBox()
        Me.txtsoBK = New System.Windows.Forms.TextBox()
        Me.DateTimePickerchuky = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtquyen = New System.Windows.Forms.TextBox()
        Me.txtsotien = New System.Windows.Forms.TextBox()
        Me.txtchitietbk = New System.Windows.Forms.TextBox()
        Me.txtsohd = New System.Windows.Forms.TextBox()
        Me.txtTongtienDv = New System.Windows.Forms.TextBox()
        Me.txttongtien = New System.Windows.Forms.TextBox()
        Me.ListViewDetail = New System.Windows.Forms.ListView()
        Me.DateTimePickerngay = New System.Windows.Forms.DateTimePicker()
        Me.txtEmployeeName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtlydo = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lbltongDV = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtsoGNT = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.cboAccounts = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtbankname = New System.Windows.Forms.TextBox()
        Me.DateTimePickerngaynop = New System.Windows.Forms.DateTimePicker()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cmdPreview = New System.Windows.Forms.Button()
        Me.cmdThemPT = New System.Windows.Forms.Button()
        Me.cmdbotPT = New System.Windows.Forms.Button()
        Me.cmdIn = New System.Windows.Forms.Button()
        Me.cmddong = New System.Windows.Forms.Button()
        Me.cmdLuu = New System.Windows.Forms.Button()
        Me.cmddeletes = New System.Windows.Forms.Button()
        Me.cbostations = New System.Windows.Forms.ComboBox()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.cmbHTthu = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDownSobanin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.txtslunc)
        Me.GroupBox1.Controls.Add(Me.NumericUpDownSobanin)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.txtso)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.txtsoBK)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerchuky)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtquyen)
        Me.GroupBox1.Controls.Add(Me.txtsotien)
        Me.GroupBox1.Controls.Add(Me.txtchitietbk)
        Me.GroupBox1.Controls.Add(Me.txtsohd)
        Me.GroupBox1.Controls.Add(Me.txtTongtienDv)
        Me.GroupBox1.Controls.Add(Me.txttongtien)
        Me.GroupBox1.Controls.Add(Me.ListViewDetail)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerngay)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtlydo)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.lbltongDV)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(5, 39)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 432)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(586, 148)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(64, 19)
        Me.Label19.TabIndex = 55
        Me.Label19.Text = "SL UNC"
        '
        'txtslunc
        '
        Me.txtslunc.BackColor = System.Drawing.Color.White
        Me.txtslunc.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtslunc.ForeColor = System.Drawing.Color.Black
        Me.txtslunc.Location = New System.Drawing.Point(651, 144)
        Me.txtslunc.Name = "txtslunc"
        Me.txtslunc.Size = New System.Drawing.Size(46, 26)
        Me.txtslunc.TabIndex = 54
        Me.txtslunc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'NumericUpDownSobanin
        '
        Me.NumericUpDownSobanin.Location = New System.Drawing.Point(654, 257)
        Me.NumericUpDownSobanin.Name = "NumericUpDownSobanin"
        Me.NumericUpDownSobanin.Size = New System.Drawing.Size(48, 26)
        Me.NumericUpDownSobanin.TabIndex = 51
        Me.NumericUpDownSobanin.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(584, 259)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(70, 19)
        Me.Label9.TabIndex = 52
        Me.Label9.Text = "Số bản in "
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Location = New System.Drawing.Point(320, 17)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(168, 27)
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
        Me.txtso.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtso.ForeColor = System.Drawing.Color.Black
        Me.txtso.Location = New System.Drawing.Point(71, 86)
        Me.txtso.Name = "txtso"
        Me.txtso.Size = New System.Drawing.Size(137, 26)
        Me.txtso.TabIndex = 4
        Me.txtso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(71, 56)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(137, 27)
        Me.Cbolydo.TabIndex = 3
        '
        'txtsoBK
        '
        Me.txtsoBK.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsoBK.ForeColor = System.Drawing.Color.Black
        Me.txtsoBK.Location = New System.Drawing.Point(71, 144)
        Me.txtsoBK.Name = "txtsoBK"
        Me.txtsoBK.Size = New System.Drawing.Size(137, 26)
        Me.txtsoBK.TabIndex = 8
        Me.txtsoBK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DateTimePickerchuky
        '
        Me.DateTimePickerchuky.CustomFormat = "MM/yyyy"
        Me.DateTimePickerchuky.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerchuky.Location = New System.Drawing.Point(71, 115)
        Me.DateTimePickerchuky.Name = "DateTimePickerchuky"
        Me.DateTimePickerchuky.Size = New System.Drawing.Size(137, 26)
        Me.DateTimePickerchuky.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(216, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 19)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Quyển số"
        '
        'txtquyen
        '
        Me.txtquyen.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtquyen.ForeColor = System.Drawing.Color.Black
        Me.txtquyen.Location = New System.Drawing.Point(292, 86)
        Me.txtquyen.Name = "txtquyen"
        Me.txtquyen.Size = New System.Drawing.Size(144, 26)
        Me.txtquyen.TabIndex = 5
        Me.txtquyen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtsotien
        '
        Me.txtsotien.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsotien.ForeColor = System.Drawing.Color.Black
        Me.txtsotien.Location = New System.Drawing.Point(292, 115)
        Me.txtsotien.Name = "txtsotien"
        Me.txtsotien.Size = New System.Drawing.Size(144, 26)
        Me.txtsotien.TabIndex = 7
        Me.txtsotien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtchitietbk
        '
        Me.txtchitietbk.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtchitietbk.ForeColor = System.Drawing.Color.Black
        Me.txtchitietbk.Location = New System.Drawing.Point(292, 144)
        Me.txtchitietbk.Name = "txtchitietbk"
        Me.txtchitietbk.Size = New System.Drawing.Size(144, 26)
        Me.txtchitietbk.TabIndex = 9
        '
        'txtsohd
        '
        Me.txtsohd.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsohd.ForeColor = System.Drawing.Color.Black
        Me.txtsohd.Location = New System.Drawing.Point(528, 144)
        Me.txtsohd.Name = "txtsohd"
        Me.txtsohd.Size = New System.Drawing.Size(56, 26)
        Me.txtsohd.TabIndex = 10
        Me.txtsohd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTongtienDv
        '
        Me.txtTongtienDv.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongtienDv.ForeColor = System.Drawing.Color.Brown
        Me.txtTongtienDv.Location = New System.Drawing.Point(528, 86)
        Me.txtTongtienDv.Name = "txtTongtienDv"
        Me.txtTongtienDv.Size = New System.Drawing.Size(168, 26)
        Me.txtTongtienDv.TabIndex = 50
        Me.txtTongtienDv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txttongtien
        '
        Me.txttongtien.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttongtien.ForeColor = System.Drawing.Color.Brown
        Me.txttongtien.Location = New System.Drawing.Point(528, 115)
        Me.txttongtien.Name = "txttongtien"
        Me.txttongtien.Size = New System.Drawing.Size(168, 26)
        Me.txttongtien.TabIndex = 50
        Me.txttongtien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ListViewDetail
        '
        Me.ListViewDetail.FullRowSelect = True
        Me.ListViewDetail.GridLines = True
        Me.ListViewDetail.Location = New System.Drawing.Point(8, 288)
        Me.ListViewDetail.Name = "ListViewDetail"
        Me.ListViewDetail.Size = New System.Drawing.Size(696, 136)
        Me.ListViewDetail.TabIndex = 28
        Me.ListViewDetail.UseCompatibleStateImageBehavior = False
        Me.ListViewDetail.View = System.Windows.Forms.View.Details
        '
        'DateTimePickerngay
        '
        Me.DateTimePickerngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngay.Location = New System.Drawing.Point(71, 17)
        Me.DateTimePickerngay.Name = "DateTimePickerngay"
        Me.DateTimePickerngay.Size = New System.Drawing.Size(137, 26)
        Me.DateTimePickerngay.TabIndex = 1
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(488, 17)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(216, 26)
        Me.txtEmployeeName.TabIndex = 50
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 19)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Số PT"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 19)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Ngày/T/N"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(220, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 19)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Mã/Tên NVT"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 58)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 19)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Dịch vụ"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(216, 117)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 19)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Số tiền"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 146)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(52, 19)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "SL BK"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(439, 146)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(83, 19)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "SL Hoá đơn"
        '
        'txtlydo
        '
        Me.txtlydo.BackColor = System.Drawing.Color.White
        Me.txtlydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtlydo.ForeColor = System.Drawing.Color.Black
        Me.txtlydo.Location = New System.Drawing.Point(208, 56)
        Me.txtlydo.Name = "txtlydo"
        Me.txtlydo.Size = New System.Drawing.Size(496, 26)
        Me.txtlydo.TabIndex = 50
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(216, 146)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(78, 19)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "Chi tiết BK"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(8, 117)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(53, 19)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Chu kỳ"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType(((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic) _
                        Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Brown
        Me.Label12.Location = New System.Drawing.Point(9, 256)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(238, 20)
        Me.Label12.TabIndex = 25
        Me.Label12.Text = "Danh sách chi tiết các phiếu thu"
        '
        'lbltongDV
        '
        Me.lbltongDV.AutoSize = True
        Me.lbltongDV.Location = New System.Drawing.Point(439, 88)
        Me.lbltongDV.Name = "lbltongDV"
        Me.lbltongDV.Size = New System.Drawing.Size(91, 19)
        Me.lbltongDV.TabIndex = 25
        Me.lbltongDV.Text = "Tổng tiền DV"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(439, 117)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(65, 19)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Tổng tiền"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtsoGNT)
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.Controls.Add(Me.cboAccounts)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.txtbankname)
        Me.GroupBox3.Controls.Add(Me.DateTimePickerngaynop)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 167)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(696, 48)
        Me.GroupBox3.TabIndex = 53
        Me.GroupBox3.TabStop = False
        '
        'txtsoGNT
        '
        Me.txtsoGNT.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsoGNT.ForeColor = System.Drawing.Color.Black
        Me.txtsoGNT.Location = New System.Drawing.Point(64, 14)
        Me.txtsoGNT.Name = "txtsoGNT"
        Me.txtsoGNT.Size = New System.Drawing.Size(56, 26)
        Me.txtsoGNT.TabIndex = 54
        Me.txtsoGNT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(2, 18)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(62, 19)
        Me.Label17.TabIndex = 55
        Me.Label17.Text = "GNT Số"
        '
        'cboAccounts
        '
        Me.cboAccounts.Location = New System.Drawing.Point(336, 14)
        Me.cboAccounts.Name = "cboAccounts"
        Me.cboAccounts.Size = New System.Drawing.Size(160, 27)
        Me.cboAccounts.TabIndex = 51
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(286, 17)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(51, 19)
        Me.Label16.TabIndex = 52
        Me.Label16.Text = "Số TK"
        '
        'txtbankname
        '
        Me.txtbankname.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbankname.ForeColor = System.Drawing.Color.Black
        Me.txtbankname.Location = New System.Drawing.Point(496, 14)
        Me.txtbankname.Name = "txtbankname"
        Me.txtbankname.Size = New System.Drawing.Size(192, 26)
        Me.txtbankname.TabIndex = 54
        Me.txtbankname.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DateTimePickerngaynop
        '
        Me.DateTimePickerngaynop.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngaynop.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngaynop.Location = New System.Drawing.Point(188, 14)
        Me.DateTimePickerngaynop.Name = "DateTimePickerngaynop"
        Me.DateTimePickerngaynop.Size = New System.Drawing.Size(95, 26)
        Me.DateTimePickerngaynop.TabIndex = 1
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(120, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(69, 19)
        Me.Label18.TabIndex = 25
        Me.Label18.Text = "Ngày nộp"
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
        Me.GroupBox2.Location = New System.Drawing.Point(8, 208)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(696, 48)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        '
        'cmdPreview
        '
        Me.cmdPreview.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPreview.Location = New System.Drawing.Point(434, 14)
        Me.cmdPreview.Name = "cmdPreview"
        Me.cmdPreview.Size = New System.Drawing.Size(80, 27)
        Me.cmdPreview.TabIndex = 52
        Me.cmdPreview.Text = "Xem"
        '
        'cmdThemPT
        '
        Me.cmdThemPT.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdThemPT.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdThemPT.Location = New System.Drawing.Point(4, 14)
        Me.cmdThemPT.Name = "cmdThemPT"
        Me.cmdThemPT.Size = New System.Drawing.Size(80, 27)
        Me.cmdThemPT.TabIndex = 12
        Me.cmdThemPT.Text = "Thêm PT"
        '
        'cmdbotPT
        '
        Me.cmdbotPT.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdbotPT.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdbotPT.Location = New System.Drawing.Point(90, 14)
        Me.cmdbotPT.Name = "cmdbotPT"
        Me.cmdbotPT.Size = New System.Drawing.Size(80, 27)
        Me.cmdbotPT.TabIndex = 23
        Me.cmdbotPT.Text = "Bớt PT"
        '
        'cmdIn
        '
        Me.cmdIn.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdIn.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdIn.Location = New System.Drawing.Point(176, 14)
        Me.cmdIn.Name = "cmdIn"
        Me.cmdIn.Size = New System.Drawing.Size(80, 27)
        Me.cmdIn.TabIndex = 23
        Me.cmdIn.Text = "In"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(608, 14)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(80, 27)
        Me.cmddong.TabIndex = 31
        Me.cmddong.Text = "Đóng"
        '
        'cmdLuu
        '
        Me.cmdLuu.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuu.Location = New System.Drawing.Point(262, 14)
        Me.cmdLuu.Name = "cmdLuu"
        Me.cmdLuu.Size = New System.Drawing.Size(80, 27)
        Me.cmdLuu.TabIndex = 23
        Me.cmdLuu.Text = "Lưu"
        '
        'cmddeletes
        '
        Me.cmddeletes.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddeletes.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddeletes.Location = New System.Drawing.Point(348, 14)
        Me.cmddeletes.Name = "cmddeletes"
        Me.cmddeletes.Size = New System.Drawing.Size(80, 27)
        Me.cmddeletes.TabIndex = 23
        Me.cmddeletes.Text = "Xóa DS"
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(503, 11)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(216, 27)
        Me.cbostations.TabIndex = 50
        '
        'cmbHTthu
        '
        Me.cmbHTthu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHTthu.Location = New System.Drawing.Point(327, 11)
        Me.cmbHTthu.Name = "cmbHTthu"
        Me.cmbHTthu.Size = New System.Drawing.Size(176, 27)
        Me.cmbHTthu.TabIndex = 51
        '
        'Label15
        '
        Me.Label15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label15.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label15.Font = New System.Drawing.Font("Arial", 21.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label15.Location = New System.Drawing.Point(5, 2)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(712, 38)
        Me.Label15.TabIndex = 52
        Me.Label15.Text = "NHẬP PHIẾU THU"
        '
        'frmPhieuThu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(722, 480)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.cmbHTthu)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label15)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmPhieuThu"
        Me.Text = "Nhập Phiếu Thu "
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.NumericUpDownSobanin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub FillDataSet()

        strSQL = "SELECT MaLoaithu,TenLoaiThu FROM Tbl_LoaiThu"
        FillCombo(cmbHTthu, strSQL, "Tbl_LoaiThu", "TenLoaiThu", "MaLoaithu")

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
    Private Sub frmPhieuThu_Expenses_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePickerngay.Value = Now
        DateTimePickerchuky.Value = Now
        SetDetailListView()
        Try
            mydataset = New DataSet
            FillDataSet()
        Catch eLoad As System.Exception
            System.Windows.Forms.MessageBox.Show(eLoad.Message)
        End Try

        strSQL = "SELECT Account_No,Bank_Name FROM Tbl_Accounts_Banks,Tbl_Banks WHERE Tbl_Banks.Bank_Code = Tbl_Accounts_Banks.Bank_Code"
        FillCombo(cboAccounts, strSQL, "Tbl_Accounts_Banks", "Account_No", "Bank_Name")
        cboAccounts.Text = vbNullString

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
            txtlydo.Text = "Thu " & Cbolydo.SelectedValue
        End If

        If (cmbHTthu.Items.Count > 0) Then
            If (cmbHTthu.FindString("TIỀN MẶT") > 0) Then
                cmbHTthu.SelectedIndex = cmbHTthu.FindString("TIỀN MẶT")
            End If
        End If

        GroupBox3.Enabled = False
        txtslunc.Visible = False
        Label19.Visible = False
        txtEmployeeName.ReadOnly = True
        rpt = New CrystalReport_Receipts
    End Sub
    Private Sub GetNumber_Vol_Ord()
        Dim strQuery As String
        strQuery = " SELECT MAX(Ordinal_No)+1  FROM Tbl_Receipts WHERE  MaLoaiThu = '" & cmbHTthu.SelectedValue & "' AND ( MONTH(Receipt_Date) = " & DateTimePickerngay.Value.Month & ") AND  (YEAR(Receipt_Date) = " & DateTimePickerngay.Value.Year() & ")"
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

    Private Sub SetDetailListView()
        ListViewDetail.Items.Clear()
        ListViewDetail.Columns.Clear()
        ListViewDetail.Columns.Add("STT", 50, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Dịch vụ", 80, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Chu kỳ", 70, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("SL BK", 60, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Chi tiết BK", 140, HorizontalAlignment.Left)
        ListViewDetail.Columns.Add("SL HĐ", 60, HorizontalAlignment.Right)
        ListViewDetail.Columns.Add("Tiền", 80, HorizontalAlignment.Right)
        ListViewDetail.Columns.Add("Mô tả", 180, HorizontalAlignment.Left)
        ListViewDetail.Columns.Add("Q Số", 50, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Số", 50, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Loại thu", 80, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Số GNT", 80, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Ngày nộp", 80, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("Tài khoản", 80, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("SL UNC", 80, HorizontalAlignment.Center)


    End Sub

    Private Sub AddSubItemListView()

        ListViewDetail.Items.Add(ListViewDetail.Items.Count + 1)
        ListViewDetail.Items(Indexlistview).SubItems.Add(Cbolydo.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(DateTimePickerchuky.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtsoBK.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtchitietbk.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtsohd.Text)
        splitn.strnumbers = CStr(CLng(txtsotien.Text))
        ListViewDetail.Items(Indexlistview).SubItems.Add(splitn.Splitnumer(","))
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtlydo.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtquyen.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(txtso.Text)
        ListViewDetail.Items(Indexlistview).SubItems.Add(CStr(cmbHTthu.SelectedValue))


        If (cmbHTthu.SelectedValue = "GNT") Then
            ListViewDetail.Items(Indexlistview).SubItems.Add(txtsoGNT.Text)
            ListViewDetail.Items(Indexlistview).SubItems.Add(DateTimePickerngaynop.Value.ToShortDateString)
            ListViewDetail.Items(Indexlistview).SubItems.Add(cboAccounts.Text)
            ListViewDetail.Items(Indexlistview).SubItems.Add("") 'sounc
        Else
            If (cmbHTthu.SelectedValue = "UNC") Then
                ListViewDetail.Items(Indexlistview).SubItems.Add("")
                ListViewDetail.Items(Indexlistview).SubItems.Add("")
                ListViewDetail.Items(Indexlistview).SubItems.Add("")
                ListViewDetail.Items(Indexlistview).SubItems.Add(txtslunc.Text)
            Else
                ListViewDetail.Items(Indexlistview).SubItems.Add("")
                ListViewDetail.Items(Indexlistview).SubItems.Add("")
                ListViewDetail.Items(Indexlistview).SubItems.Add("")
                ListViewDetail.Items(Indexlistview).SubItems.Add("")
            End If
        End If

        Indexlistview += 1
    End Sub
    Private Sub SaveReceipt(ByVal lvitem As ListViewItem)
        Dim strLoaithu As String
        strLoaithu = lvitem.SubItems(10).Text

        Select Case strLoaithu
            Case "TM"
                strSQL = " INSERT INTO Tbl_Receipts(Volume,Ordinal_No,Receipt_Date,Service_Code,Descriptions,List_Quantity,List_Detail,Invoice_Quantity,Charge_Cycle,Total_Money,Employ_Code,MaLoaiThu) " & _
                " VALUES(" & CInt(Trim$(lvitem.SubItems(8).Text)) & _
                "," & CInt(Trim$(lvitem.SubItems(9).Text)) & _
                ",'" & DateTimePickerngay.Value.ToShortDateString & _
                "','" & lvitem.SubItems(1).Text & _
                "','" & lvitem.SubItems(7).Text & _
                "'," & CInt(Trim$(lvitem.SubItems(3).Text)) & _
                ",'" & lvitem.SubItems(4).Text & _
                "'," & CInt(Trim$(lvitem.SubItems(5).Text)) & _
                ",'" & lvitem.SubItems(2).Text & _
                "'," & CLng(lvitem.SubItems(6).Text) & _
                ",'" & CboEmploy_code.Text & _
                "','" & lvitem.SubItems(10).Text & "')"
                ExcuxeSQL(strSQL)
                SaveToExpenses(lvitem)   'luu vao so quy

            Case "UNC"
                strSQL = " INSERT INTO Tbl_Receipts(Volume,Ordinal_No,Receipt_Date,Service_Code,Descriptions,List_Quantity,List_Detail,Invoice_Quantity,Charge_Cycle,Total_Money,Employ_Code,MaLoaiThu,SLUNC) " & _
                                    " VALUES(" & CInt(Trim$(lvitem.SubItems(8).Text)) & _
                                    "," & CInt(Trim$(lvitem.SubItems(9).Text)) & _
                                    ",'" & DateTimePickerngay.Value.ToShortDateString & _
                                    "','" & lvitem.SubItems(1).Text & _
                                    "','" & lvitem.SubItems(7).Text & _
                                    "'," & CInt(Trim$(lvitem.SubItems(3).Text)) & _
                                    ",'" & lvitem.SubItems(4).Text & _
                                    "'," & CInt(Trim$(lvitem.SubItems(5).Text)) & _
                                    ",'" & lvitem.SubItems(2).Text & _
                                    "'," & CLng(lvitem.SubItems(6).Text) & _
                                    ",'" & CboEmploy_code.Text & _
                                    "','" & lvitem.SubItems(10).Text & _
                                    "'," & CInt(lvitem.SubItems(14).Text) & ")"
                ExcuxeSQL(strSQL)
            Case "GNT"
                strSQL = " INSERT INTO Tbl_Receipts(Volume,Ordinal_No,Receipt_Date,Service_Code,Descriptions,List_Quantity,List_Detail,Invoice_Quantity,Charge_Cycle,Total_Money,Employ_Code,MaLoaiThu,Pay_No,Pay_Date,Account_Code,NguoiNop) " & _
                                                " VALUES(" & CInt(Trim$(lvitem.SubItems(8).Text)) & _
                                                "," & CInt(Trim$(lvitem.SubItems(9).Text)) & _
                                                ",'" & DateTimePickerngay.Value.ToShortDateString & _
                                                "','" & lvitem.SubItems(1).Text & _
                                                "','" & lvitem.SubItems(7).Text & _
                                                "'," & CInt(Trim$(lvitem.SubItems(3).Text)) & _
                                                ",'" & lvitem.SubItems(4).Text & _
                                                "'," & CInt(Trim$(lvitem.SubItems(5).Text)) & _
                                                ",'" & lvitem.SubItems(2).Text & _
                                                "'," & CLng(lvitem.SubItems(6).Text) & _
                                                ",'" & CboEmploy_code.Text & _
                                                "','" & lvitem.SubItems(10).Text & _
                                                "'," & CInt(lvitem.SubItems(11).Text) & _
                                                ",'" & lvitem.SubItems(12).Text & _
                                                "','" & lvitem.SubItems(13).Text & _
                                                "','" & CboEmploy_code.Text & "')"
                ExcuxeSQL(strSQL)
        End Select

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
            txttongtien.Text = SumColumnItem(6)
            txtTongtienDv.Text = SumColumnItem(6, Cbolydo.Text)
            DeleteTextBox()
            txtso.Text = CLng(txtso.Text)
            Computing_Vol_Ord()
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
                    DeleteTextBox()
                    Resort()
                    txttongtien.Text = SumColumnItem(6)
                    txtTongtienDv.Text = SumColumnItem(6, Cbolydo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox("Không có dòng nào được chọn để xoá", MsgBoxStyle.Critical, "Lổi xoá.")
        End Try
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
        '        SaveReceicepts()
        '    End If
        '    Exit Sub
        'End If

        value = MsgBox("Bạn có thật sự muốn ghi danh sách vào hệ thống không?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Nhắc nhở")
        If (value = vbYes) Then
            If CheckReciepts_No() Then
                SaveReceicepts()
            End If
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
                txtlydo.Text = "Thu " & Cbolydo.SelectedValue
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
            Cbolydo.Focus()
        End If
    End Sub

    Private Sub Cbolydo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Cbolydo.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (Cbolydo.FindString(Cbolydo.Text) > -1) Then
                Cbolydo.SelectedIndex = Cbolydo.FindString(Cbolydo.Text)
            Else
                MsgBox("Không tìm thấy dịch vụ tương ứng", MsgBoxStyle.Critical, "Nhập sai")
                txtlydo.Text = ""
                Cbolydo.Focus()
                Exit Sub
            End If
            DateTimePickerchuky.Focus()
        End If
    End Sub

    Private Sub DateTimePickerchuky_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerchuky.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtsotien.Focus()
        End If
    End Sub

    Private Sub txtchitietbk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtchitietbk.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtsohd.Focus()
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
        End If
    End Sub

    Private Sub txtsoBK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsoBK.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtchitietbk.Focus()
        End If
    End Sub

    Private Sub txtsohd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsohd.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdThemPT.Focus()
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
            txtsoBK.Focus()
        End If
    End Sub

    Private Sub txttongtien_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttongtien.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtTongtienDv_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTongtienDv.KeyPress
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
            MsgBox("Tên CTV chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
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
            MsgBox("Số lượng bảng kê chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
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


        If (Trim$(txtsohd.Text) = "") Then
            MsgBox("Số hóa đơn chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsohd.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txtsohd.Text))) Then
            MsgBox("Số hóa đơn phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtsohd.Focus()
            txtsohd.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (cmbHTthu.SelectedValue = "GNT") Then

            If (Trim$(txtsoGNT.Text) = "") Then
                MsgBox("Số giấy nộp tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsoGNT.Focus()
                result = False
                GoTo endFunction
            End If

            If (Not IsNumeric(Trim$(txtsoGNT.Text))) Then
                MsgBox("Số giấy nộp tiền phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsoGNT.Focus()
                txtsoGNT.SelectAll()
                result = False
                GoTo endFunction
            End If

            If (CLng(Trim$(txtsoGNT.Text)) < 1) Then
                MsgBox("Số giấy nộp tiền phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsoGNT.Focus()
                txtsoGNT.SelectAll()
                Exit Function
            End If

            If (Trim$(cboAccounts.Text) = "") Then
                MsgBox("Số tài khoản chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                cboAccounts.Focus()
                result = False
                GoTo endFunction
            End If

        End If


        If (cmbHTthu.SelectedValue = "UNC") Then

            If (Trim$(txtslunc.Text) = "") Then
                MsgBox("Số lượng UNC chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtslunc.Focus()
                result = False
                GoTo endFunction
            End If

            If (Not IsNumeric(Trim$(txtslunc.Text))) Then
                MsgBox("Số lượng UNC phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtslunc.Focus()
                txtslunc.SelectAll()
                result = False
                GoTo endFunction
            End If


            If (CLng(Trim$(txtslunc.Text)) < 1) Then
                MsgBox("Số lượng UNC phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtslunc.Focus()
                txtslunc.SelectAll()
                Exit Function
            End If
        End If

endFunction:
        Return result
    End Function

    Private Sub DeleteTextBox()
        txtsotien.Text = vbNullString
        txtsoBK.Text = vbNullString
        txtchitietbk.Text = vbNullString
        txtsohd.Text = vbNullString
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
        lbltongDV.Text = "TT DV " & Cbolydo.Text
        txtTongtienDv.Text = SumColumnItem(6, Cbolydo.Text)
    End Sub
    Private Sub SaveReceicepts()
        Dim lvitem As ListViewItem
        Dim i As Integer
        Try
            For i = 0 To ListViewDetail.Items.Count - 1
                lvitem = ListViewDetail.Items(i)
                SaveReceipt(lvitem)       ' Luu chi tiet thu
            Next
            MsgBox("Đã lưu vào hệ thống!!")
            cmdLuu.Enabled = False
            txttongtien.Text = "0"
            txtTongtienDv.Text = "0"
        Catch ex As Exception
            MsgBox("Ghi dữ liệu lỗi", MsgBoxStyle.Critical, "Lổi ghi.")
        End Try
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
        txtTongtienDv.Text = "0"
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
        Arrvar(14) = "txttenNguoiNop1"
        Arrvar(15) = "txttenphieu"

        ' gan gia tri cua bien vao
        Arrval(0) = "Đơn vị: " & cbostations.Text
        Arrval(1) = "Địa chỉ: " & GetStringName(cbostations.SelectedValue, "Tbl_Stations")
        Arrval(2) = DateTimePickerngay.Value.Day
        Arrval(3) = DateTimePickerngay.Value.Month
        Arrval(4) = DateTimePickerngay.Value.Year

        Arrval(7) = txtEmployeeName.Text
        Arrval(8) = CboEmploy_code.Text

    End Sub

    Private Sub AssignVal(ByVal listitem As ListViewItem)
        Dim strMaLoai As String
        strMaLoai = listitem.SubItems(10).Text

        Arrval(5) = listitem.SubItems(9).Text ' So thu tu
        Arrval(6) = listitem.SubItems(8).Text ' Quyen so
        Arrval(9) = listitem.SubItems(7).Text & ", Kỳ cước " & listitem.SubItems(2).Text & ". " ' Ly do nop
        Arrval(10) = listitem.SubItems(6).Text ' So tien 
        numbers.Number = CLng(Arrval(10))
        Dim str As String
        str = numbers.NumbersToString
        If (Len(str) > 0) Then
            str = str.Substring(0, 1).ToUpper + str.Remove(0, 1)
        Else
            str = "Không"
        End If
        Arrval(11) = str + " đồng." ' tien bang chu
        Arrval(12) = "- " & listitem.SubItems(3).Text & "  bảng kê :" & listitem.SubItems(4).Text & "."  ' Chung tu goc 1

        If (strMaLoai = "UNC") Then
            Arrval(15) = "BIÊN LAI THU UNC"
            Arrval(13) = "- " & listitem.SubItems(5).Text & "  hóa đơn. " & listitem.SubItems(14).Text & " giấy uỷ nhiệm chi." ' Chung tu goc 2
        Else
            If (strMaLoai = "GNT") Then
                Arrval(15) = "BIÊN LAI THU GNT"
                Arrval(13) = "- " & listitem.SubItems(5).Text & "  hóa đơn. 1 giấy nộp tiền số : " & listitem.SubItems(11).Text & "."  ' Chung tu goc 2
            Else
                Arrval(15) = "PHIẾU THU"
                Arrval(13) = "- " & listitem.SubItems(5).Text & "  hóa đơn."  ' Chung tu goc 2
            End If
        End If

        Arrval(14) = txtEmployeeName.Text
        SetFieldTextOjectReports(rpt, Arrvar, Arrval)

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

    Private Sub Prints()
        Dim i As Integer = 0
        Dim listitem As ListViewItem
        Try
            For i = 0 To ListViewDetail.Items.Count - 1
                listitem = ListViewDetail.Items(i)
                AssignVal(listitem)
                Try
                    rpt.PrintToPrinter(NumericUpDownSobanin.Value, True, 1, 1)
                Catch ex As Exception
                    MsgBox("Không tìm thấy máy in " & strPrinterName & ".", MsgBoxStyle.Critical, "Lỗi in")
                    Exit For
                End Try
                'rpt.PrintToPrinter(1, True, 1, 1)
            Next
            MsgBox("Đã in xong : " & i & " phiếu thu.")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click
        If (ListViewDetail.Items.Count < 1) Then
            MsgBox("Không có phiếu thu để xem!")
            Exit Sub
        End If
        rpt = New CrystalReport_Receipts

        AssignVar()
        Dim frm As New frmPreview
        Dim listitem As ListViewItem
        listitem = ListViewDetail.FocusedItem
        If IsNothing(listitem) Then
            listitem = ListViewDetail.Items(0)
        End If
        AssignVal(listitem)
        frm.CrystalReportViewerReceipts.ReportSource = rpt
        frm.ShowDialog()
    End Sub

    'luu vao so quy
    Private Sub SaveToExpenses(ByVal lvitem As ListViewItem)
        Dim balance As Long
        'luu so tien da thu vao so quy
        strSQL = " INSERT INTO Tbl_Receipts_Expenses(Recei_Expen_Date,Recei_No,Descriptions,Recei_Money) " & _
        " VALUES('" & DateTimePickerngay.Value.ToShortDateString & _
        "'," & CInt(Trim$(lvitem.SubItems(9).Text)) & _
        ",'" & CboEmploy_code.Text & " nộp DV " & lvitem.SubItems(1).Text & ",Kỳ cước " & lvitem.SubItems(2).Text & _
        "'," & CLng(lvitem.SubItems(6).Text) & ")"
        ExcuxeSQL(strSQL)
        '",'" & lvitem.SubItems(7).Text & " Kỳ cước " & lvitem.SubItems(2).Text & _

    End Sub

    Private Function GetBalance(ByVal strQuery As String) As Long
        Dim value As Long
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
                    value = oleread.GetValue(0)
                End If
            End If
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :" & ex.ToString)
        End Try
        oledbcon.Close()
        Return value
    End Function


    Private Sub cmbHTthu_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHTthu.SelectedIndexChanged
        Dim strMaLoaiThu As String
        Try
            strMaLoaiThu = cmbHTthu.SelectedValue
            If (strMaLoaiThu = "UNC") Then
                txtslunc.Visible = True
                Label19.Visible = True
            Else
                txtslunc.Visible = False
                Label19.Visible = False
            End If

            If (strMaLoaiThu = "GNT") Then
                GroupBox3.Enabled = True
            Else
                GroupBox3.Enabled = False
            End If
            GetNumber_Vol_Ord()
            Computing_Vol_Ord()

        Catch ex As Exception
        End Try
    End Sub

    Private Sub cboAccounts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAccounts.SelectedIndexChanged

        If (start) Then
            Try
                txtbankname.Text = cboAccounts.SelectedValue
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub DateTimePickerngay_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePickerngay.Validated
        GetNumber_Vol_Ord()
        Computing_Vol_Ord()
    End Sub

    Private Sub txtslunc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtslunc.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdThemPT.Focus()
        End If
    End Sub

    Private Sub txtsoGNT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsoGNT.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerngaynop.Focus()
        End If
    End Sub

    Private Sub DateTimePickerngaynop_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerngaynop.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cboAccounts.Focus()
        End If
    End Sub

    Private Sub cboAccounts_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboAccounts.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdThemPT.Focus()
        End If
    End Sub

    'Kiem tra danh sach cac phieu thu
    Private Function CheckReciepts_No() As Boolean
        Dim result As Boolean = True
        Dim strQuery As String
        Dim lvitem As ListViewItem
        Dim SoPT As Long
        Dim strLoaiPhieu As String
        Dim i As Integer
        Try
            For i = 0 To ListViewDetail.Items.Count - 1
                lvitem = ListViewDetail.Items(i)
                SoPT = lvitem.SubItems(9).Text
                strLoaiPhieu = lvitem.SubItems(10).Text
                strQuery = " SELECT Ordinal_No FROM Tbl_Receipts WHERE MONTH(Receipt_Date) =" & DateTimePickerngay.Value.Month & " AND YEAR(Receipt_Date) =" & DateTimePickerngay.Value.Year & " AND Ordinal_No = " & SoPT & " AND MaLoaiThu ='" & strLoaiPhieu & "'"
                If (CheckReciept_No(strQuery)) Then
                    result = False
                    MsgBox("Số phiếu thu :" & SoPT & " đã tồn tại. Vui lòng kiểm tra lại.", MsgBoxStyle.Critical)
                    Exit For
                End If
            Next
        Catch ex As Exception
        End Try
        Return result

    End Function


End Class
