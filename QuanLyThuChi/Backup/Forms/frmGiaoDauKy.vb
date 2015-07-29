Imports System.Data.OleDb
Imports QuanLyCTV.KnowDotNet.KDNGrid
Imports QuanLyCTV.DataGridTextBoxCombo
Public Class frmGiaoDauKy
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private ds As DataSet
    Dim start As Boolean = False
    Private strID As String
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
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdIn As System.Windows.Forms.Button
    Friend WithEvents cmdLuu As System.Windows.Forms.Button
    Friend WithEvents cmddeletes As System.Windows.Forms.Button
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents Banthucuoc As System.Windows.Forms.TabPage
    Friend WithEvents TabkhachhangCTV As System.Windows.Forms.TabPage
    Friend WithEvents DataGridListEmployee As System.Windows.Forms.DataGrid
    Friend WithEvents DateTimePickerchuky As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtSLHD As System.Windows.Forms.TextBox
    Friend WithEvents txtDes As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmdlends As System.Windows.Forms.Button
    Friend WithEvents DataGridListStation As System.Windows.Forms.DataGrid
    Friend WithEvents CboDichvutram As System.Windows.Forms.ComboBox
    Friend WithEvents cboDichvuNV As System.Windows.Forms.ComboBox
    Friend WithEvents txtTongtienTothu As System.Windows.Forms.TextBox
    Friend WithEvents txtSLTBC As System.Windows.Forms.TextBox
    Friend WithEvents txtSLTBCNV As System.Windows.Forms.TextBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents cmdTaoMoi As System.Windows.Forms.Button
    Friend WithEvents txttongtienNV As System.Windows.Forms.TextBox
    Friend WithEvents txtTongHD As System.Windows.Forms.TextBox
    Friend WithEvents txtTongTBC As System.Windows.Forms.TextBox
    Friend WithEvents txtTongTien As System.Windows.Forms.TextBox
    Friend WithEvents txtSLHDNV As System.Windows.Forms.TextBox
    Friend WithEvents cmdThemmoiNV As System.Windows.Forms.Button
    Friend WithEvents cmdLenDSNV As System.Windows.Forms.Button
    Friend WithEvents DateTimePickerChukyNV As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtIDNV As System.Windows.Forms.TextBox
    Friend WithEvents cmdCapNhatNV As System.Windows.Forms.Button
    Friend WithEvents cmdLuuNV As System.Windows.Forms.Button
    Friend WithEvents cmdXoaNV As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmGiaoDauKy))
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cmdIn = New System.Windows.Forms.Button
        Me.cmdLuu = New System.Windows.Forms.Button
        Me.cmddeletes = New System.Windows.Forms.Button
        Me.cmddong = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.Banthucuoc = New System.Windows.Forms.TabPage
        Me.cmdTaoMoi = New System.Windows.Forms.Button
        Me.cmdlends = New System.Windows.Forms.Button
        Me.CboDichvutram = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtDes = New System.Windows.Forms.TextBox
        Me.txtTongtienTothu = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSLTBC = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSLHD = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.DateTimePickerchuky = New System.Windows.Forms.DateTimePicker
        Me.Label11 = New System.Windows.Forms.Label
        Me.DataGridListStation = New System.Windows.Forms.DataGrid
        Me.txtID = New System.Windows.Forms.TextBox
        Me.TabkhachhangCTV = New System.Windows.Forms.TabPage
        Me.txtIDNV = New System.Windows.Forms.TextBox
        Me.cmdThemmoiNV = New System.Windows.Forms.Button
        Me.cmdLenDSNV = New System.Windows.Forms.Button
        Me.cmdCapNhatNV = New System.Windows.Forms.Button
        Me.cmdLuuNV = New System.Windows.Forms.Button
        Me.cmdXoaNV = New System.Windows.Forms.Button
        Me.cboDichvuNV = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txttongtienNV = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSLTBCNV = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtSLHDNV = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.DateTimePickerChukyNV = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.DataGridListEmployee = New System.Windows.Forms.DataGrid
        Me.txtTongHD = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtTongTBC = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtTongTien = New System.Windows.Forms.TextBox
        Me.TabControl1.SuspendLayout()
        Me.Banthucuoc.SuspendLayout()
        CType(Me.DataGridListStation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabkhachhangCTV.SuspendLayout()
        CType(Me.DataGridListEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(64, 42)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(592, 27)
        Me.cbostations.TabIndex = 54
        '
        'Label15
        '
        Me.Label15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label15.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label15.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label15.Location = New System.Drawing.Point(2, 2)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(652, 38)
        Me.Label15.TabIndex = 69
        Me.Label15.Text = "GIAO HÓA ĐƠN, TBC ĐẦU KỲ "
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(8, 44)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 22)
        Me.Label5.TabIndex = 70
        Me.Label5.Text = "Tổ thu"
        '
        'cmdIn
        '
        Me.cmdIn.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdIn.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdIn.Location = New System.Drawing.Point(332, 91)
        Me.cmdIn.Name = "cmdIn"
        Me.cmdIn.Size = New System.Drawing.Size(80, 27)
        Me.cmdIn.TabIndex = 84
        Me.cmdIn.Text = "Thay đổi"
        '
        'cmdLuu
        '
        Me.cmdLuu.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuu.Location = New System.Drawing.Point(520, 91)
        Me.cmdLuu.Name = "cmdLuu"
        Me.cmdLuu.Size = New System.Drawing.Size(80, 27)
        Me.cmdLuu.TabIndex = 83
        Me.cmdLuu.Text = "Lưu"
        '
        'cmddeletes
        '
        Me.cmddeletes.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddeletes.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddeletes.Location = New System.Drawing.Point(425, 91)
        Me.cmddeletes.Name = "cmddeletes"
        Me.cmddeletes.Size = New System.Drawing.Size(80, 27)
        Me.cmddeletes.TabIndex = 82
        Me.cmddeletes.Text = "Xóa"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(566, 450)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(80, 27)
        Me.cmddong.TabIndex = 85
        Me.cmddong.Text = "Đóng"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.Banthucuoc)
        Me.TabControl1.Controls.Add(Me.TabkhachhangCTV)
        Me.TabControl1.Location = New System.Drawing.Point(0, 72)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(656, 376)
        Me.TabControl1.TabIndex = 1
        '
        'Banthucuoc
        '
        Me.Banthucuoc.Controls.Add(Me.cmdTaoMoi)
        Me.Banthucuoc.Controls.Add(Me.cmdlends)
        Me.Banthucuoc.Controls.Add(Me.CboDichvutram)
        Me.Banthucuoc.Controls.Add(Me.Label17)
        Me.Banthucuoc.Controls.Add(Me.Label12)
        Me.Banthucuoc.Controls.Add(Me.txtDes)
        Me.Banthucuoc.Controls.Add(Me.txtTongtienTothu)
        Me.Banthucuoc.Controls.Add(Me.Label2)
        Me.Banthucuoc.Controls.Add(Me.txtSLTBC)
        Me.Banthucuoc.Controls.Add(Me.Label1)
        Me.Banthucuoc.Controls.Add(Me.txtSLHD)
        Me.Banthucuoc.Controls.Add(Me.Label6)
        Me.Banthucuoc.Controls.Add(Me.DateTimePickerchuky)
        Me.Banthucuoc.Controls.Add(Me.Label11)
        Me.Banthucuoc.Controls.Add(Me.DataGridListStation)
        Me.Banthucuoc.Controls.Add(Me.cmdIn)
        Me.Banthucuoc.Controls.Add(Me.cmdLuu)
        Me.Banthucuoc.Controls.Add(Me.cmddeletes)
        Me.Banthucuoc.Controls.Add(Me.txtID)
        Me.Banthucuoc.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Banthucuoc.ForeColor = System.Drawing.Color.Chocolate
        Me.Banthucuoc.Location = New System.Drawing.Point(4, 22)
        Me.Banthucuoc.Name = "Banthucuoc"
        Me.Banthucuoc.Size = New System.Drawing.Size(648, 350)
        Me.Banthucuoc.TabIndex = 0
        Me.Banthucuoc.Text = "Nhận HĐ - TBC đầu kỳ"
        '
        'cmdTaoMoi
        '
        Me.cmdTaoMoi.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTaoMoi.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdTaoMoi.Location = New System.Drawing.Point(240, 91)
        Me.cmdTaoMoi.Name = "cmdTaoMoi"
        Me.cmdTaoMoi.Size = New System.Drawing.Size(80, 27)
        Me.cmdTaoMoi.TabIndex = 98
        Me.cmdTaoMoi.Text = "Tạo mới"
        '
        'cmdlends
        '
        Me.cmdlends.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdlends.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdlends.Location = New System.Drawing.Point(72, 91)
        Me.cmdlends.Name = "cmdlends"
        Me.cmdlends.Size = New System.Drawing.Size(80, 27)
        Me.cmdlends.TabIndex = 97
        Me.cmdlends.Text = "Lên DS"
        '
        'CboDichvutram
        '
        Me.CboDichvutram.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboDichvutram.ItemHeight = 19
        Me.CboDichvutram.Location = New System.Drawing.Point(300, 5)
        Me.CboDichvutram.Name = "CboDichvutram"
        Me.CboDichvutram.Size = New System.Drawing.Size(137, 27)
        Me.CboDichvutram.TabIndex = 95
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label17.Location = New System.Drawing.Point(220, 8)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(55, 21)
        Me.Label17.TabIndex = 96
        Me.Label17.Text = "Dịch vụ"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label12.Location = New System.Drawing.Point(8, 67)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(42, 21)
        Me.Label12.TabIndex = 94
        Me.Label12.Text = "Mô tả"
        '
        'txtDes
        '
        Me.txtDes.BackColor = System.Drawing.Color.White
        Me.txtDes.Location = New System.Drawing.Point(71, 64)
        Me.txtDes.Name = "txtDes"
        Me.txtDes.Size = New System.Drawing.Size(572, 25)
        Me.txtDes.TabIndex = 5
        Me.txtDes.Text = ""
        '
        'txtTongtienTothu
        '
        Me.txtTongtienTothu.BackColor = System.Drawing.Color.White
        Me.txtTongtienTothu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongtienTothu.ForeColor = System.Drawing.Color.Black
        Me.txtTongtienTothu.Location = New System.Drawing.Point(504, 35)
        Me.txtTongtienTothu.Name = "txtTongtienTothu"
        Me.txtTongtienTothu.Size = New System.Drawing.Size(136, 26)
        Me.txtTongtienTothu.TabIndex = 4
        Me.txtTongtienTothu.Text = "0"
        Me.txtTongtienTothu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(440, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 21)
        Me.Label2.TabIndex = 92
        Me.Label2.Text = "Tổng tiền"
        '
        'txtSLTBC
        '
        Me.txtSLTBC.BackColor = System.Drawing.Color.White
        Me.txtSLTBC.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSLTBC.ForeColor = System.Drawing.Color.Black
        Me.txtSLTBC.Location = New System.Drawing.Point(71, 35)
        Me.txtSLTBC.Name = "txtSLTBC"
        Me.txtSLTBC.Size = New System.Drawing.Size(137, 26)
        Me.txtSLTBC.TabIndex = 2
        Me.txtSLTBC.Text = "0"
        Me.txtSLTBC.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(8, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 21)
        Me.Label1.TabIndex = 90
        Me.Label1.Text = "SL TBC"
        '
        'txtSLHD
        '
        Me.txtSLHD.BackColor = System.Drawing.Color.White
        Me.txtSLHD.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSLHD.ForeColor = System.Drawing.Color.Black
        Me.txtSLHD.Location = New System.Drawing.Point(300, 35)
        Me.txtSLHD.Name = "txtSLHD"
        Me.txtSLHD.Size = New System.Drawing.Size(137, 26)
        Me.txtSLHD.TabIndex = 3
        Me.txtSLHD.Text = "0"
        Me.txtSLHD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(220, 39)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 21)
        Me.Label6.TabIndex = 88
        Me.Label6.Text = "SL HĐ"
        '
        'DateTimePickerchuky
        '
        Me.DateTimePickerchuky.CustomFormat = "MM/yyyy"
        Me.DateTimePickerchuky.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerchuky.Location = New System.Drawing.Point(71, 6)
        Me.DateTimePickerchuky.Name = "DateTimePickerchuky"
        Me.DateTimePickerchuky.Size = New System.Drawing.Size(137, 25)
        Me.DateTimePickerchuky.TabIndex = 0
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label11.Location = New System.Drawing.Point(8, 10)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(58, 21)
        Me.Label11.TabIndex = 86
        Me.Label11.Text = "Kỳ cước"
        '
        'DataGridListStation
        '
        Me.DataGridListStation.DataMember = ""
        Me.DataGridListStation.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridListStation.Location = New System.Drawing.Point(3, 120)
        Me.DataGridListStation.Name = "DataGridListStation"
        Me.DataGridListStation.Size = New System.Drawing.Size(646, 232)
        Me.DataGridListStation.TabIndex = 82
        '
        'txtID
        '
        Me.txtID.AutoSize = False
        Me.txtID.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtID.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtID.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.ForeColor = System.Drawing.Color.Black
        Me.txtID.Location = New System.Drawing.Point(504, 5)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(136, 24)
        Me.txtID.TabIndex = 91
        Me.txtID.Text = ""
        Me.txtID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TabkhachhangCTV
        '
        Me.TabkhachhangCTV.Controls.Add(Me.txtIDNV)
        Me.TabkhachhangCTV.Controls.Add(Me.cmdThemmoiNV)
        Me.TabkhachhangCTV.Controls.Add(Me.cmdLenDSNV)
        Me.TabkhachhangCTV.Controls.Add(Me.cmdCapNhatNV)
        Me.TabkhachhangCTV.Controls.Add(Me.cmdLuuNV)
        Me.TabkhachhangCTV.Controls.Add(Me.cmdXoaNV)
        Me.TabkhachhangCTV.Controls.Add(Me.cboDichvuNV)
        Me.TabkhachhangCTV.Controls.Add(Me.Label9)
        Me.TabkhachhangCTV.Controls.Add(Me.CboEmploy_code)
        Me.TabkhachhangCTV.Controls.Add(Me.txtEmployeeName)
        Me.TabkhachhangCTV.Controls.Add(Me.Label10)
        Me.TabkhachhangCTV.Controls.Add(Me.txttongtienNV)
        Me.TabkhachhangCTV.Controls.Add(Me.Label3)
        Me.TabkhachhangCTV.Controls.Add(Me.txtSLTBCNV)
        Me.TabkhachhangCTV.Controls.Add(Me.Label4)
        Me.TabkhachhangCTV.Controls.Add(Me.txtSLHDNV)
        Me.TabkhachhangCTV.Controls.Add(Me.Label7)
        Me.TabkhachhangCTV.Controls.Add(Me.DateTimePickerChukyNV)
        Me.TabkhachhangCTV.Controls.Add(Me.Label8)
        Me.TabkhachhangCTV.Controls.Add(Me.DataGridListEmployee)
        Me.TabkhachhangCTV.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabkhachhangCTV.Location = New System.Drawing.Point(4, 22)
        Me.TabkhachhangCTV.Name = "TabkhachhangCTV"
        Me.TabkhachhangCTV.Size = New System.Drawing.Size(648, 350)
        Me.TabkhachhangCTV.TabIndex = 1
        Me.TabkhachhangCTV.Text = "Giao hóa đơn - TBC cho cộng tác viên"
        '
        'txtIDNV
        '
        Me.txtIDNV.AutoSize = False
        Me.txtIDNV.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtIDNV.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtIDNV.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIDNV.ForeColor = System.Drawing.Color.Black
        Me.txtIDNV.Location = New System.Drawing.Point(488, 9)
        Me.txtIDNV.Name = "txtIDNV"
        Me.txtIDNV.ReadOnly = True
        Me.txtIDNV.Size = New System.Drawing.Size(144, 24)
        Me.txtIDNV.TabIndex = 116
        Me.txtIDNV.Text = ""
        Me.txtIDNV.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdThemmoiNV
        '
        Me.cmdThemmoiNV.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdThemmoiNV.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdThemmoiNV.Location = New System.Drawing.Point(283, 101)
        Me.cmdThemmoiNV.Name = "cmdThemmoiNV"
        Me.cmdThemmoiNV.Size = New System.Drawing.Size(80, 27)
        Me.cmdThemmoiNV.TabIndex = 115
        Me.cmdThemmoiNV.Text = "Tạo mới"
        '
        'cmdLenDSNV
        '
        Me.cmdLenDSNV.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLenDSNV.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLenDSNV.Location = New System.Drawing.Point(96, 100)
        Me.cmdLenDSNV.Name = "cmdLenDSNV"
        Me.cmdLenDSNV.Size = New System.Drawing.Size(80, 27)
        Me.cmdLenDSNV.TabIndex = 114
        Me.cmdLenDSNV.Text = "Lên DS"
        '
        'cmdCapNhatNV
        '
        Me.cmdCapNhatNV.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCapNhatNV.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdCapNhatNV.Location = New System.Drawing.Point(371, 100)
        Me.cmdCapNhatNV.Name = "cmdCapNhatNV"
        Me.cmdCapNhatNV.Size = New System.Drawing.Size(80, 27)
        Me.cmdCapNhatNV.TabIndex = 113
        Me.cmdCapNhatNV.Text = "Cập nhật"
        '
        'cmdLuuNV
        '
        Me.cmdLuuNV.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuuNV.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuuNV.Location = New System.Drawing.Point(547, 100)
        Me.cmdLuuNV.Name = "cmdLuuNV"
        Me.cmdLuuNV.Size = New System.Drawing.Size(80, 27)
        Me.cmdLuuNV.TabIndex = 6
        Me.cmdLuuNV.Text = "Lưu"
        '
        'cmdXoaNV
        '
        Me.cmdXoaNV.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdXoaNV.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdXoaNV.Location = New System.Drawing.Point(459, 100)
        Me.cmdXoaNV.Name = "cmdXoaNV"
        Me.cmdXoaNV.Size = New System.Drawing.Size(80, 27)
        Me.cmdXoaNV.TabIndex = 111
        Me.cmdXoaNV.Text = "Xóa"
        '
        'cboDichvuNV
        '
        Me.cboDichvuNV.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDichvuNV.ItemHeight = 19
        Me.cboDichvuNV.Location = New System.Drawing.Point(282, 8)
        Me.cboDichvuNV.Name = "cboDichvuNV"
        Me.cboDichvuNV.Size = New System.Drawing.Size(137, 27)
        Me.cboDichvuNV.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label9.Location = New System.Drawing.Point(218, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(55, 21)
        Me.Label9.TabIndex = 110
        Me.Label9.Text = "Dịch vụ"
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Location = New System.Drawing.Point(75, 72)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(205, 25)
        Me.CboEmploy_code.TabIndex = 5
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(280, 72)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(360, 26)
        Me.txtEmployeeName.TabIndex = 108
        Me.txtEmployeeName.Text = ""
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label10.Location = New System.Drawing.Point(8, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 21)
        Me.Label10.TabIndex = 107
        Me.Label10.Text = "CTV "
        '
        'txttongtienNV
        '
        Me.txttongtienNV.BackColor = System.Drawing.Color.White
        Me.txttongtienNV.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttongtienNV.ForeColor = System.Drawing.Color.Black
        Me.txttongtienNV.Location = New System.Drawing.Point(487, 40)
        Me.txttongtienNV.Name = "txttongtienNV"
        Me.txttongtienNV.Size = New System.Drawing.Size(152, 26)
        Me.txttongtienNV.TabIndex = 4
        Me.txttongtienNV.Text = ""
        Me.txttongtienNV.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(420, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 21)
        Me.Label3.TabIndex = 105
        Me.Label3.Text = "Tổng tiền"
        '
        'txtSLTBCNV
        '
        Me.txtSLTBCNV.BackColor = System.Drawing.Color.White
        Me.txtSLTBCNV.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSLTBCNV.ForeColor = System.Drawing.Color.Black
        Me.txtSLTBCNV.Location = New System.Drawing.Point(72, 40)
        Me.txtSLTBCNV.Name = "txtSLTBCNV"
        Me.txtSLTBCNV.Size = New System.Drawing.Size(137, 26)
        Me.txtSLTBCNV.TabIndex = 2
        Me.txtSLTBCNV.Text = ""
        Me.txtSLTBCNV.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(8, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 21)
        Me.Label4.TabIndex = 103
        Me.Label4.Text = "SL TBC"
        '
        'txtSLHDNV
        '
        Me.txtSLHDNV.AcceptsReturn = True
        Me.txtSLHDNV.BackColor = System.Drawing.Color.White
        Me.txtSLHDNV.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSLHDNV.ForeColor = System.Drawing.Color.Black
        Me.txtSLHDNV.Location = New System.Drawing.Point(280, 40)
        Me.txtSLHDNV.Name = "txtSLHDNV"
        Me.txtSLHDNV.Size = New System.Drawing.Size(137, 26)
        Me.txtSLHDNV.TabIndex = 3
        Me.txtSLHDNV.Text = ""
        Me.txtSLHDNV.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label7.Location = New System.Drawing.Point(224, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(50, 21)
        Me.Label7.TabIndex = 101
        Me.Label7.Text = "SL HĐ"
        '
        'DateTimePickerChukyNV
        '
        Me.DateTimePickerChukyNV.CustomFormat = "MM/yyyy"
        Me.DateTimePickerChukyNV.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerChukyNV.Location = New System.Drawing.Point(75, 8)
        Me.DateTimePickerChukyNV.Name = "DateTimePickerChukyNV"
        Me.DateTimePickerChukyNV.Size = New System.Drawing.Size(137, 25)
        Me.DateTimePickerChukyNV.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(9, 11)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(58, 21)
        Me.Label8.TabIndex = 99
        Me.Label8.Text = "Kỳ cước"
        '
        'DataGridListEmployee
        '
        Me.DataGridListEmployee.DataMember = ""
        Me.DataGridListEmployee.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridListEmployee.Location = New System.Drawing.Point(1, 128)
        Me.DataGridListEmployee.Name = "DataGridListEmployee"
        Me.DataGridListEmployee.Size = New System.Drawing.Size(647, 224)
        Me.DataGridListEmployee.TabIndex = 69
        '
        'txtTongHD
        '
        Me.txtTongHD.AutoSize = False
        Me.txtTongHD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTongHD.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongHD.ForeColor = System.Drawing.Color.Brown
        Me.txtTongHD.Location = New System.Drawing.Point(208, 451)
        Me.txtTongHD.Name = "txtTongHD"
        Me.txtTongHD.Size = New System.Drawing.Size(72, 25)
        Me.txtTongHD.TabIndex = 94
        Me.txtTongHD.Text = ""
        Me.txtTongHD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label13.Location = New System.Drawing.Point(144, 456)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(49, 16)
        Me.Label13.TabIndex = 95
        Me.Label13.Text = "Tổng HĐ"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label14.Location = New System.Drawing.Point(8, 456)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 16)
        Me.Label14.TabIndex = 97
        Me.Label14.Text = "Tổng TBC"
        '
        'txtTongTBC
        '
        Me.txtTongTBC.AutoSize = False
        Me.txtTongTBC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTongTBC.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongTBC.ForeColor = System.Drawing.Color.Brown
        Me.txtTongTBC.Location = New System.Drawing.Point(72, 451)
        Me.txtTongTBC.Name = "txtTongTBC"
        Me.txtTongTBC.Size = New System.Drawing.Size(72, 25)
        Me.txtTongTBC.TabIndex = 96
        Me.txtTongTBC.Text = ""
        Me.txtTongTBC.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label16.Location = New System.Drawing.Point(312, 456)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(55, 16)
        Me.Label16.TabIndex = 99
        Me.Label16.Text = "Tổng Tiền"
        '
        'txtTongTien
        '
        Me.txtTongTien.AutoSize = False
        Me.txtTongTien.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTongTien.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongTien.ForeColor = System.Drawing.Color.Brown
        Me.txtTongTien.Location = New System.Drawing.Point(376, 451)
        Me.txtTongTien.Name = "txtTongTien"
        Me.txtTongTien.Size = New System.Drawing.Size(168, 25)
        Me.txtTongTien.TabIndex = 98
        Me.txtTongTien.Text = ""
        Me.txtTongTien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'frmGiaoDauKy
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(658, 479)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.txtTongTien)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.txtTongTBC)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.txtTongHD)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.cmddong)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cbostations)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmGiaoDauKy"
        Me.Text = "Giao khách hàng đầu kỳ"
        Me.TabControl1.ResumeLayout(False)
        Me.Banthucuoc.ResumeLayout(False)
        CType(Me.DataGridListStation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabkhachhangCTV.ResumeLayout(False)
        CType(Me.DataGridListEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Public Sub FillDataSet()
        strSQL = "SELECT StationID,Station_Name,Station_Address FROM Tbl_Stations "
        FillCombo(cbostations, strSQL, "Tbl_Stations", "Station_Name", "StationID")

        strSQL = "SELECT Service_Code,Service_Name FROM Tbl_Services "
        FillDataSet(strSQL, "Tbl_Services")
        CboDichvutram.DataSource = mydataset.Tables("Tbl_Services").DefaultView
        CboDichvutram.DisplayMember = "Service_Code"
        CboDichvutram.ValueMember = "Service_Code"

        cboDichvuNV.DataSource = mydataset.Tables("Tbl_Services").DefaultView
        cboDichvuNV.DisplayMember = "Service_Code"
        cboDichvuNV.ValueMember = "Service_Code"
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

    Private Sub FillCombo(ByVal strQuery As String, ByVal strTablename As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, strTablename)
            cmd.Dispose()
            da.Dispose()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub FillDataset(ByVal strQuery As String, ByVal tblName As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, tblName)
            cmd.Dispose()
            da.Dispose()
        Catch ex As Exception
            MsgBox("Lổi View :" & ex.ToString)
        End Try
    End Sub

    Private Sub AddColunmSTT(ByVal strtable As String)
        Dim myType As System.Type
        myType = System.Type.GetType("System.String")
        mydataset.Tables(strtable).Columns.Add(New System.Data.DataColumn("STT", myType))

        Dim i As Integer
        For i = 0 To mydataset.Tables(strtable).Rows.Count - 1
            mydataset.Tables(strtable).Rows(i).Item("STT") = i + 1
        Next
    End Sub

    Private Sub FormatDataGrid(ByVal dt As DataTable)
        With DataGridListEmployee
            .AllowNavigation = False
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Chi tiết phân HD - TBC cho nhân viên ...."
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim items() As String
        ReDim items(mydataset.Tables("Tbl_Services").Rows.Count - 1)
        mydataset.Tables("Tbl_Services").Rows.CopyTo(items, 0)
        Dim TblStyle = CGrid.GetTableStyle(dt)

        With TblStyle
            .BackColor = System.Drawing.Color.MintCream
            .ForeColor = System.Drawing.Color.Navy
            .GridLineColor = System.Drawing.Color.MediumBlue
            .HeaderBackColor = System.Drawing.Color.Lavender
            .HeaderForeColor = System.Drawing.Color.Navy
            .AlternatingBackColor = Color.LightGray
            .HeaderFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .RowHeaderWidth = 10
            .ReadOnly = False

            With .GridColumnStyles

                .Add(New CGridTextBoxStyle("STT", 50, HorizontalAlignment.Center, True, "STT", String.Empty, ""))

                .Add(New CGridComboBoxStyle("Service_Code", 100, HorizontalAlignment.Center, "Dịch vụ", String.Empty, items, ComboBoxStyle.DropDown))

                .Add(New CGridDateTimePickerStyle("Charge_Cycle", 100, True, "Kỳ cước", DateTimePickerFormat.Custom, "MM/yyyy", "MM/yyyy"))

                .Add(New CGridTextBoxStyle("Employ_Code", 220, HorizontalAlignment.Left, True, "                 Nhân viên ", String.Empty, ""))

                .Add(New CGridTextBoxStyle("Invoice_Quantity", 80, HorizontalAlignment.Center, True, "SL HĐ", String.Empty, ""))

                .Add(New CGridTextBoxStyle("Total_Money", 100, HorizontalAlignment.Right, True, "Tiền", String.Empty, ""))

                .Add(New CGridTextBoxStyle("List_Quantity", 80, HorizontalAlignment.Center, True, "SL BK", String.Empty, ""))
            End With
        End With
        CGrid.SetGridStyle(Me.DataGridListEmployee, dt, TblStyle)
        'CGrid.AddRowToTable(dt,  )
    End Sub

    Private Sub FillDataGrid()
        Try
            CGrid.ClearTableStyles(DataGridListEmployee)
            FormatDataGrid(mydataset.Tables("ListEmployee"))
        Catch ex As Exception
        End Try

    End Sub

    Private Sub FormatDataGridListEmployee()

        With DataGridListEmployee
            .AllowNavigation = False
            '.DataMember = "ListEmployee"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Chi tiết giao HĐ - TBC cho nhân viên ..."
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "ListEmployee"
            .BackColor = System.Drawing.Color.MintCream
            .ForeColor = System.Drawing.Color.Navy
            .GridLineColor = System.Drawing.Color.MediumBlue
            .HeaderBackColor = System.Drawing.Color.Lavender
            .HeaderForeColor = System.Drawing.Color.Navy
            .AlternatingBackColor = Color.LightGray
            .HeaderFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .RowHeaderWidth = 10
            .ReadOnly = False

            With .GridColumnStyles

                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "STT"
                    .HeaderText = "STT"
                    .Width = 40
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                'Dim ComboTextCol As New DataGridComboBoxColumn
                'ComboTextCol.MappingName = "Service_Code"
                'ComboTextCol.HeaderText = "        Dịch vụ"
                'ComboTextCol.Width = 120
                'ComboTextCol.ColumnComboBox.DataSource = mydataset.Tables("Tbl_Services").DefaultView
                'ComboTextCol.ColumnComboBox.DisplayMember = "Service_Code"
                'ComboTextCol.ColumnComboBox.ValueMember = "Service_Code"
                'TblStyle.PreferredRowHeight = ComboTextCol.ColumnComboBox.Height + 10
                'TblStyle.GridColumnStyles.Add(ComboTextCol)

                '.Add(New CGridDateTimePickerStyle("Charge_Cycle", 80, True, "Kỳ cước ", DateTimePickerFormat.Custom, "MM/yyyy", "MM/yyyy"))

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "Service_Code"
                    .HeaderText = "Dịch vụ"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With
                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Charge_Cycle"
                    .HeaderText = "Kỳ cước "
                    .Width = 80
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "Employ_Code"
                    .HeaderText = "         Nhân viên "
                    .Width = 150
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "List_Quantity"
                    .HeaderText = "SL TBC"
                    .Width = 70
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "Invoice_Quantity"
                    .HeaderText = "SL HĐ"
                    .Width = 70
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(6)
                    .MappingName = "Total_Money"
                    .HeaderText = "Tổng tiền"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(7)
                    .MappingName = "Recei_Date"
                    .HeaderText = "Ngày nhận"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
            End With

        End With
        DataGridListEmployee.TableStyles.Add(TblStyle)
        ' DataGridListEmployee.Table

    End Sub

    Private Sub FormatDataGridListStation()

        With DataGridListStation
            .AllowNavigation = False
            '.DataMember = "ListEmployee"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Chi tiết nhận HĐ - TBC ..."
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "ListStation"
            .BackColor = System.Drawing.Color.MintCream
            .ForeColor = System.Drawing.Color.Navy
            .GridLineColor = System.Drawing.Color.MediumBlue
            .HeaderBackColor = System.Drawing.Color.Lavender
            .HeaderForeColor = System.Drawing.Color.Navy
            .AlternatingBackColor = Color.LightGray
            .HeaderFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .RowHeaderWidth = 10
            .ReadOnly = True

            With .GridColumnStyles

                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "STT"
                    .HeaderText = "STT"
                    .Width = 40
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                'Dim ComboTextCol As New DataGridComboBoxColumn
                'ComboTextCol.MappingName = "Service_Code"
                'ComboTextCol.HeaderText = "        Dịch vụ"
                'ComboTextCol.Width = 120
                'ComboTextCol.ColumnComboBox.DataSource = mydataset.Tables("Tbl_Services").DefaultView
                'ComboTextCol.ColumnComboBox.DisplayMember = "Service_Code"
                'ComboTextCol.ColumnComboBox.ValueMember = "Service_Code"
                'TblStyle.PreferredRowHeight = ComboTextCol.ColumnComboBox.Height + 10
                'TblStyle.GridColumnStyles.Add(ComboTextCol)

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "Service_Code"
                    .HeaderText = "Dịch vụ"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                '.Add(New CGridDateTimePickerStyle("Charge_Cycle", 80, True, "Kỳ cước ", DateTimePickerFormat.Custom, "MM/yyyy", "MM/yyyy"))

                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Charge_Cycle"
                    .HeaderText = "Kỳ cước "
                    .Width = 80
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "List_Quantity"
                    .HeaderText = "SL TBC"
                    .Width = 70
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "Invoice_Quantity"
                    .HeaderText = "SL HĐ"
                    .Width = 70
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "Total_Money"
                    .HeaderText = "Tổng tiền"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(6)
                    .MappingName = "Descriptions"
                    .HeaderText = "            Diễn giải"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
            End With
        End With
        DataGridListStation.TableStyles.Add(TblStyle)
        ' DataGridListEmployee.Table

    End Sub

    Private Sub frmGiaoDauKy_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            mydataset = New DataSet
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


            strSQL = "SELECT * FROM Tbl_Employee_ChargeCycle WHERE Charge_Cycle = '00' "
            FillDataSet(strSQL, "ListEmployee")
            AddColunmSTT("ListEmployee")
            DataGridListEmployee.DataSource = mydataset.Tables("ListEmployee")
            FormatDataGridListEmployee()

            strSQL = "SELECT * FROM Tbl_Station_ChargeCycle WHERE Charge_Cycle = '00' "
            FillDataSet(strSQL, "ListStation")
            AddColunmSTT("ListStation")
            FormatDataGridListStation()
            DataGridListStation.DataSource = mydataset.Tables("ListStation")
            CutBindingTextBoxToThu()
            BindingTextBoxToThu()
            'FillDataGrid()
        Catch eLoad As System.Exception
            System.Windows.Forms.MessageBox.Show(eLoad.Message)
        End Try
    End Sub

    Private Sub cmddong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddong.Click
        Me.Close()
    End Sub

    Private Sub cmdlends_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdlends.Click
        FillDatagridToThu()
    End Sub

    Private Sub FillDatagridToThu()
        DataGridListStation.DataSource = Nothing
        mydataset.Tables("ListStation").Clear()
        strSQL = "SELECT * FROM Tbl_Station_ChargeCycle  WHERE Charge_Cycle = '" & DateTimePickerchuky.Text & "' "
        FillDataSet(strSQL, "ListStation")
        Try
            mydataset.Tables("ListStation").Columns.Remove("STT")
        Catch ex As Exception
        End Try
        AddColunmSTT("ListStation")
        DataGridListStation.DataSource = mydataset.Tables("ListStation")
        CutBindingTextBoxToThu()
        BindingTextBoxToThu()
        txtTongTBC.Text = SumColum(DataGridListStation, "List_Quantity")
        txtTongHD.Text = SumColum(DataGridListStation, "Invoice_Quantity")
        txtTongTien.Text = SumColum(DataGridListStation, "Total_Money")
    End Sub

    Private Sub cmdLuu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLuu.Click
        Dim strQuery As String
        If (checkInfor(1)) Then
            If (cmdLuu.Text = "Lưu") Then
                strQuery = " INSERT INTO Tbl_Station_ChargeCycle(StationID,Service_Code,Charge_Cycle,List_Quantity,Invoice_Quantity,Total_Money,Descriptions) VALUES('" & _
                            cbostations.SelectedValue & _
                            "','" & CboDichvutram.SelectedValue & _
                            "','" & DateTimePickerchuky.Text & _
                            "'," & CLng(Trim$(txtSLTBC.Text)) & _
                            "," & CLng(Trim$(txtSLHD.Text)) & _
                            "," & CLng(Trim$(txtTongtienTothu.Text)) & _
                            ",'" & txtDes.Text & "'" & _
                            ")"
                ExcuxeSQL(strQuery)
            Else
                strQuery = " UPDATE Tbl_Station_ChargeCycle SET Service_Code ='" & CboDichvutram.SelectedValue & _
                "',Charge_Cycle ='" & DateTimePickerchuky.Text & _
                "',List_Quantity= " & CLng(Trim$(txtSLTBC.Text)) & _
                ",Invoice_Quantity = " & CLng(Trim$(txtSLHD.Text)) & _
                ",Total_Money = " & CLng(Trim$(txtTongtienTothu.Text)) & _
                ",Descriptions = '" & txtDes.Text & _
                "',StationID = '" & cbostations.SelectedValue & _
                "' WHERE ID =" & txtID.Text
                ExcuxeSQL(strQuery)
            End If
            FillDatagridToThu()
        End If
        txtSLTBC.Focus()
    End Sub

    Private Sub CutBindingTextBoxToThu()
        CboDichvutram.DataBindings.Clear()
        txtSLHD.DataBindings.Clear()
        txtSLTBC.DataBindings.Clear()
        txtTongtienTothu.DataBindings.Clear()
        txtDes.DataBindings.Clear()
        txtID.DataBindings.Clear()
        txtSLHD.ReadOnly = False
        txtSLTBC.ReadOnly = False
        txtTongtienTothu.ReadOnly = False
        txtDes.ReadOnly = False

    End Sub
    Private Sub BindingTextBoxToThu()
        Try
            txtSLHD.ReadOnly = True
            txtSLTBC.ReadOnly = True
            txtTongtienTothu.ReadOnly = True
            txtDes.ReadOnly = True
            CboDichvutram.DataBindings.Add("Text", mydataset.Tables("ListStation"), "Service_Code")
            txtSLHD.DataBindings.Add("Text", mydataset.Tables("ListStation"), "Invoice_Quantity")
            txtSLTBC.DataBindings.Add("Text", mydataset.Tables("ListStation"), "List_Quantity")
            txtTongtienTothu.DataBindings.Add("Text", mydataset.Tables("ListStation"), "Total_Money")
            txtDes.DataBindings.Add("Text", mydataset.Tables("ListStation"), "Descriptions")
            txtID.DataBindings.Add("Text", mydataset.Tables("ListStation"), "ID")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub CutBindingTextBoxNV()
        cboDichvuNV.DataBindings.Clear()
        txtSLHDNV.DataBindings.Clear()
        txtSLTBCNV.DataBindings.Clear()
        txttongtienNV.DataBindings.Clear()
        txtIDNV.DataBindings.Clear()
        txtSLHDNV.ReadOnly = False
        txtSLTBCNV.ReadOnly = False
        txttongtienNV.ReadOnly = False

    End Sub
    Private Sub BindingTextBoxNV()
        Try
            txtSLHDNV.ReadOnly = True
            txtSLTBCNV.ReadOnly = True
            txttongtienNV.ReadOnly = True
            cboDichvuNV.DataBindings.Add("Text", mydataset.Tables("ListEmployee"), "Service_Code")
            txtSLHDNV.DataBindings.Add("Text", mydataset.Tables("ListEmployee"), "Invoice_Quantity")
            txtSLTBCNV.DataBindings.Add("Text", mydataset.Tables("ListEmployee"), "List_Quantity")
            txttongtienNV.DataBindings.Add("Text", mydataset.Tables("ListEmployee"), "Total_Money")
            txtIDNV.DataBindings.Add("Text", mydataset.Tables("ListEmployee"), "ID")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function checkInfor(ByVal index As Short) As Boolean
        Select Case index
            Case 1
                If (CboDichvutram.Text = "") Then
                    MsgBox("Chưa chọn dịch vụ để thêm vào", MsgBoxStyle.Critical)
                    CboDichvutram.Focus()
                    Return False
                End If

                If (Trim$(txtSLHD.Text) = "") Then
                    MsgBox("Số lượng hóa đơn chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLHD.Focus()
                    Return False
                End If

                If (Not IsNumeric(Trim$(txtSLHD.Text))) Then
                    MsgBox("Số lượng hóa đơn phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLHD.Focus()
                    txtSLHD.SelectAll()
                    Return False
                End If

                If (Trim$(txtSLTBC.Text) = "") Then
                    MsgBox("Số lượng TBC chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLTBC.Focus()
                    Return False
                End If

                If (Not IsNumeric(Trim$(txtSLTBC.Text))) Then
                    MsgBox("Số lượng TBC phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLTBC.Focus()
                    txtSLTBC.SelectAll()
                    Return False
                End If

                If (Trim$(txtTongtienTothu.Text) = "") Then
                    MsgBox("Tổng tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtTongtienTothu.Focus()
                    Return False
                End If

                If (Not IsNumeric(Trim$(txtTongtienTothu.Text))) Then
                    MsgBox("Tổng tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtTongtienTothu.Focus()
                    txtTongtienTothu.SelectAll()
                    Return False
                End If
            Case 2
                If (cboDichvuNV.Text = "") Then
                    MsgBox("Chưa chọn dịch vụ để thêm vào", MsgBoxStyle.Critical)
                    cboDichvuNV.Focus()
                    Return False
                End If

                If (Trim$(txtSLHDNV.Text) = "") Then
                    MsgBox("Số lượng hóa đơn chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLHDNV.Focus()
                    Return False
                End If

                If (Not IsNumeric(Trim$(txtSLHDNV.Text))) Then
                    MsgBox("Số lượng hóa đơn phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLHDNV.Focus()
                    txtSLHDNV.SelectAll()
                    Return False
                End If

                If (Trim$(txtSLTBCNV.Text) = "") Then
                    MsgBox("Số lượng TBC chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLTBCNV.Focus()
                    Return False
                End If

                If (Not IsNumeric(Trim$(txtSLTBCNV.Text))) Then
                    MsgBox("Số lượng TBC phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txtSLTBCNV.Focus()
                    txtSLTBCNV.SelectAll()
                    Return False
                End If

                If (Trim$(txttongtienNV.Text) = "") Then
                    MsgBox("Tổng tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txttongtienNV.Focus()
                    Return False
                End If

                If (Not IsNumeric(Trim$(txttongtienNV.Text))) Then
                    MsgBox("Tổng tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                    txttongtienNV.Focus()
                    txttongtienNV.SelectAll()
                    Return False
                End If
        End Select

        Return True
    End Function

    Private Sub cmddeletes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddeletes.Click
        Dim value
        If (txtID.Text = "") Then
            MsgBox("Không có dòng nào được chọn để xóa!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            Exit Sub
        End If
        value = MsgBox("Bạn có thật sự muốn xóa dòng này không!", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Xác nhận.")
        If (value = vbYes) Then
            Dim strquery As String
            strquery = "DELETE FROM Tbl_Station_ChargeCycle WHERE ID =" & txtID.Text
            ExcuxeSQL(strquery)
            FillDatagridToThu

        End If
    End Sub

    Private Sub cmdTaoMoi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTaoMoi.Click
        CutBindingTextBoxToThu()
        txtSLHD.Clear()
        txtSLTBC.Clear()
        txtTongtienTothu.Clear()
        txtDes.Clear()
        cmdLuu.Text = "Lưu"
    End Sub

    Private Sub cmdIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIn.Click
        CutBindingTextBoxToThu()
        cmdLuu.Text = "Cập nhật"
    End Sub

    Public Function SumColum(ByVal dgr As DataGrid, ByVal strcolname As String) As Long
        Dim result As Long = 0
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        dt = dgr.DataSource
        For i = 0 To dt.Rows.Count - 1
            result += dt.Rows(i).Item(strcolname)
        Next
        Return result
    End Function

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

    Private Sub cmdLenDSNV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLenDSNV.Click
        FillDatagridNV()
    End Sub

    Private Sub FillDatagridNV()
        DataGridListEmployee.DataSource = Nothing
        mydataset.Tables("ListEmployee").Clear()
        strSQL = "SELECT * FROM Tbl_Employee_ChargeCycle  WHERE Charge_Cycle = '" & DateTimePickerChukyNV.Text & "' "
        FillDataSet(strSQL, "ListEmployee")
        Try
            mydataset.Tables("ListEmployee").Columns.Remove("STT")
        Catch ex As Exception
        End Try
        AddColunmSTT("ListEmployee")
        DataGridListEmployee.DataSource = mydataset.Tables("ListEmployee")
        CutBindingTextBoxNV()
        BindingTextBoxNV()
        txtTongTBC.Text = SumColum(DataGridListEmployee, "List_Quantity")
        txtTongHD.Text = SumColum(DataGridListEmployee, "Invoice_Quantity")
        txtTongTien.Text = SumColum(DataGridListEmployee, "Total_Money")
    End Sub
    Private Sub cmdThemmoiNV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdThemmoiNV.Click
        CutBindingTextBoxNV()
        txtSLHDNV.Clear()
        txtSLTBCNV.Clear()
        txttongtienNV.Clear()
        cmdLuuNV.Text = "Lưu"
    End Sub

    Private Sub cmdLuuNV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLuuNV.Click
        Dim strQuery As String
        If (checkInfor(2)) Then
            If (cmdLuuNV.Text = "Lưu") Then
                strQuery = " INSERT INTO Tbl_Employee_ChargeCycle(Employ_Code,Service_Code,Charge_Cycle,List_Quantity,Invoice_Quantity,Total_Money) VALUES('" & _
                            CboEmploy_code.Text & _
                            "','" & cboDichvuNV.SelectedValue & _
                            "','" & DateTimePickerChukyNV.Text & _
                            "'," & CLng(Trim$(txtSLTBCNV.Text)) & _
                            "," & CLng(Trim$(txtSLHDNV.Text)) & _
                            "," & CLng(Trim$(txttongtienNV.Text)) & ")"
                ExcuxeSQL(strQuery)
            Else
                strQuery = " UPDATE Tbl_Employee_ChargeCycle SET Service_Code ='" & cboDichvuNV.SelectedValue & _
                "',Charge_Cycle ='" & DateTimePickerChukyNV.Text & _
                "',List_Quantity= " & CLng(Trim$(txtSLTBCNV.Text)) & _
                ",Invoice_Quantity = " & CLng(Trim$(txtSLHDNV.Text)) & _
                ",Total_Money = " & CLng(Trim$(txttongtienNV.Text)) & _
                ",Employ_Code = '" & CboEmploy_code.Text & _
                "' WHERE ID =" & txtIDNV.Text
                ExcuxeSQL(strQuery)
            End If
            FillDatagridNV()
        End If
        txtSLTBCNV.Focus()
    End Sub

    Private Sub cmdCapNhatNV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCapNhatNV.Click
        CutBindingTextBoxNV()
        cmdLuuNV.Text = "Cập nhật"
    End Sub

    Private Sub cmdXoaNV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdXoaNV.Click
        Dim value
        If (txtIDNV.Text = "") Then
            MsgBox("Không có dòng nào được chọn để xóa!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            Exit Sub
        End If
        value = MsgBox("Bạn có thật sự muốn xóa dòng này không!", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Xác nhận.")
        If (value = vbYes) Then
            Dim strquery As String
            strquery = "DELETE FROM Tbl_Employee_ChargeCycle WHERE ID =" & txtIDNV.Text
            ExcuxeSQL(strquery)
            FillDatagridNV()
        End If
    End Sub

    Private Sub DataGridListStation_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridListStation.Navigate

    End Sub

    Private Sub DateTimePickerchuky_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerchuky.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            CboDichvutram.Focus()
        End If
    End Sub

    Private Sub CboDichvutram_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CboDichvutram.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            txtSLTBC.Focus()
        End If
    End Sub

    Private Sub txtSLTBC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSLTBC.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            txtSLHD.Focus()
        End If
    End Sub

    Private Sub txtSLHD_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSLHD.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            txtTongtienTothu.Focus()
        End If
    End Sub

    Private Sub txtTongtienTothu_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTongtienTothu.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            txtDes.Focus()
        End If
    End Sub

    Private Sub txtDes_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDes.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            cmdLuu.Focus()
        End If
    End Sub

    Private Sub DateTimePickerChukyNV_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerChukyNV.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            cboDichvuNV.Focus()
        End If
    End Sub

    Private Sub cboDichvuNV_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboDichvuNV.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            txtSLTBCNV.Focus()
        End If
    End Sub

    Private Sub txtSLTBCNV_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSLTBCNV.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            txtSLHDNV.Focus()
        End If
    End Sub

    Private Sub txtSLHDNV_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSLHDNV.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            txttongtienNV.Focus()
        End If
    End Sub

    Private Sub txttongtienNV_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttongtienNV.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            CboEmploy_code.Focus()
        End If
    End Sub

    Private Sub CboEmploy_code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CboEmploy_code.KeyPress
        Dim Ascii As Short = Asc(e.KeyChar)
        If (Ascii = 13) Then
            cmdLuuNV.Focus()
        End If
    End Sub
End Class
