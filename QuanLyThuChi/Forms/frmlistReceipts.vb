Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports QuanLyCTV.KnowDotNet.KDNGrid

Public Class frmlistReceipts
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Dim start As Boolean = False
    Private strListReceiptno As String
    Private strListReceiptID As String
    Dim splitn As New SplitNumbers
    Private Indexlistview As Integer
    Dim Dsrpt As New DsGiayNopTien
    Dim Status As Boolean
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        FillDataset()
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
            If (CboEmploy_code.Items.Count > 0) Then
                txtEmployeeName.Text = CboEmploy_code.SelectedValue
            End If
        End If

        txttongtien.ReadOnly = True
        DateTimePickertungay.Value = Now
        DateTimePickerdenngay.Value = Now
        DateTimePickerNgayPC.Value = Now
        SetDetailListView()
        ListViewDetail.Visible = False
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
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents cmdLuu As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cmdlist As System.Windows.Forms.Button
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerchuky As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txttongtien As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents DataGridListReceipts As System.Windows.Forms.DataGrid
    Friend WithEvents DateTimePickertungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePickerdenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txttongsophieu As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtTongBK As System.Windows.Forms.TextBox
    Friend WithEvents txtTongHD As System.Windows.Forms.TextBox
    Friend WithEvents cmdlapphieuchi As System.Windows.Forms.Button
    Friend WithEvents txtdenso As System.Windows.Forms.TextBox
    Friend WithEvents txttuso As System.Windows.Forms.TextBox
    Friend WithEvents cmdIn As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtsobknt As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerNgayPC As System.Windows.Forms.DateTimePicker
    Friend WithEvents ListViewDetail As System.Windows.Forms.ListView
    Friend WithEvents CheckBoxStatus As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmlistReceipts))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CheckBoxStatus = New System.Windows.Forms.CheckBox
        Me.ListViewDetail = New System.Windows.Forms.ListView
        Me.DateTimePickerNgayPC = New System.Windows.Forms.DateTimePicker
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtsobknt = New System.Windows.Forms.TextBox
        Me.DateTimePickerdenngay = New System.Windows.Forms.DateTimePicker
        Me.txtTongBK = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtTongHD = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txttongsophieu = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.DateTimePickertungay = New System.Windows.Forms.DateTimePicker
        Me.DataGridListReceipts = New System.Windows.Forms.DataGrid
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.txttuso = New System.Windows.Forms.TextBox
        Me.DateTimePickerchuky = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtdenso = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.txttongtien = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdlapphieuchi = New System.Windows.Forms.Button
        Me.cmdlist = New System.Windows.Forms.Button
        Me.cmdIn = New System.Windows.Forms.Button
        Me.cmddong = New System.Windows.Forms.Button
        Me.cmdLuu = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridListReceipts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBoxStatus)
        Me.GroupBox1.Controls.Add(Me.ListViewDetail)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerNgayPC)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.txtsobknt)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerdenngay)
        Me.GroupBox1.Controls.Add(Me.txtTongBK)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.txtTongHD)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.txttongsophieu)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.DateTimePickertungay)
        Me.GroupBox1.Controls.Add(Me.DataGridListReceipts)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.txttuso)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerchuky)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtdenso)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.txttongtien)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(4, 41)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 432)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'CheckBoxStatus
        '
        Me.CheckBoxStatus.Location = New System.Drawing.Point(24, 80)
        Me.CheckBoxStatus.Name = "CheckBoxStatus"
        Me.CheckBoxStatus.Size = New System.Drawing.Size(72, 24)
        Me.CheckBoxStatus.TabIndex = 76
        Me.CheckBoxStatus.Text = "Đã chi"
        '
        'ListViewDetail
        '
        Me.ListViewDetail.FullRowSelect = True
        Me.ListViewDetail.GridLines = True
        Me.ListViewDetail.Location = New System.Drawing.Point(8, 152)
        Me.ListViewDetail.Name = "ListViewDetail"
        Me.ListViewDetail.Size = New System.Drawing.Size(696, 240)
        Me.ListViewDetail.TabIndex = 75
        Me.ListViewDetail.View = System.Windows.Forms.View.Details
        '
        'DateTimePickerNgayPC
        '
        Me.DateTimePickerNgayPC.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerNgayPC.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerNgayPC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerNgayPC.Location = New System.Drawing.Point(417, 43)
        Me.DateTimePickerNgayPC.Name = "DateTimePickerNgayPC"
        Me.DateTimePickerNgayPC.Size = New System.Drawing.Size(104, 26)
        Me.DateTimePickerNgayPC.TabIndex = 7
        Me.DateTimePickerNgayPC.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(347, 43)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(66, 22)
        Me.Label17.TabIndex = 74
        Me.Label17.Text = "Ngày PC"
        '
        'txtsobknt
        '
        Me.txtsobknt.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsobknt.ForeColor = System.Drawing.Color.Black
        Me.txtsobknt.Location = New System.Drawing.Point(648, 44)
        Me.txtsobknt.Name = "txtsobknt"
        Me.txtsobknt.Size = New System.Drawing.Size(56, 26)
        Me.txtsobknt.TabIndex = 8
        Me.txtsobknt.Text = ""
        Me.txtsobknt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DateTimePickerdenngay
        '
        Me.DateTimePickerdenngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerdenngay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerdenngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerdenngay.Location = New System.Drawing.Point(243, 13)
        Me.DateTimePickerdenngay.Name = "DateTimePickerdenngay"
        Me.DateTimePickerdenngay.Size = New System.Drawing.Size(104, 26)
        Me.DateTimePickerdenngay.TabIndex = 2
        Me.DateTimePickerdenngay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'txtTongBK
        '
        Me.txtTongBK.AutoSize = False
        Me.txtTongBK.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtTongBK.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongBK.ForeColor = System.Drawing.Color.Brown
        Me.txtTongBK.Location = New System.Drawing.Point(432, 401)
        Me.txtTongBK.Name = "txtTongBK"
        Me.txtTongBK.Size = New System.Drawing.Size(64, 24)
        Me.txtTongBK.TabIndex = 72
        Me.txtTongBK.Text = ""
        Me.txtTongBK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label16.Location = New System.Drawing.Point(344, 402)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(86, 22)
        Me.Label16.TabIndex = 71
        Me.Label16.Text = "Tổng số BK"
        '
        'txtTongHD
        '
        Me.txtTongHD.AutoSize = False
        Me.txtTongHD.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtTongHD.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongHD.ForeColor = System.Drawing.Color.Brown
        Me.txtTongHD.Location = New System.Drawing.Point(270, 401)
        Me.txtTongHD.Name = "txtTongHD"
        Me.txtTongHD.Size = New System.Drawing.Size(64, 24)
        Me.txtTongHD.TabIndex = 70
        Me.txtTongHD.Text = ""
        Me.txtTongHD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label12.Location = New System.Drawing.Point(184, 403)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(87, 22)
        Me.Label12.TabIndex = 69
        Me.Label12.Text = "Tổng số HĐ"
        '
        'txttongsophieu
        '
        Me.txttongsophieu.AutoSize = False
        Me.txttongsophieu.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txttongsophieu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttongsophieu.ForeColor = System.Drawing.Color.Brown
        Me.txttongsophieu.Location = New System.Drawing.Point(108, 402)
        Me.txttongsophieu.Name = "txttongsophieu"
        Me.txttongsophieu.Size = New System.Drawing.Size(64, 24)
        Me.txttongsophieu.TabIndex = 68
        Me.txttongsophieu.Text = ""
        Me.txttongsophieu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label10.Location = New System.Drawing.Point(15, 404)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 67
        Me.Label10.Text = "Tổng số phiếu"
        '
        'DateTimePickertungay
        '
        Me.DateTimePickertungay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickertungay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickertungay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickertungay.Location = New System.Drawing.Point(64, 13)
        Me.DateTimePickertungay.Name = "DateTimePickertungay"
        Me.DateTimePickertungay.Size = New System.Drawing.Size(104, 26)
        Me.DateTimePickertungay.TabIndex = 1
        Me.DateTimePickertungay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'DataGridListReceipts
        '
        Me.DataGridListReceipts.DataMember = ""
        Me.DataGridListReceipts.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridListReceipts.Location = New System.Drawing.Point(8, 152)
        Me.DataGridListReceipts.Name = "DataGridListReceipts"
        Me.DataGridListReceipts.Size = New System.Drawing.Size(696, 240)
        Me.DataGridListReceipts.TabIndex = 66
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(177, 15)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(69, 22)
        Me.Label8.TabIndex = 65
        Me.Label8.Text = "Đến ngày"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label14.Location = New System.Drawing.Point(5, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(700, 2)
        Me.Label14.TabIndex = 34
        '
        'txttuso
        '
        Me.txttuso.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttuso.ForeColor = System.Drawing.Color.Black
        Me.txttuso.Location = New System.Drawing.Point(64, 42)
        Me.txttuso.Name = "txttuso"
        Me.txttuso.Size = New System.Drawing.Size(104, 26)
        Me.txttuso.TabIndex = 5
        Me.txttuso.Text = ""
        Me.txttuso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DateTimePickerchuky
        '
        Me.DateTimePickerchuky.CustomFormat = "MM/yyyy"
        Me.DateTimePickerchuky.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerchuky.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerchuky.Location = New System.Drawing.Point(609, 13)
        Me.DateTimePickerchuky.Name = "DateTimePickerchuky"
        Me.DateTimePickerchuky.Size = New System.Drawing.Size(99, 26)
        Me.DateTimePickerchuky.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(177, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 22)
        Me.Label1.TabIndex = 63
        Me.Label1.Text = "Đến số"
        '
        'txtdenso
        '
        Me.txtdenso.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdenso.ForeColor = System.Drawing.Color.Black
        Me.txtdenso.Location = New System.Drawing.Point(243, 42)
        Me.txtdenso.Name = "txtdenso"
        Me.txtdenso.Size = New System.Drawing.Size(104, 26)
        Me.txtdenso.TabIndex = 6
        Me.txtdenso.Text = ""
        Me.txtdenso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 22)
        Me.Label2.TabIndex = 61
        Me.Label2.Text = "Từ số"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(520, 14)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(90, 22)
        Me.Label11.TabIndex = 62
        Me.Label11.Text = "Chu kỳ cước"
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Location = New System.Drawing.Point(243, 77)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(180, 27)
        Me.CboEmploy_code.TabIndex = 9
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(424, 77)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(280, 26)
        Me.txtEmployeeName.TabIndex = 57
        Me.txtEmployeeName.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(112, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(124, 22)
        Me.Label4.TabIndex = 56
        Me.Label4.Text = "NV nhận/nộp tiền"
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(417, 13)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(104, 27)
        Me.Cbolydo.TabIndex = 3
        '
        'txttongtien
        '
        Me.txttongtien.BackColor = System.Drawing.Color.White
        Me.txttongtien.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttongtien.ForeColor = System.Drawing.Color.Brown
        Me.txttongtien.Location = New System.Drawing.Point(592, 400)
        Me.txttongtien.Name = "txttongtien"
        Me.txttongtien.Size = New System.Drawing.Size(112, 26)
        Me.txttongtien.TabIndex = 50
        Me.txttongtien.Text = ""
        Me.txttongtien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 22)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Từ ngày"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(345, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Dịch vụ"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label13.Location = New System.Drawing.Point(504, 400)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(69, 22)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Tổng tiền"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdlapphieuchi)
        Me.GroupBox2.Controls.Add(Me.cmdlist)
        Me.GroupBox2.Controls.Add(Me.cmdIn)
        Me.GroupBox2.Controls.Add(Me.cmddong)
        Me.GroupBox2.Controls.Add(Me.cmdLuu)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 101)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(696, 45)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        '
        'cmdlapphieuchi
        '
        Me.cmdlapphieuchi.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdlapphieuchi.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdlapphieuchi.Location = New System.Drawing.Point(131, 13)
        Me.cmdlapphieuchi.Name = "cmdlapphieuchi"
        Me.cmdlapphieuchi.Size = New System.Drawing.Size(108, 27)
        Me.cmdlapphieuchi.TabIndex = 13
        Me.cmdlapphieuchi.Text = "Lập phiếu chi "
        '
        'cmdlist
        '
        Me.cmdlist.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdlist.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdlist.Location = New System.Drawing.Point(9, 13)
        Me.cmdlist.Name = "cmdlist"
        Me.cmdlist.Size = New System.Drawing.Size(108, 27)
        Me.cmdlist.TabIndex = 11
        Me.cmdlist.Text = "Lên DS"
        '
        'cmdIn
        '
        Me.cmdIn.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdIn.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdIn.Location = New System.Drawing.Point(248, 13)
        Me.cmdIn.Name = "cmdIn"
        Me.cmdIn.Size = New System.Drawing.Size(108, 27)
        Me.cmdIn.TabIndex = 11
        Me.cmdIn.Text = "In bảng kê"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(584, 13)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(108, 27)
        Me.cmddong.TabIndex = 31
        Me.cmddong.Text = "Đóng"
        '
        'cmdLuu
        '
        Me.cmdLuu.Enabled = False
        Me.cmdLuu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuu.Location = New System.Drawing.Point(375, 13)
        Me.cmdLuu.Name = "cmdLuu"
        Me.cmdLuu.Size = New System.Drawing.Size(108, 27)
        Me.cmdLuu.TabIndex = 12
        Me.cmdLuu.Text = "Lưu"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label9.Location = New System.Drawing.Point(520, 47)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(136, 22)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "Số BK nộp tiền NH"
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(522, 16)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(192, 27)
        Me.cbostations.TabIndex = 53
        '
        'Label15
        '
        Me.Label15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label15.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label15.Font = New System.Drawing.Font("Arial", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label15.Location = New System.Drawing.Point(4, 3)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(710, 40)
        Me.Label15.TabIndex = 52
        Me.Label15.Text = "LẬP BẢNG KÊ NỘP TIỀN NGÂN HÀNG"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'frmlistReceipts
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 478)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmlistReceipts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Lập bảng kê nộp tiền"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridListReceipts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    ' Private Indexlistview As Integer
    Private ds As DataSet
    Private ds1 As DataSet
    Private Sub FormatDataGridListReceipts()

        With DataGridListReceipts
            .AllowNavigation = False
            .DataMember = "ListReceipts"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách phiếu thu ...."
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "ListReceipts"
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
                ' Set datagrid ColumnStyle for ID field

                .Add(New DataGridBoolColumn)
                With .Item(0)
                    .MappingName = "Check"
                    .HeaderText = "In"
                    .Width = 40
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "Ordinal_No"
                    .HeaderText = "Số PT"
                    .Width = 50
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
                '.Add(New DataGridDateTimePicker)
                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Receipt_Date"
                    .HeaderText = "    Ngày thu"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "Employ_Code"
                    .HeaderText = "                 Nhân viên "
                    .Width = 220
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "Invoice_Quantity"
                    .HeaderText = "SL HĐ"
                    .Width = 80
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "Total_Money"
                    .HeaderText = "Tiền"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(6)
                    .MappingName = "List_Quantity"
                    .HeaderText = "SL BK"
                    .Width = 80
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

            End With
        End With
        DataGridListReceipts.TableStyles.Add(TblStyle)
    End Sub

    Private Sub FillDataset(ByVal strQuery As String)
        Try
            ds = New DataSet
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds, "ListReceipts")

            Dim myTypeCheck As System.Type

            myTypeCheck = System.Type.GetType("System.Boolean")
            ds.Tables("ListReceipts").Columns.Add(New System.Data.DataColumn("Check", myTypeCheck))

            Dim i As Integer
            For i = 0 To ds.Tables("ListReceipts").Rows.Count - 1
                ds.Tables("ListReceipts").Rows(i).Item("Check") = True
            Next
        Catch ex As Exception
            MsgBox("Lổi View :" & ex.ToString)
        End Try
    End Sub

    Private Sub frmlistReceipts_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'FormatDataGridListReceipts()
        strSQL = "SELECT Employ_Code,Receipt_Date,Invoice_Quantity,Total_Money,List_Quantity,Ordinal_No,ID FROM Tbl_Receipts WHERE Charge_Cycle = #01/01/2000# "
        FillDataset(strSQL)
        FillDataGrid()
    End Sub

    Private Sub cmddong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddong.Click
        Me.Close()
    End Sub

    Private Sub cmdlist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdlist.Click
        Status = CheckBoxStatus.Checked
        Indexlistview = 0
        txttongsophieu.Text = "0"
        txtTongBK.Text = "0"
        txttongtien.Text = "0"
        txtTongHD.Text = "0"
        strListReceiptID = "0"
        strListReceiptno = "0"
        ds.Clear()
        'DataGridListReceipts.DataSource = Nothing
        FillReceipts()
        'DataGridListReceipts.DataSource = ds.Tables("ListReceipts")
        FillDataGrid()
        If (ds.Tables("ListReceipts").Rows.Count > 0) Then
            txttongsophieu.Text = ds.Tables("ListReceipts").Rows.Count
            txtTongBK.Text = SumMoney("List_Quantity")
            txttongtien.Text = SumMoney("Total_Money")
            splitn.strnumbers = txttongtien.Text
            txttongtien.Text = splitn.Splitnumer(",")
            txtTongHD.Text = SumMoney("Invoice_Quantity")
            strListReceiptID = BillStrListReceipts(DataGridListReceipts, "ID")
            strListReceiptno = BillStrListReceipts(DataGridListReceipts, "Ordinal_No")
        End If

        If (Cbolydo.Text = "PSTNDL") Then
            AddListView()
            ListViewDetail.Visible = True
        Else
            ListViewDetail.Visible = False
        End If

        Dim strQuery As String
        strQuery = " SELECT MAX(Ordinal_No_List)+1  FROM Tbl_Expenses WHERE ( MONTH(Expense_Date) = " & DateTimePickerNgayPC.Value.Month & ") AND  (YEAR(Expense_Date) = " & DateTimePickerNgayPC.Value.Year() & ")"
        txtsobknt.Text = GetMaxNumber(strQuery)
    End Sub

    Private Sub FillReceipts()
        If checkSQL() Then
            FillDataset(strSQL)
        End If
    End Sub

    Private Sub cmdLuu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLuu.Click
        UpdateReceipts()
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

    Public Sub UpdateReceipts()
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        Dim valueID As Long
        Dim strQuery As String
        dt = DataGridListReceipts.DataSource
        Try
            For i = 0 To dt.Rows.Count - 1
                value = dt.Rows(i).Item("Check")
                If (value) Then
                    valueID = dt.Rows(i).Item("ID")
                    strQuery = " UPDATE Tbl_Receipts SET Status = True WHERE ID = " & valueID
                    UpdateReceipt(strQuery)
                End If
            Next
            MsgBox("Đã cập nhật. Xác nhận các phiếu thu đã chi nộp tiền vào ngân hàng.")
        Catch ex As Exception
        End Try

    End Sub

    Public Sub UpdateExpenes()
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        Dim valueID As Long
        Dim strQuery As String
        dt = DataGridListReceipts.DataSource
        Try
            For i = 0 To dt.Rows.Count - 1
                value = dt.Rows(i).Item("Check")
                If (value) Then
                    valueID = dt.Rows(i).Item("ID")
                    strQuery = " UPDATE Tbl_Receipts SET Status = True WHERE ID = " & valueID
                    UpdateReceipt(strQuery)
                End If
            Next
            MsgBox("Đã cập nhật. Xác nhận các phiếu thu đã chi nộp tiền vào ngân hàng.")
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
    Private Function SumMoney(ByVal columnname As String) As Long
        Dim table As New DataTable
        Dim result As Long = 0
        Try
            table = DataGridListReceipts.DataSource
            Dim i As Integer
            For i = 0 To table.Rows.Count - 1
                result += table.Rows(i).Item(columnname)
            Next
        Catch ex As Exception

        End Try

        Return result
    End Function

    Private Sub DataGridListReceipts_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridListReceipts.Validated
       
    End Sub

    Public Function SumColum(ByVal dgr As DataGrid, ByVal strcolname As String) As Long
        Dim result As Long = 0
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        dt = dgr.DataSource
        For i = 0 To dt.Rows.Count - 1
            value = dt.Rows(i).Item("Check")
            If (value) Then
                result += dt.Rows(i).Item(strcolname)
            End If
        Next
        Return result
    End Function

    Public Function CountSoPhieu(ByVal dgr As DataGrid) As Long
        Dim result As Long = 0
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        dt = dgr.DataSource
        For i = 0 To dt.Rows.Count - 1
            value = dt.Rows(i).Item("Check")
            If (value) Then
                result += 1
            End If
        Next
        Return result
    End Function

    Private Function BillStrListReceipts(ByVal dgr As DataGrid, ByVal strcolname As String) As String
        Dim strResult As String = ""
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        dt = dgr.DataSource
        Try

            For i = 0 To dt.Rows.Count - 1
                value = dt.Rows(i).Item("Check")
                If (value) Then
                    strResult += dt.Rows(i).Item(strcolname) & ","
                End If
            Next
        Catch ex As Exception
        End Try
        If (Trim$(strResult) <> "") Then
            strResult = strResult.Remove(strResult.Length - 1, 1)
        End If
        Return strResult
    End Function
    Private Sub cmdlapphieuchi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdlapphieuchi.Click
        If checkInfo() Then

            If (Trim$(txtsobknt.Text) = "") Then
                MsgBox("Số bảng kê nộp tiền ngân hàng chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsobknt.Focus()
                txtsobknt.SelectAll()
                Exit Sub
            End If

            If (Not IsNumeric(Trim$(txtsobknt.Text))) Then
                MsgBox("Số bảng kê nộp tiền ngân hàng phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsobknt.Focus()
                txtsobknt.SelectAll()
                Exit Sub
            End If

            If (CLng(Trim$(txtsobknt.Text)) < 1) Then
                MsgBox("Số bảng kê nộp tiền ngân hàng phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtsobknt.Focus()
                txtsobknt.SelectAll()
                Exit Sub
            End If

            Dim frm As New frmExpenses(cbostations.Text, Cbolydo.Text, CInt(txtTongHD.Text) + CInt(txtTongBK.Text), txtTongHD.Text & ": Hóa đơn, " & txtTongBK.Text & " Bảng kê.", txttongtien.Text, DateTimePickerchuky.Text, strListReceiptID, txtsobknt.Text, strListReceiptno, CboEmploy_code.Text, DateTimePickerNgayPC.Value, Status)
            frm.ShowDialog()
        End If

    End Sub

    Public Sub FillDataSet()
        mydataset = New DataSet
        strSQL = "SELECT StationID,Station_Name,Station_Address FROM Tbl_Stations "
        FillCombo(cbostations, strSQL, "Tbl_Stations", "Station_Name", "StationID")

        strSQL = "SELECT Service_Code,Service_Name FROM Tbl_Services "
        FillCombo(Cbolydo, strSQL, "Tbl_Services", "Service_Code", "Service_Name")

    End Sub

    Private Sub FillDatasetEx(ByVal strQuery As String)
        Try
            ds1 = New DataSet
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds1, "ListExpenes")
        Catch ex As Exception
            MsgBox("Lổi View :" & ex.ToString)
        End Try
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

    Private Function checkInfo() As Boolean
        Dim result As Boolean = True

        If (Trim$(txttongtien.Text) = "") Then
            MsgBox("Số tiền không có!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txttongtien.Focus()
            result = False
            GoTo endFunction
        End If

        If (Not IsNumeric(Trim$(txttongtien.Text))) Then
            MsgBox("Số tiền phải là số !", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txttongtien.Focus()
            txttongtien.SelectAll()
            result = False
            GoTo endFunction
        End If

        If (CLng(txttongtien.Text) = 0) Then
            MsgBox("Số tiền chi không có!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            result = False
            GoTo endFunction
        End If

        If (Trim$(Cbolydo.Text) = "") Then
            MsgBox("Chưa chọn loại dịch vụ để lập phiếu!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            Cbolydo.Focus()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txttuso.Text) <> "") Then
            If (Not IsNumeric(Trim$(txttuso.Text))) Then
                MsgBox("Điều kiện từ số phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txttuso.Focus()
                txttuso.SelectAll()
                result = False
                GoTo endFunction
            End If


            If (CLng(Trim$(txttuso.Text)) < 1) Then
                MsgBox("Điều kiện từ số phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txttuso.Focus()
                txttuso.SelectAll()
                Exit Function
            End If

        End If

        If (Trim$(txtdenso.Text) <> "") Then
            If (Not IsNumeric(Trim$(txtdenso.Text))) Then
                MsgBox("Điều kiện đến số phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtdenso.Focus()
                txtdenso.SelectAll()
                result = False
                GoTo endFunction
            End If

            If (CLng(Trim$(txtdenso.Text)) < 1) Then
                MsgBox("Điều kiện đến số phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtdenso.Focus()
                txtdenso.SelectAll()
                Exit Function
            End If
        End If


        'If (Trim$(txttumuctien.Text) <> "") Then
        '    If (Not IsNumeric(Trim$(txttumuctien.Text))) Then
        '        MsgBox("Điều kiện từ mức tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
        '        txttumuctien.Focus()
        '        txttumuctien.SelectAll()
        '        result = False
        '        GoTo endFunction
        '    End If
        'End If

        'If (Trim$(txtDenmuctien.Text) <> "") Then
        '    If (Not IsNumeric(Trim$(txtDenmuctien.Text))) Then
        '        MsgBox("Điều kiện từ mức tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
        '        txtDenmuctien.Focus()
        '        txtDenmuctien.SelectAll()
        '        result = False
        '        GoTo endFunction
        '    End If
        'End If

endFunction:
        Return result
    End Function

    Private Function checkSQL() As Boolean
        Dim result As Boolean = True

        If (Trim$(Cbolydo.Text) = "") Then
            MsgBox("Chưa chọn loại dịch vụ để lập phiếu!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            Cbolydo.Focus()
            result = False
            GoTo endFunction
        End If

        Select Case Cbolydo.Text
            Case "PSTNDL"
                strSQL = "SELECT Employ_Code,Receipt_Date,Invoice_Quantity,Total_Money,List_Quantity,Ordinal_No,ID,TenDaily,DiaChiDaily,STTDL FROM Tbl_Receipts WHERE  " & _
                         " Receipt_Date BETWEEN #" & DateTimePickertungay.Value.ToShortDateString & " # AND #" & DateTimePickerdenngay.Value.ToShortDateString & "# AND " & _
                         " Charge_Cycle = #" & DateTimePickerchuky.Text & "# AND MaLoaiThu ='TM'"
                strSQL += " AND (Service_Code ='PSTNDS' OR Service_Code  ='PSTNTNDN') "

            Case "PSTNDS", "PSTNTNDN"
                result = False
                GoTo endFunction
            Case Else
                strSQL = "SELECT Employ_Code,Receipt_Date,Invoice_Quantity,Total_Money,List_Quantity,Ordinal_No,ID FROM Tbl_Receipts WHERE  " & _
                         " Receipt_Date BETWEEN #" & DateTimePickertungay.Value.ToShortDateString & "# AND #" & DateTimePickerdenngay.Value.ToShortDateString & "# AND " & _
                         " Charge_Cycle = #" & DateTimePickerchuky.Text & "# AND MaLoaiThu ='TM'"
                strSQL += " AND Service_Code ='" & Cbolydo.Text & "'"

        End Select

        If (Trim$(txttuso.Text) <> "") Then
            If (Not IsNumeric(Trim$(txttuso.Text))) Then
                MsgBox("Điều kiện từ số phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txttuso.Focus()
                txttuso.SelectAll()
                result = False
                GoTo endFunction
            End If

            If (CLng(Trim$(txttuso.Text)) < 1) Then
                MsgBox("Điều kiện từ số phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txttuso.Focus()
                txttuso.SelectAll()
                Exit Function
            End If
            strSQL += " AND Ordinal_No >=" & CInt(txttuso.Text)
        End If

        If (Trim$(txtdenso.Text) <> "") Then
            If (Not IsNumeric(Trim$(txtdenso.Text))) Then
                MsgBox("Điều kiện đến số phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtdenso.Focus()
                txtdenso.SelectAll()
                result = False
                GoTo endFunction
            End If

            If (CLng(Trim$(txtdenso.Text)) < 1) Then
                MsgBox("Điều kiện đến số phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtdenso.Focus()
                txtdenso.SelectAll()
                Exit Function
            End If
            strSQL += " AND Ordinal_No <=" & CInt(txtdenso.Text)
        End If

        'If (Trim$(CboEmploy_code.Text) <> "") Then
        '    strSQL += " AND Employ_Code ='" & CboEmploy_code.Text & "'"
        'End If

        strSQL += " AND Status = " & CheckBoxStatus.Checked & " ORDER BY Receipt_Date,Ordinal_No "

endFunction:
        Return result
    End Function


    Private Sub cmdIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIn.Click
        If (Trim$(txtsobknt.Text) = "" OrElse Not IsNumeric(txtsobknt.Text)) Then
            MsgBox("Số BKNT chưa được nhập vào!")
            txtsobknt.Focus()
            Exit Sub
        End If

        Dim rpt As CrystalReportBKNT
        rpt = New CrystalReportBKNT
        Dsrpt.Clear()
        Dim strQuery As String

        'Lay thong tin trung tam va don vi thu cuoc
        strQuery = "SELECT Tbl_Countries.CountryName, Tbl_Stations.Station_Name FROM Tbl_Stations INNER JOIN Tbl_Countries ON Tbl_Stations.CountryCode = Tbl_Countries.CountryCode WHERE Tbl_Stations.StationID='" & cbostations.SelectedValue & "'"
        FillReports(strQuery, "GetCountry_station")

        'Fill : GNT, chukycuoc,Ngay,MaNhanVien,TenNhanVien,DichVu
        Dim Newrow As DataRow
        Newrow = Dsrpt.Tables("GetValuePar").NewRow

        Dsrpt.Tables("GetValuePar").Rows.Add(Newrow)
        Dsrpt.Tables("GetValuePar").Rows(0).Item("SBKNT") = CInt(txtsobknt.Text)
        Dsrpt.Tables("GetValuePar").Rows(0).Item("ChuKy") = DateTimePickerchuky.Text
        Dsrpt.Tables("GetValuePar").Rows(0).Item("DichVu") = Cbolydo.Text
        Dsrpt.Tables("GetValuePar").Rows(0).Item("MaNguoiNop") = CboEmploy_code.Text
        Dsrpt.Tables("GetValuePar").Rows(0).Item("TenNguoiNop") = txtEmployeeName.Text

        'Tao bang master 
        strQuery = "SELECT Employ_Code FROM Tbl_Employee "
        FillReports(strQuery, "QueryEmployeeCodes")

        AddTableDetail()
        rpt.SetDataSource(Dsrpt)
        Dim frm As New frmPreview
        frm.CrystalReportViewerReceipts.ReportSource = rpt
        frm.ShowDialog()
    End Sub

    Private Sub cbostations_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbostations.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickertungay.Focus()
        End If
    End Sub

    Private Sub DateTimePickertungay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickertungay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerdenngay.Focus()
        End If
    End Sub

    Private Sub DateTimePickerdenngay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerdenngay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
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
                Cbolydo.Focus()
                Exit Sub
            End If
            DateTimePickerchuky.Focus()
        End If
    End Sub

    Private Sub DateTimePickerchuky_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerchuky.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttuso.Focus()
        End If
    End Sub

    Private Sub txttuso_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttuso.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtdenso.Focus()
        End If
    End Sub

    Private Sub txtdenso_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdenso.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerNgayPC.Focus()
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
            cmdlist.Focus()
        End If
    End Sub

    Private Sub Export_ToExcel()
        Dim listitem As ListViewItem
        Dim index As Integer
        Dim MyXLApp As Excel.Application
        Dim MyXLBook As Excel.Workbook
        Dim MyXLWorksheet As Excel.Worksheet
        Dim SumMoneyEmployee As Long = 0
        Dim SumTotalMoney As Long = 0
        Dim SumTotalInvoice As Integer = 0
        Dim SumListQuantityEmployee As Integer = 0
        Dim SumInvoiceEmployee As Long = 0
        Dim SumInvoiceTotal As Long = 0
        Dim indexSum As Integer
        Try
            MyXLApp = CType(CreateObject("Excel.Application"), Excel.Application)
            MyXLBook = CType(MyXLApp.Workbooks.Add, Excel.Workbook)
            MyXLWorksheet = CType(MyXLBook.Worksheets(1), Excel.Worksheet)
            MyXLApp.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        With MyXLWorksheet

            .Columns().ColumnWidth = 15
            .Range("A1").ColumnWidth = 3
            .Range("A1").Value = "CÔNG TY THU CƯỚC VÀ DỊCH VỤ VIETTEL"
            .Range("B1").ColumnWidth = 40
            .Range("A1", "B1").Merge()
            .Range("A2").Value = "Trung tâm thu cước Hồ Chí Minh"
            .Range("A2", "B2").Merge()
            .Range("A3").Value = "Đơn vị thu cước " & cbostations.Text
            .Range("A3", "B3").Merge()

            .Range("C1").ColumnWidth = 7

            .Range("D1").ColumnWidth = 7
            .Range("D1").Value = "Mẫu số BC-10/BKNTNH"
            .Range("D1", "E1").Merge()
            .Range("D1", "E1").HorizontalAlignment = Alignment.HorizontalCenterAlign

            .Range("E1").ColumnWidth = 18
            .Range("E4").Value = "Bảng kê số: " & txtsobknt.Text
            .Range("E4").Font.Bold = True
            .Range("A5").Value = "BẢNG KÊ NỘP TIỀN CƯỚC VÀO TÀI KHOẢN NGÂN HÀNG KỲ CƯỚC " & DateTimePickerchuky.Text
            .Range("A5", "E5").Merge()
            .Range("A5", "E5").Font.Bold = True
            .Range("A5", "E5").HorizontalAlignment = Alignment.HorizontalCenterAlign
            .Range("A6").Value = "Ngày " & Now.Day & " Tháng  " & Now.Month & " Năm " & Now.Year
            .Range("A6", "E6").Merge()
            .Range("A7").Value = "Dịch vụ: " & Cbolydo.Text & " ;N.viên nộp tiền :.................................. Mã .................................."
            .Range("A7", "E7").Merge()
            .Range("A8").Value = "TT"
            .Range("A8").Font.Bold = True
            .Range("A8", "A9").Merge()
            .Range("A8", "A9").VerticalAlignment = Alignment.HorizontalCenterAlign
            .Range("B8").Value = "Nhân viên/bảng kê"
            .Range("B8").Font.Bold = True
            .Range("B8", "B9").Merge()
            .Range("B8", "B9").VerticalAlignment = Alignment.HorizontalCenterAlign

            .Range("C8").Value = "SL HĐ"
            .Range("C8").Font.Bold = True
            .Range("C8", "C9").Merge()
            .Range("C8", "C9").VerticalAlignment = Alignment.HorizontalCenterAlign

            .Range("D8").Value = "SL BK"
            .Range("D8").Font.Bold = True
            .Range("D8", "D9").Merge()
            .Range("D8", "D9").VerticalAlignment = Alignment.HorizontalCenterAlign

            .Range("E8").Value = "Số tiền"
            .Range("E8").Font.Bold = True
            .Range("E8", "E9").Merge()
            .Range("E8", "E9").VerticalAlignment = Alignment.HorizontalCenterAlign
            .Range("E8", "E9").HorizontalAlignment = Alignment.HorizontalCenterAlign

            Dim i, count As Integer
            index = 10
            If (Cbolydo.Text = "PSTNDL") Then
                For i = 0 To ListViewDetail.Items.Count - 1
                    listitem = ListViewDetail.Items(i)

                    .Range("B" & index.ToString).Value = listitem.SubItems(0).Text
                    If (Trim$(listitem.SubItems(1).Text) = "") Then
                        .Range("B" & index.ToString).Font.Bold = True
                    End If
                    '.Range("C" & index.ToString).Value = listitem.SubItems(2).Text
                    .Range("C" & index.ToString).Value = listitem.SubItems(3).Text
                    .Range("D" & index.ToString).Value = listitem.SubItems(5).Text
                    .Range("E" & index.ToString).Value = listitem.SubItems(4).Text

                    index += 1
                Next i
            Else
                Dim dt As DataTable
                Dim value As Boolean
                Dim strEmployee_Code As String
                dt = DataGridListReceipts.DataSource
                index = 10
                For i = 0 To dt.Rows.Count - 1
                    value = dt.Rows(i).Item("Check")
                    If (value) Then
                        strEmployee_Code = dt.Rows(i).Item("Employ_Code")
                        indexSum = index
                        .Range("B" & index.ToString).Value = "Nhân viên thu cước " & strEmployee_Code
                        .Range("B" & index.ToString).Font.Bold = True
                        .Range("B" & index.ToString).Font.Italic = True
                        count = i
                        index += 1
                        SumMoneyEmployee = 0
                        SumInvoiceEmployee = 0
                        SumListQuantityEmployee = 0
                        While count < dt.Rows.Count
                            If (value And (strEmployee_Code = dt.Rows(count).Item("Employ_Code"))) Then
                                .Range("B" & index.ToString).Value = "Phiếu thu số " & CStr(dt.Rows(count).Item("Ordinal_No"))
                                .Range("C" & index.ToString).Value = CStr(dt.Rows(count).Item("Invoice_Quantity"))
                                SumInvoiceEmployee += CInt(dt.Rows(count).Item("Invoice_Quantity"))
                                .Range("D" & index.ToString).Value = CStr(dt.Rows(count).Item("List_Quantity"))
                                SumListQuantityEmployee += CInt(dt.Rows(count).Item("List_Quantity"))
                                '.Range("D" & index.ToString).Value = CStr(dt.Rows(count).Item("Total_Money"))
                                SumMoneyEmployee += CLng(dt.Rows(count).Item("Total_Money"))
                                'Range("E" & index.ToString).Value = CStr(dt.Rows(count).Item("List_Quantity"))
                                .Range("E" & index.ToString).Value = CStr(dt.Rows(count).Item("Total_Money"))
                                dt.Rows(count).Item("Check") = False
                                index += 1
                            End If
                            count += 1
                            If (count < dt.Rows.Count - 1) Then
                                value = dt.Rows(count).Item("Check")
                            End If
                        End While

                        .Range("C" & indexSum.ToString).Value = SumInvoiceEmployee 'So hoa don theo nhan vien
                        .Range("C" & indexSum.ToString).Font.Bold = True
                        .Range("D" & indexSum.ToString).Value = SumListQuantityEmployee 'Tong so bang ke theo nhan vien
                        .Range("D" & indexSum.ToString).Font.Bold = True
                        .Range("E" & indexSum.ToString).Value = SumMoneyEmployee 'Tong tien theo nhan ven
                        .Range("E" & indexSum.ToString).Font.Bold = True

                    End If
                Next

                .Range("A" & index.ToString).Value = "Tổng cộng"
                .Range("A" & index.ToString, "B" & index.ToString).Merge()
                .Range("A" & index.ToString, "B" & index.ToString).Font.Bold = True
                .Range("A" & index.ToString, "B" & index.ToString).HorizontalAlignment = Alignment.HorizontalCenterAlign
            End If
            Dim cl As Color
            Dim border As Excel.Border
            '.Range("A8", "E" & index.ToString).Borders.Color = cl.Black
            '.Range("A8", "E" & index.ToString).Borders(Excel.XlBordersIndex.xlInsideHorizontal) =  border. 

            index += 2
            .Range("A" & index.ToString).Value = "        Người lập                        Kế toán                              Phụ trách đơn vị thu cước"
            .Range("A" & index.ToString, "E" & index.ToString).Merge()
            .Range("A" & index.ToString, "E" & index.ToString).Font.Bold = True
            .Range("A" & index.ToString, "E" & index.ToString).HorizontalAlignment = Alignment.HorizontalCenterAlign

        End With

    End Sub

    Private Sub Fill_Cells_ToExcel()


    End Sub

    Private Sub SetDetailListView()
        ListViewDetail.Items.Clear()
        ListViewDetail.Columns.Clear()
        ListViewDetail.Columns.Add("   Số phiếu thu/chi", 150, HorizontalAlignment.Left)
        ListViewDetail.Columns.Add("  Ngày thu/chi", 100, HorizontalAlignment.Center)
        ListViewDetail.Columns.Add("        Nhân viên ", 170, HorizontalAlignment.Left)
        ListViewDetail.Columns.Add("SL BK ", 80, HorizontalAlignment.Right)
        ListViewDetail.Columns.Add("SL HĐ ", 70, HorizontalAlignment.Right)
        ListViewDetail.Columns.Add("Tiền thu/chi    ", 120, HorizontalAlignment.Right)


    End Sub

    Private Sub AddSubItemListView(ByVal strSP As String, ByVal strngay As String, ByVal strNhanvien As String, ByVal strSHD As String, ByVal strTien As String, ByVal strSoBK As String)

        ListViewDetail.Items.Add(strSP)
        ListViewDetail.Items(Indexlistview).SubItems.Add(strngay)
        ListViewDetail.Items(Indexlistview).SubItems.Add(strNhanvien)
        ListViewDetail.Items(Indexlistview).SubItems.Add(strSoBK)
        ListViewDetail.Items(Indexlistview).SubItems.Add(strSHD)
        ListViewDetail.Items(Indexlistview).SubItems.Add(strTien)
        Indexlistview += 1
    End Sub
    Private Sub AddListView()
        ListViewDetail.Items.Clear()
        Dim i As Integer
        Dim strPT As String
        Dim strNgay As String
        Dim strNhanvien As String
        Dim strHD As String
        Dim strTien As String
        Dim strSoBK As String
        Dim strTenDaily As String
        Dim SoTTDL As Long
        Dim STTDL As Long

        For i = 0 To ds.Tables("ListReceipts").Rows.Count - 1

            strTenDaily = ds.Tables("ListReceipts").Rows(i).Item("TenDaily")

            AddSubItemListView(strTenDaily, "", "", "", "", "")
            ListViewDetail.Items(Indexlistview - 1).ForeColor = Color.Blue
            ListViewDetail.Items(Indexlistview - 1).Font = New Font("Times New Roman", 12, FontStyle.Bold)
            ListViewDetail.Items(Indexlistview - 1).BackColor = Color.Cornsilk

            SoTTDL = ds.Tables("ListReceipts").Rows(i).Item("STTDL")
            STTDL = SoTTDL
            strPT = " Phiếu thu số " & ds.Tables("ListReceipts").Rows(i).Item("Ordinal_No")
            strNgay = CDate(ds.Tables("ListReceipts").Rows(i).Item("Receipt_Date")).ToShortDateString
            strNhanvien = ds.Tables("ListReceipts").Rows(i).Item("Employ_Code")
            strHD = ds.Tables("ListReceipts").Rows(i).Item("Invoice_Quantity")
            strTien = ds.Tables("ListReceipts").Rows(i).Item("Total_Money")
            strSoBK = ds.Tables("ListReceipts").Rows(i).Item("List_Quantity")
            AddSubItemListView(strPT, strNgay, strNhanvien, strHD, strTien, strSoBK)
            i += 1
            While (STTDL = SoTTDL) AndAlso (i <= ds.Tables("ListReceipts").Rows.Count - 1)
                strPT = " Phiếu thu số " & ds.Tables("ListReceipts").Rows(i).Item("Ordinal_No")
                strNgay = CDate(ds.Tables("ListReceipts").Rows(i).Item("Receipt_Date")).ToShortDateString
                strNhanvien = ds.Tables("ListReceipts").Rows(i).Item("Employ_Code")
                strHD = ds.Tables("ListReceipts").Rows(i).Item("Invoice_Quantity")
                strTien = ds.Tables("ListReceipts").Rows(i).Item("Total_Money")
                strSoBK = ds.Tables("ListReceipts").Rows(i).Item("List_Quantity")
                AddSubItemListView(strPT, strNgay, strNhanvien, strHD, strTien, strSoBK)
                i += 1
                If (i <= ds.Tables("ListReceipts").Rows.Count - 1) Then
                    SoTTDL = ds.Tables("ListReceipts").Rows(i).Item("STTDL")
                End If
            End While

            'Lay tien phieu chi
            Dim strQuery As String
            strQuery = "SELECT List_Quantity,Ordinal_No_List,ID, Ordinal_No, Expense_Date, Service_Code, Total_Money, Employ_Code, Account_No, Bank_Code, Pay_Date, Pay_No, NguoiNop FROM Tbl_Expenses WHERE STTDL = " & STTDL
            FillDatasetEx(strQuery)
            Dim count As Integer
            For count = 0 To ds1.Tables("ListExpenes").Rows.Count - 1
                strPT = " Phiếu chi số " & ds1.Tables("ListExpenes").Rows(count).Item("Ordinal_No")
                strNgay = CDate(ds1.Tables("ListExpenes").Rows(count).Item("Expense_Date")).ToShortDateString
                strNhanvien = ds1.Tables("ListExpenes").Rows(count).Item("Employ_Code")
                strHD = "0" 'ds.Tables("ListExpenes").Rows(i).Item("0")
                strTien = ds1.Tables("ListExpenes").Rows(count).Item("Total_Money")
                strSoBK = ds1.Tables("ListExpenes").Rows(count).Item("List_Quantity")
                AddSubItemListView(strPT, strNgay, strNhanvien, strHD, strTien, strSoBK)
                txttongtien.Text = CLng(txttongtien.Text) - CLng(strTien)
            Next
            i -= 1
        Next
    End Sub

    Private Sub FillReports(ByVal strQuery As String, ByVal strTableName As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(Dsrpt, strTableName)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub AddTableDetail()
        Dim Newrow As DataRow
        Dim dt As DataTable
        Dim value As Boolean
        Dim strEmployee_Code As String
        Dim i, count As Integer
        Dim lvitem As ListViewItem
        Dim vlue As Integer
        Select Case Cbolydo.Text
            Case "PSTNDL"
                count = 0
                For i = 0 To ListViewDetail.Items.Count - 1
                    Newrow = Dsrpt.Tables("QueryReceipts").NewRow
                    Dsrpt.Tables("QueryReceipts").Rows.Add(Newrow)
                    lvitem = ListViewDetail.Items(i)
                    If (Trim$(lvitem.SubItems(1).Text) = "") Then
                        Dim litem As ListViewItem
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("FillName") = lvitem.SubItems(0).Text
                        litem = ListViewDetail.Items(i + 1)
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("Employ_Code") = litem.SubItems(2).Text
                    Else
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("FillName") = lvitem.SubItems(0).Text
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("Employ_Code") = lvitem.SubItems(2).Text
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("List_Quantity") = lvitem.SubItems(3).Text
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("Invoice_Quantity") = lvitem.SubItems(4).Text

                        vlue = lvitem.SubItems(0).Text.IndexOf("chi")
                        If (vlue < 1) Then
                            Dsrpt.Tables("QueryReceipts").Rows(count).Item("Total_Money") = lvitem.SubItems(5).Text
                        Else
                            Dsrpt.Tables("QueryReceipts").Rows(count).Item("Total_Money") = -CDbl(lvitem.SubItems(5).Text)
                        End If

                    End If
                    count += 1
                Next

            Case "PSTNDS", "PSTNTNDN"

            Case Else
                dt = DataGridListReceipts.DataSource
                count = 0
                For i = 0 To dt.Rows.Count - 1
                    value = dt.Rows(i).Item("Check")
                    If (value) Then
                        Newrow = Dsrpt.Tables("QueryReceipts").NewRow
                        Dsrpt.Tables("QueryReceipts").Rows.Add(Newrow)
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("Ordinal_No") = dt.Rows(i).Item("Ordinal_No")
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("List_Quantity") = dt.Rows(i).Item("List_Quantity")
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("Invoice_Quantity") = dt.Rows(i).Item("Invoice_Quantity")
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("Total_Money") = dt.Rows(i).Item("Total_Money")
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("Employ_Code") = dt.Rows(i).Item("Employ_Code")
                        Dsrpt.Tables("QueryReceipts").Rows(count).Item("FillName") = " Phiếu thu số"
                        count += 1
                    End If
                Next
        End Select

    End Sub

    Private Sub DateTimePickerNgayPC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerNgayPC.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtsobknt.Focus()
        End If
    End Sub


    Private Sub txtsobknt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsobknt.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            CboEmploy_code.Focus()
        End If
    End Sub

    Private Sub FormatDataGrid(ByVal dt As DataTable)
        With DataGridListReceipts
            .AllowNavigation = False
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách phiếu thu ...."
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

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

                .Add(New CGridCheckBoxStyle("Check", 40, HorizontalAlignment.Center, False, "In", "", False, True, False, False))

                .Add(New CGridTextBoxStyle("Ordinal_No", 50, HorizontalAlignment.Center, True, "Số PT", String.Empty, ""))

                .Add(New CGridDateTimePickerStyle("Receipt_Date", 100, True, "    Ngày thu", DateTimePickerFormat.Custom, "dd/MM/yyyy", "dd/MM/yyyy"))

                .Add(New CGridTextBoxStyle("Employ_Code", 220, HorizontalAlignment.Left, True, "                 Nhân viên ", String.Empty, ""))

                .Add(New CGridTextBoxStyle("Invoice_Quantity", 80, HorizontalAlignment.Center, True, "SL HĐ", String.Empty, ""))

                .Add(New CGridTextBoxStyle("Total_Money", 100, HorizontalAlignment.Right, True, "Tiền", String.Empty, ""))

                .Add(New CGridTextBoxStyle("List_Quantity", 80, HorizontalAlignment.Center, True, "SL BK", String.Empty, ""))
            End With
        End With
        CGrid.SetGridStyle(Me.DataGridListReceipts, dt, TblStyle)
        CGrid.DisableAddNew(DataGridListReceipts, Me)
    End Sub

    Private Sub FillDataGrid()
        CGrid.ClearTableStyles(DataGridListReceipts)
        FormatDataGrid(ds.Tables("ListReceipts"))
    End Sub

    Private Sub DataGridListReceipts_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridListReceipts.MouseUp
        Dim ClickedRowIndex As Integer
        Dim bChecked As Boolean
        Dim ClickedColumnName As String
        Dim result As Object = Nothing

        ClickedRowIndex = CGrid.GetClickedCellAndRow(CType(DataGridListReceipts.DataSource, DataTable), Me.DataGridListReceipts, ClickedColumnName, result, False)
        If ClickedRowIndex > -1 AndAlso ClickedColumnName = "Check" Then
            ClickedRowIndex = CGrid.SelectCheckBoxRow(CType(DataGridListReceipts.DataSource, DataTable), Me.DataGridListReceipts, e, "Check", bChecked, 0, True)
            result = bChecked
        End If

        If ClickedRowIndex > -1 Then
            If Not result Is Nothing Then
                txtTongBK.Text = SumColum(DataGridListReceipts, "List_Quantity")
                txttongtien.Text = SumColum(DataGridListReceipts, "Total_Money")
                splitn.strnumbers = CStr(CLng(txttongtien.Text))
                txttongtien.Text = splitn.Splitnumer(",")
                txtTongHD.Text = SumColum(DataGridListReceipts, "Invoice_Quantity")
                strListReceiptID = BillStrListReceipts(DataGridListReceipts, "ID")
                strListReceiptno = BillStrListReceipts(DataGridListReceipts, "Ordinal_No")
                txttongsophieu.Text = CountSoPhieu(DataGridListReceipts)
            End If
        End If

    End Sub

    Private Sub DataGridListReceipts_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridListReceipts.KeyUp
        If (e.KeyCode = Keys.Space) Then

            Dim ClickedRowIndex As Integer
            Dim bChecked As Boolean
            Dim ClickedColumnName As String
            Dim result As Object = Nothing

            ClickedRowIndex = CGrid.GetClickedCellAndRow(CType(DataGridListReceipts.DataSource, DataTable), Me.DataGridListReceipts, ClickedColumnName, result, False)
            If ClickedRowIndex > -1 AndAlso ClickedColumnName = "Check" Then
                ClickedRowIndex = CGrid.SelectCheckBoxRow(CType(DataGridListReceipts.DataSource, DataTable), Me.DataGridListReceipts, "Check", bChecked, True)
                result = bChecked
            End If

            If ClickedRowIndex > -1 Then
                If Not result Is Nothing Then
                    txtTongBK.Text = SumColum(DataGridListReceipts, "List_Quantity")
                    txttongtien.Text = SumColum(DataGridListReceipts, "Total_Money")
                    splitn.strnumbers = CStr(CLng(txttongtien.Text))
                    txttongtien.Text = splitn.Splitnumer(",")
                    txtTongHD.Text = SumColum(DataGridListReceipts, "Invoice_Quantity")
                    strListReceiptID = BillStrListReceipts(DataGridListReceipts, "ID")
                    strListReceiptno = BillStrListReceipts(DataGridListReceipts, "Ordinal_No")
                    txttongsophieu.Text = CountSoPhieu(DataGridListReceipts)
                End If
            End If
        End If
    End Sub

    Private Sub ListViewDetail_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewDetail.SelectedIndexChanged

    End Sub
End Class
