Imports System.Data.OleDb
Public Class frmEmployee
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdNewEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdEditEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdSaveEmployee As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtsodt As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePickerngaysinh As System.Windows.Forms.DateTimePicker
    Friend WithEvents txttiendatcoc As System.Windows.Forms.TextBox
    Friend WithEvents txthinhthuc As System.Windows.Forms.TextBox
    Friend WithEvents txtdiachitamtru As System.Windows.Forms.TextBox
    Friend WithEvents txtdiachitthuongtru As System.Windows.Forms.TextBox
    Friend WithEvents txtnoisinh As System.Windows.Forms.TextBox
    Friend WithEvents txttenctv As System.Windows.Forms.TextBox
    Friend WithEvents txtmactv As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtsocmnd As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxphai As System.Windows.Forms.CheckBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerNgayvaolam As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckBoxStatus As System.Windows.Forms.CheckBox
    Friend WithEvents cbotrinhdo As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEmployee))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CheckBoxStatus = New System.Windows.Forms.CheckBox
        Me.DateTimePickerNgayvaolam = New System.Windows.Forms.DateTimePicker
        Me.Label13 = New System.Windows.Forms.Label
        Me.CheckBoxphai = New System.Windows.Forms.CheckBox
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtsodt = New System.Windows.Forms.TextBox
        Me.txtsocmnd = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.DateTimePickerngaysinh = New System.Windows.Forms.DateTimePicker
        Me.txttiendatcoc = New System.Windows.Forms.TextBox
        Me.txthinhthuc = New System.Windows.Forms.TextBox
        Me.txtdiachitamtru = New System.Windows.Forms.TextBox
        Me.txtdiachitthuongtru = New System.Windows.Forms.TextBox
        Me.txtnoisinh = New System.Windows.Forms.TextBox
        Me.txttenctv = New System.Windows.Forms.TextBox
        Me.txtmactv = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdSaveEmployee = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdEditEmployee = New System.Windows.Forms.Button
        Me.cmdNewEmployee = New System.Windows.Forms.Button
        Me.cbotrinhdo = New System.Windows.Forms.ComboBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbotrinhdo)
        Me.GroupBox1.Controls.Add(Me.CheckBoxStatus)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerNgayvaolam)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.CheckBoxphai)
        Me.GroupBox1.Controls.Add(Me.cbostations)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.txtsodt)
        Me.GroupBox1.Controls.Add(Me.txtsocmnd)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerngaysinh)
        Me.GroupBox1.Controls.Add(Me.txttiendatcoc)
        Me.GroupBox1.Controls.Add(Me.txthinhthuc)
        Me.GroupBox1.Controls.Add(Me.txtdiachitamtru)
        Me.GroupBox1.Controls.Add(Me.txtdiachitthuongtru)
        Me.GroupBox1.Controls.Add(Me.txtnoisinh)
        Me.GroupBox1.Controls.Add(Me.txttenctv)
        Me.GroupBox1.Controls.Add(Me.txtmactv)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 48)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(546, 304)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'CheckBoxStatus
        '
        Me.CheckBoxStatus.Checked = True
        Me.CheckBoxStatus.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxStatus.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxStatus.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBoxStatus.Location = New System.Drawing.Point(298, 272)
        Me.CheckBoxStatus.Name = "CheckBoxStatus"
        Me.CheckBoxStatus.Size = New System.Drawing.Size(238, 22)
        Me.CheckBoxStatus.TabIndex = 15
        Me.CheckBoxStatus.Text = "Tình trạng hiện tại"
        '
        'DateTimePickerNgayvaolam
        '
        Me.DateTimePickerNgayvaolam.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerNgayvaolam.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerNgayvaolam.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerNgayvaolam.Location = New System.Drawing.Point(112, 206)
        Me.DateTimePickerNgayvaolam.Name = "DateTimePickerNgayvaolam"
        Me.DateTimePickerNgayvaolam.Size = New System.Drawing.Size(168, 26)
        Me.DateTimePickerNgayvaolam.TabIndex = 10
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label13.Location = New System.Drawing.Point(10, 208)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 24)
        Me.Label13.TabIndex = 13
        Me.Label13.Text = "Ngày vào làm "
        '
        'CheckBoxphai
        '
        Me.CheckBoxphai.Checked = True
        Me.CheckBoxphai.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxphai.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxphai.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBoxphai.Location = New System.Drawing.Point(296, 80)
        Me.CheckBoxphai.Name = "CheckBoxphai"
        Me.CheckBoxphai.Size = New System.Drawing.Size(88, 16)
        Me.CheckBoxphai.TabIndex = 5
        Me.CheckBoxphai.Text = "Nam/Nữ"
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(376, 14)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(160, 27)
        Me.cbostations.TabIndex = 2
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label12.Location = New System.Drawing.Point(296, 15)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(64, 24)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "Tổ thu"
        '
        'txtsodt
        '
        Me.txtsodt.BackColor = System.Drawing.Color.White
        Me.txtsodt.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsodt.Location = New System.Drawing.Point(376, 206)
        Me.txtsodt.Name = "txtsodt"
        Me.txtsodt.Size = New System.Drawing.Size(160, 26)
        Me.txtsodt.TabIndex = 11
        Me.txtsodt.Text = ""
        Me.txtsodt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtsocmnd
        '
        Me.txtsocmnd.BackColor = System.Drawing.Color.White
        Me.txtsocmnd.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsocmnd.ForeColor = System.Drawing.Color.Blue
        Me.txtsocmnd.Location = New System.Drawing.Point(376, 112)
        Me.txtsocmnd.Name = "txtsocmnd"
        Me.txtsocmnd.Size = New System.Drawing.Size(160, 26)
        Me.txtsocmnd.TabIndex = 7
        Me.txtsocmnd.Text = ""
        Me.txtsocmnd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label10.Location = New System.Drawing.Point(296, 210)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 24)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "Số ĐT "
        '
        'DateTimePickerngaysinh
        '
        Me.DateTimePickerngaysinh.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngaysinh.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerngaysinh.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngaysinh.Location = New System.Drawing.Point(112, 78)
        Me.DateTimePickerngaysinh.Name = "DateTimePickerngaysinh"
        Me.DateTimePickerngaysinh.Size = New System.Drawing.Size(168, 26)
        Me.DateTimePickerngaysinh.TabIndex = 4
        '
        'txttiendatcoc
        '
        Me.txttiendatcoc.BackColor = System.Drawing.Color.White
        Me.txttiendatcoc.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttiendatcoc.ForeColor = System.Drawing.Color.Firebrick
        Me.txttiendatcoc.Location = New System.Drawing.Point(112, 269)
        Me.txttiendatcoc.Name = "txttiendatcoc"
        Me.txttiendatcoc.Size = New System.Drawing.Size(168, 26)
        Me.txttiendatcoc.TabIndex = 14
        Me.txttiendatcoc.Text = "0"
        Me.txttiendatcoc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txthinhthuc
        '
        Me.txthinhthuc.BackColor = System.Drawing.Color.White
        Me.txthinhthuc.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txthinhthuc.Location = New System.Drawing.Point(376, 238)
        Me.txthinhthuc.Name = "txthinhthuc"
        Me.txthinhthuc.Size = New System.Drawing.Size(160, 26)
        Me.txthinhthuc.TabIndex = 13
        Me.txthinhthuc.Text = ""
        '
        'txtdiachitamtru
        '
        Me.txtdiachitamtru.BackColor = System.Drawing.Color.White
        Me.txtdiachitamtru.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdiachitamtru.Location = New System.Drawing.Point(112, 174)
        Me.txtdiachitamtru.Name = "txtdiachitamtru"
        Me.txtdiachitamtru.Size = New System.Drawing.Size(424, 26)
        Me.txtdiachitamtru.TabIndex = 9
        Me.txtdiachitamtru.Text = ""
        '
        'txtdiachitthuongtru
        '
        Me.txtdiachitthuongtru.BackColor = System.Drawing.Color.White
        Me.txtdiachitthuongtru.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdiachitthuongtru.Location = New System.Drawing.Point(112, 142)
        Me.txtdiachitthuongtru.Name = "txtdiachitthuongtru"
        Me.txtdiachitthuongtru.Size = New System.Drawing.Size(424, 26)
        Me.txtdiachitthuongtru.TabIndex = 8
        Me.txtdiachitthuongtru.Text = ""
        '
        'txtnoisinh
        '
        Me.txtnoisinh.BackColor = System.Drawing.Color.White
        Me.txtnoisinh.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnoisinh.Location = New System.Drawing.Point(112, 110)
        Me.txtnoisinh.Name = "txtnoisinh"
        Me.txtnoisinh.Size = New System.Drawing.Size(168, 26)
        Me.txtnoisinh.TabIndex = 6
        Me.txtnoisinh.Text = ""
        '
        'txttenctv
        '
        Me.txttenctv.BackColor = System.Drawing.Color.White
        Me.txttenctv.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttenctv.Location = New System.Drawing.Point(112, 46)
        Me.txttenctv.Name = "txttenctv"
        Me.txttenctv.Size = New System.Drawing.Size(424, 26)
        Me.txttenctv.TabIndex = 3
        Me.txttenctv.Text = ""
        '
        'txtmactv
        '
        Me.txtmactv.BackColor = System.Drawing.Color.White
        Me.txtmactv.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmactv.ForeColor = System.Drawing.Color.Blue
        Me.txtmactv.Location = New System.Drawing.Point(112, 14)
        Me.txtmactv.Name = "txtmactv"
        Me.txtmactv.Size = New System.Drawing.Size(168, 26)
        Me.txtmactv.TabIndex = 1
        Me.txtmactv.Text = ""
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(10, 270)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(88, 24)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Tiền đặt cọc"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label7.Location = New System.Drawing.Point(296, 238)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(85, 24)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "HT đặt cọc"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(10, 175)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(109, 24)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Nơi ở hiện nay "
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(10, 143)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(109, 24)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "HK thường trú "
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(10, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 24)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Nơi sinh"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(10, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 24)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Năm sinh "
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(10, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(91, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Họ tên CTV "
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(10, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Mã CTV"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label11.Location = New System.Drawing.Point(296, 112)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(88, 24)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Số CMND "
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label14.Location = New System.Drawing.Point(10, 238)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(64, 24)
        Me.Label14.TabIndex = 9
        Me.Label14.Text = "Trình độ"
        '
        'Label9
        '
        Me.Label9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label9.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 24.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label9.Location = New System.Drawing.Point(6, 2)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(545, 46)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "THÔNG TIN CỘNG TÁC VIÊN"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdSaveEmployee)
        Me.GroupBox2.Controls.Add(Me.cmdClose)
        Me.GroupBox2.Controls.Add(Me.cmdEditEmployee)
        Me.GroupBox2.Controls.Add(Me.cmdNewEmployee)
        Me.GroupBox2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox2.Location = New System.Drawing.Point(6, 346)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(544, 54)
        Me.GroupBox2.TabIndex = 15
        Me.GroupBox2.TabStop = False
        '
        'cmdSaveEmployee
        '
        Me.cmdSaveEmployee.Location = New System.Drawing.Point(224, 14)
        Me.cmdSaveEmployee.Name = "cmdSaveEmployee"
        Me.cmdSaveEmployee.Size = New System.Drawing.Size(88, 32)
        Me.cmdSaveEmployee.TabIndex = 16
        Me.cmdSaveEmployee.Text = "Lưu"
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(440, 14)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(88, 32)
        Me.cmdClose.TabIndex = 13
        Me.cmdClose.Text = "Đóng"
        '
        'cmdEditEmployee
        '
        Me.cmdEditEmployee.Location = New System.Drawing.Point(120, 14)
        Me.cmdEditEmployee.Name = "cmdEditEmployee"
        Me.cmdEditEmployee.Size = New System.Drawing.Size(88, 32)
        Me.cmdEditEmployee.TabIndex = 13
        Me.cmdEditEmployee.Text = "Thay Đổi"
        '
        'cmdNewEmployee
        '
        Me.cmdNewEmployee.Location = New System.Drawing.Point(16, 14)
        Me.cmdNewEmployee.Name = "cmdNewEmployee"
        Me.cmdNewEmployee.Size = New System.Drawing.Size(88, 32)
        Me.cmdNewEmployee.TabIndex = 17
        Me.cmdNewEmployee.Text = "Nhập Mới"
        '
        'cbotrinhdo
        '
        Me.cbotrinhdo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbotrinhdo.Location = New System.Drawing.Point(112, 237)
        Me.cbotrinhdo.Name = "cbotrinhdo"
        Me.cbotrinhdo.Size = New System.Drawing.Size(168, 27)
        Me.cbotrinhdo.TabIndex = 16
        '
        'frmEmployee
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(554, 402)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmEmployee"
        Me.Text = "Nhập cộng tác viên"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
    Public Sub LoadDataSet()
        mydataset = New DataSet
        strSQL = "SELECT StationID,Station_Name FROM Tbl_Stations "
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_Stations")
            cbostations.DataSource = mydataset.Tables("Tbl_Stations").DefaultView
            cbostations.DisplayMember = "Station_Name"
            cbostations.ValueMember = "StationID"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        strSQL = "SELECT QualificationCode,QualificationName FROM Tbl_Qualifications "
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_Qualifications")
            cbotrinhdo.DataSource = mydataset.Tables("Tbl_Qualifications").DefaultView
            cbotrinhdo.DisplayMember = "QualificationName"
            cbotrinhdo.ValueMember = "QualificationCode"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Private Sub DeleteTextbox()
        txtmactv.Text = vbNullString
        txttenctv.Text = vbNullString
        txtnoisinh.Text = vbNullString
        txtsocmnd.Text = vbNullString
        txtsodt.Text = vbNullString
        txtdiachitthuongtru.Text = vbNullString
        txtdiachitamtru.Text = vbNullString
        txthinhthuc.Text = vbNullString
        txttiendatcoc.Text = vbNullString

    End Sub

    Private Sub LockTextbox()
        txtmactv.ReadOnly = True
        txttenctv.ReadOnly = True
        txtnoisinh.ReadOnly = True
        txtsocmnd.ReadOnly = True
        txtsodt.ReadOnly = True
        txtdiachitthuongtru.ReadOnly = True
        txtdiachitamtru.ReadOnly = True
        txthinhthuc.ReadOnly = True
        txttiendatcoc.ReadOnly = True
    End Sub

    Private Sub UnLockTextbox()
        txtmactv.ReadOnly = False
        txttenctv.ReadOnly = False
        txtnoisinh.ReadOnly = False
        txtsocmnd.ReadOnly = False
        txtsodt.ReadOnly = False
        txtdiachitthuongtru.ReadOnly = False
        txtdiachitamtru.ReadOnly = False
        txthinhthuc.ReadOnly = False
        txttiendatcoc.ReadOnly = False
    End Sub

    Private Function CheckEmploy_Code(ByVal strMaCTV As String) As Boolean
        Dim result As Boolean = False
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = " SELECT Employ_Code FROM tbl_Employee WHERE Employ_Code = '" & strMaCTV & "'"
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then
                result = True
            End If
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :" & ex.ToString)
        End Try
        oledbcon.Close()
        Return result
    End Function
    Private Function CheckInfo() As Boolean
        Dim CheckValue As Boolean = True
        If (cmdSaveEmployee.Text = "Lưu") Then
            If Trim$(txtmactv.Text) = "" Then
                MsgBox("Bạn chưa nhập mã CTV! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
                txtmactv.Focus()
                CheckValue = False
                GoTo EndFunction
            End If

            If (CheckEmploy_Code(Trim$(txtmactv.Text))) Then
                MsgBox("Mã CTV này đã tồn tại! Vui lòng xem lại ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
                txtmactv.Focus()
                CheckValue = False
                GoTo EndFunction
            End If
        End If

        If Trim$(cbostations.Text) = "" Then
            MsgBox("Bạn chưa chọn tổ thu cước! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            cbostations.Focus()
            CheckValue = False
            GoTo EndFunction
        End If

        If Trim$(txttenctv.Text) = "" Then
            MsgBox("Bạn chưa nhập tên CTV! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            txttenctv.Focus()
            CheckValue = False
            GoTo EndFunction
        End If


        If Trim$(txtsocmnd.Text) = "" Then
            MsgBox("Bạn chưa nhập số CMND CTV! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            txtsocmnd.Focus()
            CheckValue = False
            GoTo EndFunction
        End If

        If Trim$(txtnoisinh.Text) = "" Then
            MsgBox("Bạn chưa nhập nơi sinh CTV! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            txtnoisinh.Focus()
            CheckValue = False
            GoTo EndFunction
        End If

        If Trim$(txtdiachitthuongtru.Text) = "" Then
            MsgBox("Bạn chưa nhập địa chỉ thường trú CTV! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            txtdiachitthuongtru.Focus()
            CheckValue = False
            GoTo EndFunction
        End If

        If Trim$(txthinhthuc.Text) = "" Then
            MsgBox("Bạn chưa nhập hình thức đặt cọc! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            txthinhthuc.Focus()
            CheckValue = False
            GoTo EndFunction
        End If


        If Trim$(txttiendatcoc.Text) = "" Then
            MsgBox("Bạn chưa nhập tiền đặt cọc! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            txttiendatcoc.Focus()
            CheckValue = False
            GoTo EndFunction
        End If

        If Not IsNumeric(Trim$(txttiendatcoc.Text)) Then
            MsgBox("Tiền đặt cọc phải là số! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
            txttiendatcoc.Focus()
            txttiendatcoc.SelectAll()
            CheckValue = False
            GoTo EndFunction
        End If
EndFunction:
        Return CheckValue
    End Function
    Private Sub SaveEmployee()
        strSQL = " INSERT INTO tbl_Employee(Employ_Code,Employ_Name,Identity_Card,Address1,Address2,DateOfBirth,PlaceOfBirth,Phone_No,form_Deposit,Deposit,StationID,Sex,Qualification,DateOfWork,Status) " & _
        " VALUES('" & Trim$(txtmactv.Text) & _
        "','" & Trim$(txttenctv.Text) & "','" & _
        Trim$(txtsocmnd.Text) & "','" & _
        txtdiachitthuongtru.Text & "','" & _
        txtdiachitamtru.Text & _
        "','" & DateTimePickerngaysinh.Value.ToShortDateString & _
        "','" & txtnoisinh.Text & _
        "','" & txtsodt.Text & _
        "','" & txthinhthuc.Text & _
        "'," & CLng(txttiendatcoc.Text) & _
        ",'" & cbostations.SelectedValue & _
        "'," & CheckBoxphai.Checked & _
        ",'" & cbotrinhdo.SelectedValue & _
        "','" & DateTimePickerNgayvaolam.Value.ToShortDateString & _
        "'," & CheckBoxStatus.Checked & _
        ")"
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strSQL
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
            MsgBox("Đã lưu vào hệ thống!!")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        oledbcon.Close()
    End Sub

    Private Sub UpdateEmployee()

        strSQL = " UPDATE tbl_Employee SET Employ_Name ='" & Trim$(txttenctv.Text) & _
         "',Identity_Card ='" & Trim$(txtsocmnd.Text) & _
         "',Address1 = '" & txtdiachitthuongtru.Text & _
         "',Address2 ='" & txtdiachitamtru.Text & _
         "',DateOfBirth ='" & DateTimePickerngaysinh.Value.ToShortDateString & _
         "',PlaceOfBirth ='" & txtnoisinh.Text & _
         "',Phone_No ='" & txtsodt.Text & _
         "',form_Deposit ='" & txthinhthuc.Text & _
         "',Deposit =" & CLng(txttiendatcoc.Text) & _
         ",StationID = '" & cbostations.SelectedValue & _
         "',Sex = " & CheckBoxphai.Checked & _
         ",Qualification = '" & cbotrinhdo.SelectedValue & _
         "',DateOfWork ='" & DateTimePickerNgayvaolam.Value.ToShortDateString & _
         "',Status = " & CheckBoxStatus.Checked & _
         " WHERE Employ_Code = '" & Trim$(txtmactv.Text) & "'"
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strSQL
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
            MsgBox("Đã cập nhật vào hệ thống!!")
        Catch ex As Exception
            MsgBox(" Lổi rồi người ơi!!!" & ex.ToString)
        End Try
        oledbcon.Close()
    End Sub
    Private Sub LoadInfo(ByVal strMaCTV As String)
        Dim value
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = " SELECT Employ_Name,Identity_Card,Address1,Address2,DateOfBirth,PlaceOfBirth,Phone_No,form_Deposit,Deposit,StationID,Sex,Qualification,DateOfWork,Status FROM tbl_Employee WHERE Employ_Code = '" & strMaCTV & "'"
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then
                If Not oleread.IsDBNull(0) Then
                    txttenctv.Text = oleread.GetString(0)
                End If

                If Not oleread.IsDBNull(1) Then
                    txtsocmnd.Text = oleread.GetString(1)
                End If

                If Not oleread.IsDBNull(2) Then
                    txtdiachitthuongtru.Text = oleread.GetString(2)
                End If

                If Not oleread.IsDBNull(3) Then
                    txtdiachitamtru.Text = oleread.GetString(3)
                End If

                If Not oleread.IsDBNull(4) Then
                    DateTimePickerngaysinh.Value = oleread.GetDateTime(4)
                End If

                If Not oleread.IsDBNull(5) Then
                    txtnoisinh.Text = oleread.GetString(5)
                End If

                If Not oleread.IsDBNull(6) Then
                    txtsodt.Text = oleread.GetString(6)
                End If

                If Not oleread.IsDBNull(7) Then
                    txthinhthuc.Text = oleread.GetString(7)
                End If

                If Not oleread.IsDBNull(8) Then
                    txttiendatcoc.Text = oleread.GetValue(8)
                End If

                If Not oleread.IsDBNull(9) Then
                    cbostations.SelectedIndex = cbostations.FindString(GetStringName("Tbl_Stations", oleread.GetString(9)))
                End If

                If Not oleread.IsDBNull(10) Then
                    CheckBoxphai.Checked = oleread.GetBoolean(10)
                End If

                If Not oleread.IsDBNull(11) Then
                    cbotrinhdo.SelectedIndex = cbotrinhdo.FindString(GetStringName("Tbl_Qualifications", oleread.GetString(11)))
                End If

                If Not oleread.IsDBNull(12) Then
                    DateTimePickerNgayvaolam.Value = oleread.GetDateTime(12)
                End If

                If Not oleread.IsDBNull(13) Then
                    CheckBoxStatus.Checked = oleread.GetBoolean(13)
                End If

            End If
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :" & ex.ToString)
        End Try
        oledbcon.Close()
    End Sub

    Private Sub cmdSaveEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveEmployee.Click
        If (CheckInfo()) Then
            If (cmdSaveEmployee.Text = "Lưu") Then
                SaveEmployee()
            Else
                UpdateEmployee()
            End If
            cmdNewEmployee.Focus()
        End If
    End Sub

    Private Sub txtmactv_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmactv.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If Trim$(txtmactv.Text) = "" Then
                MsgBox("Bạn chưa nhập mã CTV! Vui lòng nhập ...", MsgBoxStyle.Critical, "Lổi nhập liệu!")
                txtmactv.Focus()
                Exit Sub
            End If
            txtmactv.Text = txtmactv.Text.ToUpper
            Dim value
            If (CheckEmploy_Code(Trim$(txtmactv.Text))) Then
                value = MsgBox("Mã cộng tác viên này đã tồn tại. Bạn có muốn lấy lại thông tin CTV này không?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Thông báo!")
                If (value = vbYes) Then
                    LockTextbox()
                    LoadInfo(Trim(txtmactv.Text))
                    cmdSaveEmployee.Enabled = False
                End If
            End If
            cbostations.Focus()
        End If
    End Sub

    Private Sub cmdNewEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewEmployee.Click
        cmdSaveEmployee.Enabled = True
        cmdSaveEmployee.Text = "Lưu"
        UnLockTextbox()
        DeleteTextbox()
        txtmactv.Focus()
    End Sub

    Private Sub frmEmployee_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDataSet()
        cmdSaveEmployee.Enabled = False
        LockTextbox()
        DateTimePickerngaysinh.Value = Now
        DateTimePickerngaysinh.Value = DateTimePickerngaysinh.Value.AddYears(-18)
    End Sub

    Private Sub cmdEditEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditEmployee.Click
        cmdSaveEmployee.Enabled = True
        cmdSaveEmployee.Text = "Cập Nhật"
        UnLockTextbox()
        txtmactv.ReadOnly = True
    End Sub

    Private Sub DateTimePickerngaysinh_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerngaysinh.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            CheckBoxphai.Focus()
        End If
    End Sub

    Private Sub txtdiachitamtru_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdiachitamtru.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerNgayvaolam.Focus()
        End If
    End Sub

    Private Sub txtdiachitthuongtru_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdiachitthuongtru.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtdiachitamtru.Focus()
        End If
    End Sub

    Private Sub txthinhthuc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txthinhthuc.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttiendatcoc.Focus()
        End If
    End Sub

    Private Sub txtnoisinh_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtnoisinh.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtsocmnd.Focus()
        End If
    End Sub

    Private Sub txtsocmnd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsocmnd.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtdiachitthuongtru.Focus()
        End If
    End Sub

    Private Sub txtsodt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsodt.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cbotrinhdo.Focus()
        End If
    End Sub

    Private Sub txttenctv_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttenctv.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttenctv.Text = txttenctv.Text.ToUpper
            DateTimePickerngaysinh.Focus()
            'KeyAscii = 0
            'System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txttiendatcoc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttiendatcoc.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdSaveEmployee.Focus()
        End If
    End Sub

    Private Sub DateTimePickerngaysinh_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePickerngaysinh.ValueChanged
        If DateTimePickerngaysinh.Value > Now Then
            DateTimePickerngaysinh.Value = Now
        End If
    End Sub
    Private Function GetStringName(ByVal strTablename As String, ByVal strCode As String) As String
        Dim i As Integer
        Dim strresult As String
        For i = 0 To mydataset.Tables(strTablename).Rows.Count - 1
            strresult = mydataset.Tables(strTablename).Rows(i).Item(0)
            If (strresult.Equals(strCode)) Then
                strresult = mydataset.Tables(strTablename).Rows(i).Item(1)
                Exit For
            End If
        Next
        Return strresult
    End Function

    Private Sub cbostations_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbostations.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttenctv.Focus()
        End If
    End Sub

    Private Sub CheckBoxphai_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CheckBoxphai.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtnoisinh.Focus()
        End If
    End Sub

    Private Sub DateTimePickerNgayvaolam_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerNgayvaolam.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtsodt.Focus()
        End If
    End Sub

    Private Sub cbotrinhdo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbotrinhdo.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txthinhthuc.Focus()
        End If
    End Sub

    Private Sub CheckBoxStatus_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CheckBoxStatus.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdSaveEmployee.Focus()
        End If
    End Sub
End Class
