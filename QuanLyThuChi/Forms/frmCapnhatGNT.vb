Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmCapnhatphieuthu
    Inherits System.Windows.Forms.Form
    Private ds As DataSet
    Private mydataset As DataSet
    Dim start As Boolean = False
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        FillDataset()

        start = True
        If (cbostations.Items.Count > 0) Then
            cbostations.SelectedIndex = 0

            Try
                strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
                Try
                    CboEmploy_code.DataSource = Nothing
                    cbonguoinhan.DataSource = Nothing
                    Try
                        mydataset.Tables("Tbl_Employee").Clear()
                    Catch ex As Exception
                    End Try

                    CboEmploy_code.Items.Clear()
                    FillCombo(CboEmploy_code, strSQL, "Tbl_Employee", "Employ_Code", "Employ_Name")

                    cbonguoinhan.Items.Clear()
                    FillCombo(cbonguoinhan, strSQL, "cbonguoinhan", "Employ_Code", "Employ_Name")

                    txtEmployeeName.Text = CboEmploy_code.SelectedValue
                    If (cbonguoinhan.Items.Count > 0) Then
                        cbonguoinhan.SelectedIndex = 0
                        txtnguoinop.Text = cbonguoinhan.SelectedValue
                    End If

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
            End Try

        End If

        If (Cbolydo.Items.Count > 0) Then
            Cbolydo.SelectedIndex = 0
            txtlydo.Text = "Chi nộp " & Cbolydo.SelectedValue
        End If

        If (cboAccounts_Banks.Items.Count > 0) Then
            cboAccounts_Banks.SelectedIndex = 0
            txtBank_Name.Text = cboAccounts_Banks.SelectedValue
        End If

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
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents cmdLuu As System.Windows.Forms.Button
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtlydo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtso As System.Windows.Forms.TextBox
    Friend WithEvents txtSoBKNT As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lbltongDV As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents DataGridListExpenes As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmdlist As System.Windows.Forms.Button
    Friend WithEvents DateTimePickerTungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePickerDenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txttusoBK As System.Windows.Forms.TextBox
    Friend WithEvents txtDensoBK As System.Windows.Forms.TextBox
    Friend WithEvents txtnguoinop As System.Windows.Forms.TextBox
    Friend WithEvents cbonguoinhan As System.Windows.Forms.ComboBox
    Friend WithEvents cmdUpdateDatagrid As System.Windows.Forms.Button
    Friend WithEvents cboAccounts_Banks As System.Windows.Forms.ComboBox
    Friend WithEvents txtBank_Name As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePickerNgayNop As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label15 = New System.Windows.Forms.Label
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.cmdUpdateDatagrid = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.txtBank_Name = New System.Windows.Forms.TextBox
        Me.cboAccounts_Banks = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.cbonguoinhan = New System.Windows.Forms.ComboBox
        Me.txtnguoinop = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.DateTimePickerNgayNop = New System.Windows.Forms.DateTimePicker
        Me.Label10 = New System.Windows.Forms.Label
        Me.cmdLuu = New System.Windows.Forms.Button
        Me.cmddong = New System.Windows.Forms.Button
        Me.DataGridListExpenes = New System.Windows.Forms.DataGrid
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtlydo = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DateTimePickerTungay = New System.Windows.Forms.DateTimePicker
        Me.DateTimePickerDenngay = New System.Windows.Forms.DateTimePicker
        Me.txttusoBK = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDensoBK = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmdlist = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DataGridListExpenes, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.Label15.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label15.Location = New System.Drawing.Point(3, 4)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(712, 44)
        Me.Label15.TabIndex = 55
        Me.Label15.Text = "CẬP NHẬT GIẤY NỘP TIỀN                 "
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(522, 20)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(192, 27)
        Me.cbostations.TabIndex = 56
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.DataGridListExpenes)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtlydo)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(4, 45)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 460)
        Me.GroupBox1.TabIndex = 54
        Me.GroupBox1.TabStop = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cmdUpdateDatagrid)
        Me.GroupBox4.Controls.Add(Me.Button1)
        Me.GroupBox4.Controls.Add(Me.txtBank_Name)
        Me.GroupBox4.Controls.Add(Me.cboAccounts_Banks)
        Me.GroupBox4.Controls.Add(Me.Label12)
        Me.GroupBox4.Controls.Add(Me.cbonguoinhan)
        Me.GroupBox4.Controls.Add(Me.txtnguoinop)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Controls.Add(Me.DateTimePickerNgayNop)
        Me.GroupBox4.Controls.Add(Me.Label10)
        Me.GroupBox4.Controls.Add(Me.cmdLuu)
        Me.GroupBox4.Controls.Add(Me.cmddong)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 345)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(696, 111)
        Me.GroupBox4.TabIndex = 9
        Me.GroupBox4.TabStop = False
        '
        'cmdUpdateDatagrid
        '
        Me.cmdUpdateDatagrid.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateDatagrid.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdUpdateDatagrid.Location = New System.Drawing.Point(72, 78)
        Me.cmdUpdateDatagrid.Name = "cmdUpdateDatagrid"
        Me.cmdUpdateDatagrid.Size = New System.Drawing.Size(112, 28)
        Me.cmdUpdateDatagrid.TabIndex = 74
        Me.cmdUpdateDatagrid.Text = "Cập nhật bảng"
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Button1.Location = New System.Drawing.Point(336, 78)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 28)
        Me.Button1.TabIndex = 73
        Me.Button1.Text = "Sửa"
        '
        'txtBank_Name
        '
        Me.txtBank_Name.BackColor = System.Drawing.Color.White
        Me.txtBank_Name.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBank_Name.ForeColor = System.Drawing.Color.Blue
        Me.txtBank_Name.Location = New System.Drawing.Point(264, 46)
        Me.txtBank_Name.Name = "txtBank_Name"
        Me.txtBank_Name.Size = New System.Drawing.Size(424, 26)
        Me.txtBank_Name.TabIndex = 72
        Me.txtBank_Name.Text = ""
        '
        'cboAccounts_Banks
        '
        Me.cboAccounts_Banks.Location = New System.Drawing.Point(66, 46)
        Me.cboAccounts_Banks.Name = "cboAccounts_Banks"
        Me.cboAccounts_Banks.Size = New System.Drawing.Size(198, 27)
        Me.cboAccounts_Banks.TabIndex = 12
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(1, 46)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(49, 22)
        Me.Label12.TabIndex = 71
        Me.Label12.Text = "Số TK"
        '
        'cbonguoinhan
        '
        Me.cbonguoinhan.Location = New System.Drawing.Point(264, 15)
        Me.cbonguoinhan.Name = "cbonguoinhan"
        Me.cbonguoinhan.Size = New System.Drawing.Size(195, 27)
        Me.cbonguoinhan.TabIndex = 11
        '
        'txtnguoinop
        '
        Me.txtnguoinop.BackColor = System.Drawing.Color.White
        Me.txtnguoinop.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnguoinop.ForeColor = System.Drawing.Color.Blue
        Me.txtnguoinop.Location = New System.Drawing.Point(459, 15)
        Me.txtnguoinop.Name = "txtnguoinop"
        Me.txtnguoinop.Size = New System.Drawing.Size(232, 26)
        Me.txtnguoinop.TabIndex = 69
        Me.txtnguoinop.Text = ""
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(187, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(77, 22)
        Me.Label11.TabIndex = 68
        Me.Label11.Text = "Người nộp"
        '
        'DateTimePickerNgayNop
        '
        Me.DateTimePickerNgayNop.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerNgayNop.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerNgayNop.Location = New System.Drawing.Point(66, 15)
        Me.DateTimePickerNgayNop.Name = "DateTimePickerNgayNop"
        Me.DateTimePickerNgayNop.Size = New System.Drawing.Size(116, 26)
        Me.DateTimePickerNgayNop.TabIndex = 10
        Me.DateTimePickerNgayNop.Value = New Date(2005, 6, 29, 0, 0, 0, 0)
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label10.Location = New System.Drawing.Point(1, 17)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(70, 22)
        Me.Label10.TabIndex = 65
        Me.Label10.Text = "Ngày nộp"
        '
        'cmdLuu
        '
        Me.cmdLuu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuu.Location = New System.Drawing.Point(204, 78)
        Me.cmdLuu.Name = "cmdLuu"
        Me.cmdLuu.Size = New System.Drawing.Size(112, 28)
        Me.cmdLuu.TabIndex = 13
        Me.cmdLuu.Text = "Lưu"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(576, 78)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(112, 28)
        Me.cmddong.TabIndex = 31
        Me.cmddong.Text = "Đóng"
        '
        'DataGridListExpenes
        '
        Me.DataGridListExpenes.DataMember = ""
        Me.DataGridListExpenes.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridListExpenes.Location = New System.Drawing.Point(8, 138)
        Me.DataGridListExpenes.Name = "DataGridListExpenes"
        Me.DataGridListExpenes.Size = New System.Drawing.Size(696, 206)
        Me.DataGridListExpenes.TabIndex = 67
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Location = New System.Drawing.Point(77, 48)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(195, 27)
        Me.CboEmploy_code.TabIndex = 2
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(77, 16)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(192, 27)
        Me.Cbolydo.TabIndex = 1
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(272, 48)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(432, 26)
        Me.txtEmployeeName.TabIndex = 50
        Me.txtEmployeeName.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(84, 22)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Người nhận"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Dịch vụ"
        '
        'txtlydo
        '
        Me.txtlydo.BackColor = System.Drawing.Color.White
        Me.txtlydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtlydo.ForeColor = System.Drawing.Color.Black
        Me.txtlydo.Location = New System.Drawing.Point(272, 16)
        Me.txtlydo.Name = "txtlydo"
        Me.txtlydo.Size = New System.Drawing.Size(432, 26)
        Me.txtlydo.TabIndex = 50
        Me.txtlydo.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DateTimePickerTungay)
        Me.GroupBox2.Controls.Add(Me.DateTimePickerDenngay)
        Me.GroupBox2.Controls.Add(Me.txttusoBK)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtDensoBK)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.cmdlist)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 80)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(696, 56)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Thông tin bảng kê nộp tiền vào NH"
        '
        'DateTimePickerTungay
        '
        Me.DateTimePickerTungay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerTungay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerTungay.Location = New System.Drawing.Point(61, 19)
        Me.DateTimePickerTungay.Name = "DateTimePickerTungay"
        Me.DateTimePickerTungay.Size = New System.Drawing.Size(116, 26)
        Me.DateTimePickerTungay.TabIndex = 4
        Me.DateTimePickerTungay.Value = New Date(2005, 6, 29, 0, 0, 0, 0)
        '
        'DateTimePickerDenngay
        '
        Me.DateTimePickerDenngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerDenngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerDenngay.Location = New System.Drawing.Point(243, 19)
        Me.DateTimePickerDenngay.Name = "DateTimePickerDenngay"
        Me.DateTimePickerDenngay.Size = New System.Drawing.Size(116, 26)
        Me.DateTimePickerDenngay.TabIndex = 5
        Me.DateTimePickerDenngay.Value = New Date(2005, 6, 29, 0, 0, 0, 0)
        '
        'txttusoBK
        '
        Me.txttusoBK.BackColor = System.Drawing.Color.White
        Me.txttusoBK.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttusoBK.ForeColor = System.Drawing.Color.Black
        Me.txttusoBK.Location = New System.Drawing.Point(425, 19)
        Me.txttusoBK.Name = "txttusoBK"
        Me.txttusoBK.Size = New System.Drawing.Size(56, 26)
        Me.txttusoBK.TabIndex = 6
        Me.txttusoBK.Text = ""
        Me.txttusoBK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label7.Location = New System.Drawing.Point(359, 21)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(71, 22)
        Me.Label7.TabIndex = 63
        Me.Label7.Text = "Từ số BK"
        '
        'txtDensoBK
        '
        Me.txtDensoBK.BackColor = System.Drawing.Color.White
        Me.txtDensoBK.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDensoBK.ForeColor = System.Drawing.Color.Black
        Me.txtDensoBK.Location = New System.Drawing.Point(552, 19)
        Me.txtDensoBK.Name = "txtDensoBK"
        Me.txtDensoBK.Size = New System.Drawing.Size(56, 26)
        Me.txtDensoBK.TabIndex = 7
        Me.txtDensoBK.Text = ""
        Me.txtDensoBK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(480, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 22)
        Me.Label1.TabIndex = 62
        Me.Label1.Text = "Đến số BK"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(3, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 22)
        Me.Label3.TabIndex = 63
        Me.Label3.Text = "Từ ngày"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label9.Location = New System.Drawing.Point(177, 21)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(69, 22)
        Me.Label9.TabIndex = 62
        Me.Label9.Text = "Đến ngày"
        '
        'cmdlist
        '
        Me.cmdlist.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdlist.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdlist.Location = New System.Drawing.Point(611, 18)
        Me.cmdlist.Name = "cmdlist"
        Me.cmdlist.Size = New System.Drawing.Size(80, 28)
        Me.cmdlist.TabIndex = 8
        Me.cmdlist.Text = "Lên DS"
        '
        'frmCapnhatphieuthu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 520)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmCapnhatphieuthu"
        Me.Text = "Cập nhật phiếu chi nộp tiền vào ngân hàng"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.DataGridListExpenes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub cmddong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddong.Click
        Me.Close()
    End Sub

    Private Sub FormatDataGridListExpenes()

        With DataGridListExpenes
            .AllowNavigation = False
            .DataMember = "ListExpenes"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách phiếu chi ...."
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "ListExpenes"
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
                    .MappingName = "Ordinal_No_List"
                    .HeaderText = "Số BK   "
                    .Width = 80
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Ordinal_No"
                    .HeaderText = "Số PC"
                    .Width = 80
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
                '.Add(New DataGridDateTimePicker)
                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "Expense_Date"
                    .HeaderText = " Ngày chi"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "Employ_Code"
                    .HeaderText = "               Người nhận "
                    .Width = 220
                    .Alignment = HorizontalAlignment.Left
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
                    .MappingName = "Pay_No"
                    .HeaderText = "Số GNT"
                    .Width = 80
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = False
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(7)
                    .MappingName = "Pay_Date"
                    .HeaderText = "Ngày nộp"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(8)
                    .MappingName = "Account_No"
                    .HeaderText = "Số TK"
                    .Width = 80
                    .Alignment = HorizontalAlignment.Right
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(9)
                    .MappingName = "Bank_Code"
                    .HeaderText = "         Ngân hàng nộp"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(10)
                    .MappingName = "NguoiNop"
                    .HeaderText = "             Người nộp"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

            End With
        End With
        DataGridListExpenes.TableStyles.Add(TblStyle)
    End Sub

    Private Sub FillDataset(ByVal strQuery As String)
        Try
            ds = New DataSet
            DataGridListExpenes.DataSource = Nothing
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds, "ListExpenes")

            Dim myTypeCheck As System.Type

            myTypeCheck = System.Type.GetType("System.Boolean")
            ds.Tables("ListExpenes").Columns.Add(New System.Data.DataColumn("Check", myTypeCheck))


            Dim i As Integer
            For i = 0 To ds.Tables("ListExpenes").Rows.Count - 1
                ds.Tables("ListExpenes").Rows(i).Item("Check") = True
            Next
            DataGridListExpenes.DataSource = ds.Tables("ListExpenes")
        Catch ex As Exception
            MsgBox("Lổi View :" & ex.ToString)
        End Try
    End Sub

    Public Sub FillDataSet()
        mydataset = New DataSet
        strSQL = "SELECT StationID,Station_Name,Station_Address FROM Tbl_Stations "
        FillCombo(cbostations, strSQL, "Tbl_Stations", "Station_Name", "StationID")

        strSQL = "SELECT Service_Code,Service_Name FROM Tbl_Services "
        FillCombo(Cbolydo, strSQL, "Tbl_Services", "Service_Code", "Service_Name")

        strSQL = "SELECT Account_No,Tbl_Banks.Bank_Code,Bank_Name FROM Tbl_Banks,Tbl_Accounts_Banks WHERE Tbl_Accounts_Banks.Bank_Code = Tbl_Banks.Bank_Code "
        FillCombo(cboAccounts_Banks, strSQL, "Tbl_Accounts_Banks", "Account_No", "Bank_Name")

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
    Private Sub frmCapnhatphieuthu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormatDataGridListExpenes()
        strSQL = "SELECT Ordinal_No_List,ID, Ordinal_No, Expense_Date, Service_Code, Total_Money, Employ_Code, Account_No, Bank_Code, Pay_Date, Pay_No, NguoiNop FROM Tbl_Expenses WHERE Pay_No = -1 "
        FillDataset(strSQL)
    End Sub

    Private Sub cmdlist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdlist.Click
        CheckSQL()
        FillDataset(strSQL)
    End Sub

    Private Function CheckSQL() As Boolean
        Dim result As Boolean = True
        strSQL = "SELECT Ordinal_No_List,ID, Ordinal_No, Expense_Date, Service_Code, Total_Money, Employ_Code, Account_No, Bank_Code, Pay_Date, Pay_No, NguoiNop FROM Tbl_Expenses WHERE Expense_Date BETWEEN #" & DateTimePickerTungay.Value.ToShortDateString & " 00:00:00# AND #" & DateTimePickerDenngay.Value.ToShortDateString & " 23:59:59# AND Status = False "


        If (Trim$(Cbolydo.Text) <> "") Then
            strSQL += " AND Service_Code ='" & Cbolydo.Text & "'"
        Else
            MsgBox("Dịch vụ nộp chưa được chọn!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            Cbolydo.Focus()
            result = False
            GoTo endFunction
        End If

        If (Trim$(txttusoBK.Text) <> "") Then
            If (Not IsNumeric(Trim$(txttusoBK.Text))) Then
                MsgBox("Điều kiện từ số phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txttusoBK.Focus()
                txttusoBK.SelectAll()
                result = False
                GoTo endFunction
            End If

            If (CLng(Trim$(txttusoBK.Text)) < 1) Then
                MsgBox("Điều kiện từ số phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txttusoBK.Focus()
                txttusoBK.SelectAll()
                Exit Function
            End If
            strSQL += " AND Ordinal_No_List >=" & CInt(txttusoBK.Text)
        End If

        If (Trim$(txtDensoBK.Text) <> "") Then
            If (Not IsNumeric(Trim$(txtDensoBK.Text))) Then
                MsgBox("Điều kiện đến số phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtDensoBK.Focus()
                txtDensoBK.SelectAll()
                result = False
                GoTo endFunction
            End If

            If (CLng(Trim$(txtDensoBK.Text)) < 1) Then
                MsgBox("Điều kiện đến số phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtDensoBK.Focus()
                txtDensoBK.SelectAll()
                Exit Function
            End If
            strSQL += " AND Ordinal_No_List <=" & CInt(txtDensoBK.Text)
        End If

        If (Trim$(CboEmploy_code.Text) <> "") Then
            strSQL += " AND Employ_Code ='" & CboEmploy_code.Text & "'"
        End If

endFunction:
    End Function


    Private Sub cbostations_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbostations.SelectedIndexChanged
        If (cbostations.Items.Count > 0) Then
            Try
                strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
                Try
                    CboEmploy_code.DataSource = Nothing
                    cbonguoinhan.DataSource = Nothing
                    Try
                        mydataset.Tables("Tbl_Employee").Clear()
                        mydataset.Tables("cbonguoinhan").Clear()
                    Catch ex As Exception
                    End Try
                    CboEmploy_code.Items.Clear()
                    FillCombo(CboEmploy_code, strSQL, "Tbl_Employee", "Employ_Code", "Employ_Name")

                    cbonguoinhan.Items.Clear()
                    FillCombo(cbonguoinhan, strSQL, "cbonguoinhan", "Employ_Code", "Employ_Name")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
            End Try
            txtEmployeeName.Text = CboEmploy_code.SelectedValue
        End If
    End Sub

    Private Sub CboEmploy_code_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboEmploy_code.SelectedIndexChanged
        If (start) Then
            Try
                txtEmployeeName.Text = CboEmploy_code.SelectedValue
                cbonguoinhan.SelectedIndex = cbonguoinhan.FindString(CboEmploy_code.Text)
            Catch ex As Exception
            End Try
        End If
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

    Private Sub cmdLuu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLuu.Click
        UpdateExpenses()
    End Sub

    Public Sub UpdateExpenses()
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        Dim valueID As Long
        Dim SoGNT As Integer
        Dim strQuery As String
        Dim strBankcode As String
        Dim strAccountCode As String
        Dim strNguoiNop As String
        Dim soBK As Integer
        Dim NgayNop As Date
        NgayNop = DateTimePickerNgayNop.Value
        strAccountCode = cboAccounts_Banks.Text
        strNguoiNop = cbonguoinhan.Text
        strBankcode = GetStringName(mydataset, strAccountCode, "Tbl_Accounts_Banks", "Account_No", "Bank_Code")
        dt = DataGridListExpenes.DataSource
        Try
            For i = 0 To dt.Rows.Count - 1
                value = dt.Rows(i).Item("Check")
                If (value) Then
                    valueID = dt.Rows(i).Item("ID")
                    soBK = dt.Rows(i).Item("Ordinal_No_List")
                    Try
                        SoGNT = dt.Rows(i).Item("Pay_No")
                        If (SoGNT < 1) Then
                            Dim rt = MsgBox("Số bảng kê " & soBK & " không thể cập nhật được. Số giấy nộp tiền chưa được nhập." & vbCrLf & "Bạn có muốn tiếp tục không?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Lổi cập nhật")
                            If (rt = vbNo) Then
                                Exit Sub
                            Else
                                GoTo endUpdate
                            End If
                        End If
                    Catch ex As Exception
                        Dim rt = MsgBox("Số bảng kê " & soBK & " không thể cập nhật được. Số giấy nộp tiền chưa được nhập." & vbCrLf & "Bạn có muốn tiếp tục không?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Lổi cập nhật")
                        If (rt = vbNo) Then
                            Exit Sub
                        Else
                            GoTo endUpdate
                        End If
                    End Try

                    strQuery = " UPDATE Tbl_Expenses SET Pay_Date ='" & NgayNop & "', Account_No ='" & strAccountCode & "' , Bank_Code = '" & strBankcode & "', NguoiNop ='" & strNguoiNop & "', Pay_No = " & SoGNT & ", Status = True WHERE ID = " & valueID
                    UpdateExpense(strQuery)
endUpdate:
                End If
            Next
            MsgBox("Đã cập nhật. Xác nhận các phiếu thu đã chi nộp tiền vào ngân hàng.")
        Catch ex As Exception
        End Try
    End Sub

    Public Sub UpdateExpense(ByVal strQuery As String)
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
    Public Sub UpdateDatagrid()
        Dim i As Integer
        Dim dt As DataTable
        Dim value As Boolean
        Dim valueID As Long
        Dim strBankname As String
        Dim strAccountCode As String
        Dim strNguoiNop As String
        Dim strNgayNop As Date
        strAccountCode = cboAccounts_Banks.Text
        strNguoiNop = cbonguoinhan.Text
        strNgayNop = DateTimePickerNgayNop.Value
        strBankname = txtBank_Name.Text
        dt = DataGridListExpenes.DataSource
        Try
            For i = 0 To dt.Rows.Count - 1
                value = dt.Rows(i).Item("Check")
                If (value) Then
                    dt.Rows(i).Item("NguoiNop") = strNguoiNop
                    dt.Rows(i).Item("Bank_Code") = strBankname
                    dt.Rows(i).Item("Account_No") = strAccountCode
                    dt.Rows(i).Item("Pay_Date") = strNgayNop
                End If
            Next
        Catch ex As Exception
        End Try

    End Sub

    Private Sub cbonguoinhan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbonguoinhan.SelectedIndexChanged
        If (start) Then
            Try
                txtnguoinop.Text = cbonguoinhan.SelectedValue
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub cmdUpdateDatagrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateDatagrid.Click
        UpdateDatagrid()
    End Sub

    Private Sub cboAccounts_Banks_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAccounts_Banks.SelectedIndexChanged
        If (start) Then
            Try
                txtBank_Name.Text = cboAccounts_Banks.SelectedValue
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub cbostations_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbostations.KeyPress
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
                txtlydo.Text = ""
                Cbolydo.Focus()
                Exit Sub
            End If
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
            DateTimePickerTungay.Focus()
        End If
    End Sub

    Private Sub DateTimePickerTungay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerTungay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerDenngay.Focus()
        End If
    End Sub

    Private Sub DateTimePickerDenngay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerDenngay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttusoBK.Focus()
        End If
    End Sub

    Private Sub txttusoBK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttusoBK.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtDensoBK.Focus()
        End If
    End Sub

    Private Sub txtDensoBK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDensoBK.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdlist.Focus()
        End If
    End Sub

    Private Sub DateTimePickerNgayNop_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerNgayNop.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cbonguoinhan.Focus()
        End If
    End Sub

    Private Sub cbonguoinhan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbonguoinhan.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cboAccounts_Banks.Focus()
        End If
    End Sub

    Private Sub cboAccounts_Banks_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboAccounts_Banks.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdLuu.Focus()
        End If
    End Sub
End Class

