Imports System.Data.OleDb
Public Class frmBaoCaoChiNop
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Dim start As Boolean = False
    Dim Dsrpt As New DsChiNop
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        FillDataSet()
        FillCombo()


        If (cbostations.Items.Count > 0) Then
            cbostations.SelectedIndex = 0
            Try
                strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
                Try
                    Dim cmd As New OleDbCommand(strSQL, oledbcon)
                    da = New OleDbDataAdapter(cmd)
                    da.Fill(mydataset, "Tbl_Employee")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                CboEmploy_code.DataSource = mydataset.Tables("Tbl_Employee")
                CboEmploy_code.DisplayMember = "Employ_Code"
                CboEmploy_code.ValueMember = "Employ_Name"

                txtEmployeeName.Text = CboEmploy_code.SelectedValue
            Catch ex As Exception
            End Try
        End If

        DateTimePickerdenngay.Value = Now
        DateTimePickertungay.Value = Now
        start = True
        If (cboAccounts.Items.Count > 0) Then
            txtBank_Accountname.Text = cboAccounts.SelectedValue
        End If

        If (Cbolydo.Items.Count > 0) Then
            txtTenLydo.Text = Cbolydo.SelectedValue
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
    Friend WithEvents cmdclose As System.Windows.Forms.Button
    Friend WithEvents cmdxem As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboAccounts As System.Windows.Forms.ComboBox
    Friend WithEvents txtBank_Accountname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerdenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickertungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtDensoPC As System.Windows.Forms.TextBox
    Friend WithEvents txtTusoPC As System.Windows.Forms.TextBox
    Friend WithEvents txtTenLydo As System.Windows.Forms.TextBox
    Friend WithEvents RadioButtonDaNop As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonChuaNop As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonAll As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBaoCaoChiNop))
        Me.cmdclose = New System.Windows.Forms.Button
        Me.cmdxem = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RadioButtonAll = New System.Windows.Forms.RadioButton
        Me.RadioButtonChuaNop = New System.Windows.Forms.RadioButton
        Me.RadioButtonDaNop = New System.Windows.Forms.RadioButton
        Me.txtDensoPC = New System.Windows.Forms.TextBox
        Me.txtTusoPC = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtTenLydo = New System.Windows.Forms.TextBox
        Me.cboAccounts = New System.Windows.Forms.ComboBox
        Me.txtBank_Accountname = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.DateTimePickerdenngay = New System.Windows.Forms.DateTimePicker
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.DateTimePickertungay = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdclose
        '
        Me.cmdclose.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdclose.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdclose.Location = New System.Drawing.Point(286, 211)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(80, 27)
        Me.cmdclose.TabIndex = 11
        Me.cmdclose.Text = "Đóng"
        '
        'cmdxem
        '
        Me.cmdxem.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdxem.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdxem.Location = New System.Drawing.Point(110, 211)
        Me.cmdxem.Name = "cmdxem"
        Me.cmdxem.Size = New System.Drawing.Size(80, 27)
        Me.cmdxem.TabIndex = 10
        Me.cmdxem.Text = "Xem BC"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.txtDensoPC)
        Me.GroupBox1.Controls.Add(Me.txtTusoPC)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtTenLydo)
        Me.GroupBox1.Controls.Add(Me.cboAccounts)
        Me.GroupBox1.Controls.Add(Me.txtBank_Accountname)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerdenngay)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.DateTimePickertungay)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(4, 30)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(476, 176)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButtonAll)
        Me.GroupBox2.Controls.Add(Me.RadioButtonChuaNop)
        Me.GroupBox2.Controls.Add(Me.RadioButtonDaNop)
        Me.GroupBox2.Location = New System.Drawing.Point(232, 136)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(240, 35)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        '
        'RadioButtonAll
        '
        Me.RadioButtonAll.Checked = True
        Me.RadioButtonAll.ForeColor = System.Drawing.SystemColors.Desktop
        Me.RadioButtonAll.Location = New System.Drawing.Point(169, 12)
        Me.RadioButtonAll.Name = "RadioButtonAll"
        Me.RadioButtonAll.Size = New System.Drawing.Size(65, 20)
        Me.RadioButtonAll.TabIndex = 2
        Me.RadioButtonAll.TabStop = True
        Me.RadioButtonAll.Text = "Tất cả"
        '
        'RadioButtonChuaNop
        '
        Me.RadioButtonChuaNop.ForeColor = System.Drawing.SystemColors.Desktop
        Me.RadioButtonChuaNop.Location = New System.Drawing.Point(82, 12)
        Me.RadioButtonChuaNop.Name = "RadioButtonChuaNop"
        Me.RadioButtonChuaNop.Size = New System.Drawing.Size(92, 20)
        Me.RadioButtonChuaNop.TabIndex = 1
        Me.RadioButtonChuaNop.Text = "Chưa nộp"
        '
        'RadioButtonDaNop
        '
        Me.RadioButtonDaNop.ForeColor = System.Drawing.SystemColors.Desktop
        Me.RadioButtonDaNop.Location = New System.Drawing.Point(10, 12)
        Me.RadioButtonDaNop.Name = "RadioButtonDaNop"
        Me.RadioButtonDaNop.Size = New System.Drawing.Size(72, 20)
        Me.RadioButtonDaNop.TabIndex = 0
        Me.RadioButtonDaNop.Text = "Đã nộp"
        '
        'txtDensoPC
        '
        Me.txtDensoPC.BackColor = System.Drawing.Color.White
        Me.txtDensoPC.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDensoPC.ForeColor = System.Drawing.Color.Blue
        Me.txtDensoPC.Location = New System.Drawing.Point(185, 144)
        Me.txtDensoPC.Name = "txtDensoPC"
        Me.txtDensoPC.Size = New System.Drawing.Size(42, 26)
        Me.txtDensoPC.TabIndex = 8
        Me.txtDensoPC.Text = ""
        '
        'txtTusoPC
        '
        Me.txtTusoPC.BackColor = System.Drawing.Color.White
        Me.txtTusoPC.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTusoPC.ForeColor = System.Drawing.Color.Blue
        Me.txtTusoPC.Location = New System.Drawing.Point(70, 143)
        Me.txtTusoPC.Name = "txtTusoPC"
        Me.txtTusoPC.Size = New System.Drawing.Size(42, 26)
        Me.txtTusoPC.TabIndex = 7
        Me.txtTusoPC.Text = ""
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label7.Location = New System.Drawing.Point(112, 147)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(76, 22)
        Me.Label7.TabIndex = 86
        Me.Label7.Text = "Đến số PC"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(4, 144)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 22)
        Me.Label1.TabIndex = 85
        Me.Label1.Text = "Từ số PC"
        '
        'txtTenLydo
        '
        Me.txtTenLydo.BackColor = System.Drawing.Color.White
        Me.txtTenLydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTenLydo.ForeColor = System.Drawing.Color.Blue
        Me.txtTenLydo.Location = New System.Drawing.Point(232, 48)
        Me.txtTenLydo.Name = "txtTenLydo"
        Me.txtTenLydo.Size = New System.Drawing.Size(240, 26)
        Me.txtTenLydo.TabIndex = 84
        Me.txtTenLydo.Text = ""
        '
        'cboAccounts
        '
        Me.cboAccounts.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAccounts.Location = New System.Drawing.Point(70, 112)
        Me.cboAccounts.Name = "cboAccounts"
        Me.cboAccounts.Size = New System.Drawing.Size(162, 27)
        Me.cboAccounts.TabIndex = 6
        '
        'txtBank_Accountname
        '
        Me.txtBank_Accountname.BackColor = System.Drawing.Color.White
        Me.txtBank_Accountname.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBank_Accountname.ForeColor = System.Drawing.Color.Blue
        Me.txtBank_Accountname.Location = New System.Drawing.Point(232, 112)
        Me.txtBank_Accountname.Name = "txtBank_Accountname"
        Me.txtBank_Accountname.Size = New System.Drawing.Size(240, 26)
        Me.txtBank_Accountname.TabIndex = 83
        Me.txtBank_Accountname.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(6, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 22)
        Me.Label2.TabIndex = 81
        Me.Label2.Text = "Số TK"
        '
        'DateTimePickerdenngay
        '
        Me.DateTimePickerdenngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerdenngay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerdenngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerdenngay.Location = New System.Drawing.Point(336, 16)
        Me.DateTimePickerdenngay.Name = "DateTimePickerdenngay"
        Me.DateTimePickerdenngay.Size = New System.Drawing.Size(136, 26)
        Me.DateTimePickerdenngay.TabIndex = 3
        Me.DateTimePickerdenngay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboEmploy_code.Location = New System.Drawing.Point(70, 80)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(162, 27)
        Me.CboEmploy_code.TabIndex = 5
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(232, 80)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(240, 26)
        Me.txtEmployeeName.TabIndex = 72
        Me.txtEmployeeName.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(8, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 22)
        Me.Label4.TabIndex = 71
        Me.Label4.Text = "Ng nhận"
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(70, 48)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(162, 27)
        Me.Cbolydo.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 22)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Từ ngày"
        '
        'DateTimePickertungay
        '
        Me.DateTimePickertungay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickertungay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickertungay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickertungay.Location = New System.Drawing.Point(70, 16)
        Me.DateTimePickertungay.Name = "DateTimePickertungay"
        Me.DateTimePickertungay.Size = New System.Drawing.Size(136, 26)
        Me.DateTimePickertungay.TabIndex = 2
        Me.DateTimePickertungay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(248, 19)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(69, 22)
        Me.Label8.TabIndex = 76
        Me.Label8.Text = "Đến ngày"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(8, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 69
        Me.Label5.Text = "Dịch vụ"
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(76, 8)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(360, 27)
        Me.cbostations.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(12, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 22)
        Me.Label6.TabIndex = 86
        Me.Label6.Text = "Tổ thu"
        '
        'frmBaoCaoChiNop
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 246)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.cmdclose)
        Me.Controls.Add(Me.cmdxem)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label6)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmBaoCaoChiNop"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Báo cáo chi - nộp tiền ngân hàng"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub FillDataSet()
        mydataset = New DataSet

        strSQL = "SELECT StationID,Station_Name,Station_Address FROM Tbl_Stations "
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_Stations")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        strSQL = "SELECT Service_Code,Service_Name FROM Tbl_Services "
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_Services")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        strSQL = "SELECT Account_No,Tbl_Banks.Bank_Code,Bank_Name FROM Tbl_Banks,Tbl_Accounts_Banks WHERE Tbl_Accounts_Banks.Bank_Code = Tbl_Banks.Bank_Code "
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_Banks_Accounts")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub FillCombo()
        
        cbostations.DataSource = mydataset.Tables("Tbl_Stations")
        cbostations.DisplayMember = "Station_Name"
        cbostations.ValueMember = "StationID"

        Cbolydo.DataSource = mydataset.Tables("Tbl_Services")
        Cbolydo.DisplayMember = "Service_Code"
        Cbolydo.ValueMember = "Service_Name"

        cboAccounts.DataSource = mydataset.Tables("Tbl_Banks_Accounts")
        cboAccounts.DisplayMember = "Account_No"
        cboAccounts.ValueMember = "Bank_Name"
    End Sub
    Private Sub cmdclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdclose.Click
        Me.Close()
    End Sub

    Private Sub cbostations_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbostations.SelectedIndexChanged
        If (start) Then
            Try
                CboEmploy_code.DataSource = Nothing
                Try
                    mydataset.Tables("Tbl_Employee").Clear()
                Catch ex As Exception
                End Try
                CboEmploy_code.Items.Clear()

                strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
                Try
                    Dim cmd As New OleDbCommand(strSQL, oledbcon)
                    da = New OleDbDataAdapter(cmd)
                    da.Fill(mydataset, "Tbl_Employee")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                
                CboEmploy_code.DataSource = mydataset.Tables("Tbl_Employee")
                CboEmploy_code.DisplayMember = "Employ_Code"
                CboEmploy_code.ValueMember = "Employ_Name"
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Cbolydo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbolydo.SelectedIndexChanged

        If (start) Then
            txtTenLydo.Text = Cbolydo.SelectedValue
        End If

    End Sub

    Private Sub CboEmploy_code_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboEmploy_code.SelectedIndexChanged
        If (start) Then
            txtEmployeeName.Text = CboEmploy_code.SelectedValue
        End If
    End Sub

    Private Sub cboAccounts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAccounts.SelectedIndexChanged
        If (start) Then
            txtBank_Accountname.Text = cboAccounts.SelectedValue
        End If
    End Sub

    Private Sub cmdxem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdxem.Click
        Dim rpt As CrystalReportChiNop
        rpt = New CrystalReportChiNop
        Dsrpt.Clear()
        Dim strQuery As String

        'Lay thong tin trung tam va don vi thu cuoc
        strQuery = "SELECT Tbl_Countries.CountryName, Tbl_Stations.Station_Name FROM Tbl_Stations INNER JOIN Tbl_Countries ON Tbl_Stations.CountryCode = Tbl_Countries.CountryCode WHERE Tbl_Stations.StationID='" & cbostations.SelectedValue & "'"
        FillReports(strQuery, "GetCountry_station")

        'Fill : GNT, chukycuoc,Ngay,MaNhanVien,TenNhanVien,DichVu
        Dim Newrow As DataRow
        Newrow = Dsrpt.Tables("valueNgay").NewRow

        Dsrpt.Tables("valueNgay").Rows.Add(Newrow)
        Dsrpt.Tables("valueNgay").Rows(0).Item("TuNgay") = DateTimePickertungay.Text
        Dsrpt.Tables("valueNgay").Rows(0).Item("DenNgay") = DateTimePickerdenngay.Text

        ' Lay bang detail
        strQuery = "SELECT Tbl_Expenses.Expense_Date, Tbl_Expenses.Ordinal_No, Tbl_Expenses.Descriptions, Tbl_Expenses.List_Quantity, Tbl_Expenses.Charge_Cycle, Tbl_Expenses.Employ_Code, Tbl_Expenses.Account_No, Tbl_Expenses.Bank_Code, Tbl_Expenses.Status, Tbl_Expenses.Service_Code, Tbl_Expenses.Total_Money,Pay_No FROM Tbl_Expenses INNER JOIN Tbl_Employee ON Tbl_Expenses.Employ_Code = Tbl_Employee.Employ_Code " & _
                   " WHERE  Tbl_Employee.StationID = '" & cbostations.SelectedValue & "' AND Expense_Date BETWEEN #" & DateTimePickertungay.Value.ToShortDateString & " # AND #" & DateTimePickerdenngay.Value.ToShortDateString & "# "

        Select Case Cbolydo.Text
            Case "PSTNHH"
                strQuery += " AND Service_Code = 'PSTNHH'"
            Case Else
                If (Trim$(Cbolydo.Text) <> "") Then
                    strQuery += " AND Service_Code = '" & Cbolydo.Text & "'"
                Else
                    strQuery += " AND Service_Code <> 'PSTNHH'"
                End If
        End Select

        If (Trim$(CboEmploy_code.Text) <> "") Then
            strQuery += " AND Employ_Code = '" & CboEmploy_code.Text & "'"
        End If

        If (Trim$(cboAccounts.Text) <> "") Then
            strQuery += " AND Account_No = '" & cboAccounts.Text & "'"
        End If

        If (RadioButtonDaNop.Checked) Then
            strQuery += " AND Tbl_Expenses.Status = True "
        Else
            If (RadioButtonChuaNop.Checked) Then
                strQuery += " AND Tbl_Expenses.Status = False "
            End If
        End If

        If (Trim$(txtTusoPC.Text) <> "" And IsNumeric(Trim$(txtTusoPC.Text))) Then
            strQuery += " AND Ordinal_No >= " & Trim$(txtTusoPC.Text)
        End If

        If (Trim$(txtDensoPC.Text) <> "" And IsNumeric(Trim$(txtDensoPC.Text))) Then
            strQuery += " AND Ordinal_No <= " & Trim$(txtDensoPC.Text)
        End If

        FillReports(strQuery, "QryReportsChiNop")
        rpt.SetDataSource(Dsrpt)
        Dim frm As New frmPreview
        frm.CrystalReportViewerReceipts.ReportSource = rpt
        frm.ShowDialog()
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
            cboAccounts.Focus()
        End If
    End Sub

    Private Sub cboAccounts_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboAccounts.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtTusoPC.Focus()
        End If
    End Sub

    Private Sub txtTusoPC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTusoPC.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtDensoPC.Focus()
        End If
    End Sub

    Private Sub txtDensoPC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDensoPC.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdxem.Focus()
        End If
    End Sub
End Class
