Imports System.Data.OleDb
Public Class frmTongHopGNT
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Dim start As Boolean = False
    Dim Dsrpt As New DsTongHopGNT

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
    Friend WithEvents cmdxem As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtTenLydo As System.Windows.Forms.TextBox
    Friend WithEvents cboAccounts As System.Windows.Forms.ComboBox
    Friend WithEvents txtBank_Accountname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerdenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickertungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdclose As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmTongHopGNT))
        Me.cmdxem = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
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
        Me.cmdclose = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdxem
        '
        Me.cmdxem.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdxem.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdxem.Location = New System.Drawing.Point(117, 182)
        Me.cmdxem.Name = "cmdxem"
        Me.cmdxem.Size = New System.Drawing.Size(80, 27)
        Me.cmdxem.TabIndex = 93
        Me.cmdxem.Text = "Xem BC"
        '
        'GroupBox1
        '
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
        Me.GroupBox1.Location = New System.Drawing.Point(5, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(476, 144)
        Me.GroupBox1.TabIndex = 92
        Me.GroupBox1.TabStop = False
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
        Me.cboAccounts.TabIndex = 82
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
        Me.Label2.Location = New System.Drawing.Point(7, 112)
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
        Me.DateTimePickerdenngay.TabIndex = 75
        Me.DateTimePickerdenngay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboEmploy_code.Location = New System.Drawing.Point(70, 80)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(162, 27)
        Me.CboEmploy_code.TabIndex = 70
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
        Me.Label4.Location = New System.Drawing.Point(7, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 22)
        Me.Label4.TabIndex = 71
        Me.Label4.Text = "Ng Nộp"
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(70, 48)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(162, 27)
        Me.Cbolydo.TabIndex = 67
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(7, 16)
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
        Me.DateTimePickertungay.TabIndex = 66
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
        Me.Label5.Location = New System.Drawing.Point(7, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 69
        Me.Label5.Text = "Dịch vụ"
        '
        'cmdclose
        '
        Me.cmdclose.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdclose.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdclose.Location = New System.Drawing.Point(285, 182)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(80, 27)
        Me.cmdclose.TabIndex = 94
        Me.cmdclose.Text = "Đóng"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(13, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 22)
        Me.Label6.TabIndex = 91
        Me.Label6.Text = "Tổ thu"
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(77, 8)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(403, 27)
        Me.cbostations.TabIndex = 90
        '
        'frmTongHopGNT
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 214)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.cmdxem)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdclose)
        Me.Controls.Add(Me.Label6)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmTongHopGNT"
        Me.Text = "Tổng hợp giấy nộp tiền"
        Me.GroupBox1.ResumeLayout(False)
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

    Private Sub FillReports(ByVal strQuery As String, ByVal strTableName As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(Dsrpt, strTableName)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub cmdxem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdxem.Click
        Dim rpt As CrystalReportGNT
        rpt = New CrystalReportGNT
        Dsrpt.Clear()
        Dim strQuery As String

        'Fill gia tri ten khu vuc va don vi thu
        strQuery = "SELECT Tbl_Countries.CountryName, Tbl_Stations.Station_Name as StationName  FROM Tbl_Stations INNER JOIN Tbl_Countries ON Tbl_Stations.CountryCode = Tbl_Countries.CountryCode WHERE Tbl_Stations.StationID='" & cbostations.SelectedValue & "'"
        FillReports(strQuery, "qryCountry")

        'Fill : GNT, chukycuoc,Ngay,MaNhanVien,TenNhanVien,DichVu
        Dim Newrow As DataRow
        Newrow = Dsrpt.Tables("valueNgay").NewRow

        Dsrpt.Tables("valueNgay").Rows.Add(Newrow)
        Dsrpt.Tables("valueNgay").Rows(0).Item("TuNgay") = DateTimePickertungay.Text
        Dsrpt.Tables("valueNgay").Rows(0).Item("DenNgay") = DateTimePickerdenngay.Text


        'Fill GNT phieu thu
        strQuery = " SELECT Receipt_Date AS NgayThuchi, Ordinal_No AS SoPhieuThuchi, Charge_Cycle AS KyCuoc, Total_Money  AS TienThu, Employ_Code AS Nhanvien, Account_Code AS SoTaiKhoan, Pay_Date AS NgayNopThu, Pay_No AS SoGNT, Service_Code AS Dichvu FROM Tbl_Receipts WHERE MaLoaiThu='GNT'" & _
                   " AND Receipt_Date BETWEEN #" & DateTimePickertungay.Value.ToShortDateString & " # AND #" & DateTimePickerdenngay.Value.ToShortDateString & "# "

        If (Trim$(Cbolydo.Text) <> "") Then
            strQuery += " AND Service_Code = '" & Cbolydo.Text & "'"
        End If

        If (Trim$(CboEmploy_code.Text) <> "") Then
            strQuery += " AND Employ_Code = '" & CboEmploy_code.Text & "'"
        End If

        If (Trim$(cboAccounts.Text) <> "") Then
            strQuery += " AND Account_Code = '" & cboAccounts.Text & "'"
        End If
        FillReports(strQuery, "QryGNTNopThu")

        'Fill GNT chi nop
        AddNewRows()

        rpt.SetDataSource(Dsrpt)
        Dim frm As New frmPreview
        frm.CrystalReportViewerReceipts.ReportSource = rpt
        frm.ShowDialog()
    End Sub

    Private Sub AddNewRows()
        Dim Newrow As DataRow
        Dim i As Integer
        Dim ds As New DataSet
        Dim strQuery As String
        Dim count As Integer
        count = Dsrpt.Tables("QryGNTNopThu").Rows.Count

        strQuery = "SELECT Expense_Date AS NgayThuchi, Ordinal_No AS SoPhieuThuchi , Charge_Cycle AS KyCuoc, Total_Money AS TienNop, Employ_Code AS Nhanvien, Account_No AS SoTaiKhoan, Pay_Date AS NgayNopThu, Pay_No AS SoGNT, Service_Code AS Dichvu FROM Tbl_Expenses WHERE Pay_no > 0 " & _
                       " AND Expense_Date BETWEEN #" & DateTimePickertungay.Value.ToShortDateString & " # AND #" & DateTimePickerdenngay.Value.ToShortDateString & "# AND Status = True "

        Select Case Cbolydo.Text
            Case "PSTNHH"
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



        Dim cmd As New OleDbCommand(strQuery, oledbcon)
        da = New OleDbDataAdapter(cmd)
        da.Fill(ds, "ListThuGNT")

        For i = 0 To ds.Tables("ListThuGNT").Rows.Count - 1
            Newrow = Dsrpt.Tables("QryGNTNopThu").NewRow
            Dsrpt.Tables("QryGNTNopThu").Rows.Add(Newrow)
            Dsrpt.Tables("QryGNTNopThu").Rows(count).Item("SoPhieuThuchi") = ds.Tables("ListThuGNT").Rows(i).Item("SoPhieuThuchi")
            Dsrpt.Tables("QryGNTNopThu").Rows(count).Item("KyCuoc") = ds.Tables("ListThuGNT").Rows(i).Item("KyCuoc")
            Dsrpt.Tables("QryGNTNopThu").Rows(count).Item("Dichvu") = ds.Tables("ListThuGNT").Rows(i).Item("Dichvu")
            Dsrpt.Tables("QryGNTNopThu").Rows(count).Item("NgayNopThu") = ds.Tables("ListThuGNT").Rows(i).Item("NgayNopThu")
            Dsrpt.Tables("QryGNTNopThu").Rows(count).Item("TienNop") = ds.Tables("ListThuGNT").Rows(i).Item("TienNop")
            Dsrpt.Tables("QryGNTNopThu").Rows(count).Item("Nhanvien") = ds.Tables("ListThuGNT").Rows(i).Item("Nhanvien")
            Dsrpt.Tables("QryGNTNopThu").Rows(count).Item("SoTaiKhoan") = ds.Tables("ListThuGNT").Rows(i).Item("SoTaiKhoan")
            count += 1
        Next
        
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
End Class
