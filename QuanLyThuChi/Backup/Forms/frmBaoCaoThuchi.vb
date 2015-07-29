Imports System.Data.OleDb
Public Class frmBaoCaoThuchi
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private table As New DataTable
    Dim start As Boolean = False
    Dim Dsrpt As New dsBaoCaoThuChi

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
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

                CboEmploy_code.DataSource = mydataset.Tables("Tbl_Employee") 'dt
                CboEmploy_code.DisplayMember = "Employ_Code"
                CboEmploy_code.ValueMember = "Employ_Name"

                If (CboEmploy_code.Items.Count > 0) Then
                    CboEmploy_code.SelectedIndex = 0 'CboEmploy_code.FindString("All")
                    txtEmployeeName.Text = CboEmploy_code.SelectedValue
                End If
            Catch ex As Exception
            End Try
        End If

        start = True
        dtpdenngay.Value = Now
        dtptungay.Value = Now


        If (Cbolydo.Items.Count > 0) Then
            Cbolydo.SelectedIndex = 0 'CboEmploy_code.FindString("All")
            txttendichvu.Text = Cbolydo.SelectedValue
        End If

        If (cmbHTthu.Items.Count > 0) Then
            If (cmbHTthu.FindString("TIỀN MẶT") > 0) Then
                cmbHTthu.SelectedIndex = cmbHTthu.FindString("TIỀN MẶT")
            End If
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
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdclose As System.Windows.Forms.Button
    Friend WithEvents cmdxem As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpdenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtptungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents txttendichvu As System.Windows.Forms.TextBox
    Friend WithEvents cmbHTthu As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxDetail As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBaoCaoThuchi))
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmdclose = New System.Windows.Forms.Button
        Me.cmdxem = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txttendichvu = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.dtpdenngay = New System.Windows.Forms.DateTimePicker
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtptungay = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cmbHTthu = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CheckBoxDetail = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(85, 8)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(195, 27)
        Me.cbostations.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(16, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 22)
        Me.Label6.TabIndex = 81
        Me.Label6.Text = "Tổ thu"
        '
        'cmdclose
        '
        Me.cmdclose.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdclose.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdclose.Location = New System.Drawing.Point(440, 168)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(80, 27)
        Me.cmdclose.TabIndex = 84
        Me.cmdclose.Text = "Đóng"
        '
        'cmdxem
        '
        Me.cmdxem.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdxem.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdxem.Location = New System.Drawing.Point(312, 168)
        Me.cmdxem.Name = "cmdxem"
        Me.cmdxem.Size = New System.Drawing.Size(80, 27)
        Me.cmdxem.TabIndex = 9
        Me.cmdxem.Text = "Xem BC"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.txttendichvu)
        Me.GroupBox1.Controls.Add(Me.dtpdenngay)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtptungay)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(2, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(534, 128)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'txttendichvu
        '
        Me.txttendichvu.BackColor = System.Drawing.Color.White
        Me.txttendichvu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendichvu.ForeColor = System.Drawing.Color.Blue
        Me.txttendichvu.Location = New System.Drawing.Point(248, 57)
        Me.txttendichvu.Name = "txttendichvu"
        Me.txttendichvu.Size = New System.Drawing.Size(280, 26)
        Me.txttendichvu.TabIndex = 80
        Me.txttendichvu.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CheckBox2)
        Me.GroupBox2.Controls.Add(Me.CheckBox1)
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox2.Location = New System.Drawing.Point(388, 9)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(136, 46)
        Me.GroupBox2.TabIndex = 77
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Loại phiếu"
        '
        'CheckBox2
        '
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBox2.Location = New System.Drawing.Point(72, 16)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(56, 24)
        Me.CheckBox2.TabIndex = 1
        Me.CheckBox2.Text = "Chi"
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBox1.Location = New System.Drawing.Point(11, 16)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(56, 24)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Thu"
        '
        'dtpdenngay
        '
        Me.dtpdenngay.CustomFormat = "dd/MM/yyyy"
        Me.dtpdenngay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpdenngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpdenngay.Location = New System.Drawing.Point(277, 22)
        Me.dtpdenngay.Name = "dtpdenngay"
        Me.dtpdenngay.Size = New System.Drawing.Size(104, 26)
        Me.dtpdenngay.TabIndex = 4
        Me.dtpdenngay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboEmploy_code.Location = New System.Drawing.Point(79, 89)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(169, 27)
        Me.CboEmploy_code.TabIndex = 8
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(248, 89)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(280, 26)
        Me.txtEmployeeName.TabIndex = 72
        Me.txtEmployeeName.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(5, 89)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 22)
        Me.Label4.TabIndex = 71
        Me.Label4.Text = "Người nộp"
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(80, 57)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(168, 27)
        Me.Cbolydo.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(5, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 22)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Từ ngày"
        '
        'dtptungay
        '
        Me.dtptungay.CustomFormat = "dd/MM/yyyy"
        Me.dtptungay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtptungay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtptungay.Location = New System.Drawing.Point(80, 22)
        Me.dtptungay.Name = "dtptungay"
        Me.dtptungay.Size = New System.Drawing.Size(104, 26)
        Me.dtptungay.TabIndex = 3
        Me.dtptungay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(197, 25)
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
        Me.Label5.Location = New System.Drawing.Point(5, 57)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 69
        Me.Label5.Text = "Dịch vụ"
        '
        'cmbHTthu
        '
        Me.cmbHTthu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHTthu.Location = New System.Drawing.Point(383, 8)
        Me.cmbHTthu.Name = "cmbHTthu"
        Me.cmbHTthu.Size = New System.Drawing.Size(152, 27)
        Me.cmbHTthu.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(344, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 22)
        Me.Label2.TabIndex = 86
        Me.Label2.Text = "Loại"
        '
        'CheckBoxDetail
        '
        Me.CheckBoxDetail.Checked = True
        Me.CheckBoxDetail.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxDetail.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxDetail.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBoxDetail.Location = New System.Drawing.Point(16, 168)
        Me.CheckBoxDetail.Name = "CheckBoxDetail"
        Me.CheckBoxDetail.Size = New System.Drawing.Size(272, 24)
        Me.CheckBoxDetail.TabIndex = 87
        Me.CheckBoxDetail.Text = "Xem chi tiết theo nhân viên, dịch vụ"
        '
        'frmBaoCaoThuchi
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(538, 207)
        Me.Controls.Add(Me.CheckBoxDetail)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmbHTthu)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cmdclose)
        Me.Controls.Add(Me.cmdxem)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmBaoCaoThuchi"
        Me.Text = "Báo cáo thu - chi"
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

        strSQL = "SELECT MaLoaithu,TenLoaiThu FROM Tbl_LoaiThu"
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_LoaiThu")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Private Sub FillCombo()
        Dim dt As DataTable
        Dim dr As DataRow
        dt = mydataset.Tables("Tbl_Stations")
        'dr = dt.NewRow
        'dr(0) = "Tất cả"
        'dr(1) = "All"
        'dt.Rows.Add(dr)
        cbostations.DataSource = dt
        cbostations.DisplayMember = "Station_Name"
        cbostations.ValueMember = "StationID"

        dt = mydataset.Tables("Tbl_Services")
        'dr = dt.NewRow
        'dr(0) = "Tất cả"
        'dr(1) = "All"
        'dt.Rows.Add(dr)
        Cbolydo.DataSource = dt
        Cbolydo.DisplayMember = "Service_Code"
        Cbolydo.ValueMember = "Service_Name"

        dt = mydataset.Tables("Tbl_LoaiThu")
        'dr = dt.NewRow
        'dr(0) = "Tất cả"
        'dr(1) = "All"
        'dt.Rows.Add(dr)
        cmbHTthu.DataSource = dt
        cmbHTthu.DisplayMember = "TenLoaiThu"
        cmbHTthu.ValueMember = "MaLoaithu"


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
                'Dim dt As DataTable
                'Dim dr As DataRow
                'dt = mydataset.Tables("Tbl_Employee")
                'dr = dt.NewRow
                'dr(0) = "Tất cả"
                'dr(1) = "All"
                'dt.Rows.Add(dr)
                CboEmploy_code.DataSource = mydataset.Tables("Tbl_Employee") 'dt
                CboEmploy_code.DisplayMember = "Employ_Code"
                CboEmploy_code.ValueMember = "Employ_Name"
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub CboEmploy_code_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboEmploy_code.SelectedIndexChanged
        If (start) Then
            txtEmployeeName.Text = CboEmploy_code.SelectedValue
        End If
    End Sub

    Private Sub cmdclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdclose.Click
        Me.Close()
    End Sub

    Private Sub cmdxem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdxem.Click
        
        Dsrpt.Clear()
        Dim strQuery As String
        Dim dtRow As DataRow

        'Fill gia tri tu ngay den ngay
        dtRow = Dsrpt.Tables("qryNgay").NewRow
        Dsrpt.Tables("qryNgay").Rows.Add(dtRow)
        Dsrpt.Tables("qryNgay").Rows(0).Item("fromDate") = dtptungay.Value.ToShortDateString
        Dsrpt.Tables("qryNgay").Rows(0).Item("toDate") = dtpdenngay.Value.ToShortDateString

        'Fill gia tri ten khu vuc va don vi thu
        strQuery = "SELECT Tbl_Countries.CountryName, Tbl_Stations.Station_Name as StationName  FROM Tbl_Stations INNER JOIN Tbl_Countries ON Tbl_Stations.CountryCode = Tbl_Countries.CountryCode WHERE Tbl_Stations.StationID='" & cbostations.SelectedValue & "'"
        FillReports(strQuery, "qryCountry")

        Select Case cmbHTthu.SelectedValue
            Case "TM"

                strQuery = "SELECT Receipt_Date AS Recei_Expen_Date, Ordinal_No AS Recei_No,  Charge_Cycle, Descriptions , Total_Money AS Recei_Money, Tbl_Receipts.Total_Money, Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Receipts.Service_Code,Invoice_Quantity,List_Quantity,List_Detail"
                strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                strQuery = strQuery & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "'  AND Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND  MaLoaiThu='" & cmbHTthu.SelectedValue & "' "
                'val(LEFT(Tbl_Receipts.Charge_Cycle,2)) BETWEEN " & dtptuchuky.Value.Month & " AND " & dtpdenchuky.Value.Month & " AND val(RIGHT(Tbl_Receipts.Charge_Cycle,4)) BETWEEN " & dtptuchuky.Value.Year & " AND " & dtpdenchuky.Value.Year & " AND

                If Trim(Cbolydo.Text) <> "" Then
                    strQuery = strQuery & "  AND Service_Code ='" & Cbolydo.Text & "'"
                End If

                If Trim(CboEmploy_code.Text) <> "" Then
                    strQuery = strQuery & "  AND Tbl_Receipts.Employ_Code ='" & CboEmploy_code.Text & "'"
                End If

                If CheckBox1.Checked = True And CheckBox2.Checked = True Then
                    FillReports(strQuery, "qryBaoCaoThuChi")
                    AddNewRow()
                Else
                    If CheckBox1.Checked = True And CheckBox2.Checked = False Then
                        FillReports(strQuery, "qryBaoCaoThuChi")
                    ElseIf CheckBox1.Checked = False And CheckBox2.Checked = True Then
                        AddNewRow()
                    End If
                End If
                If CheckBoxDetail.Checked Then
                    Dim rpt As CrystalReportThuChiDetail
                    rpt = New CrystalReportThuChiDetail
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()
                Else
                    Dim rpt As CrystalReportThuChi
                    rpt = New CrystalReportThuChi
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()

                End If

            Case "GNT"
                

                strQuery = "SELECT Tbl_Receipts.Receipt_Date as NgayNop, Tbl_Receipts.Ordinal_No as SoPT, Tbl_Receipts.Pay_No as SoGNT, Tbl_Receipts.Charge_Cycle as KyCuoc, Tbl_Receipts.Descriptions as MoTa, Tbl_Receipts.Account_Code as SoTK, Tbl_Receipts.Total_Money as SoTT, Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS NguoiNop, Tbl_Receipts.Service_Code as DichVu,List_Quantity AS SLBK,List_Detail AS CTBK, Invoice_Quantity AS SLHD"
                strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                strQuery = strQuery & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "'  AND Tbl_Receipts.Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND Tbl_Receipts.MaLoaiThu='" & cmbHTthu.SelectedValue & "' "
                'val(LEFT(Tbl_Receipts.Charge_Cycle,2)) BETWEEN " & dtptuchuky.Value.Month & " AND " & dtpdenchuky.Value.Month & " AND val(RIGHT(Tbl_Receipts.Charge_Cycle,4)) BETWEEN " & dtptuchuky.Value.Year & " AND " & dtpdenchuky.Value.Year & " AND
                If Trim(Cbolydo.Text) <> "" Then
                    strQuery = strQuery & "  AND Service_Code ='" & Cbolydo.Text & "'"
                End If

                If Trim(CboEmploy_code.Text) <> "" Then
                    strQuery = strQuery & "  AND Tbl_Receipts.Employ_Code ='" & CboEmploy_code.Text & "'"
                End If

                FillReports(strQuery, "qryBaoCaoGNT")
                If CheckBoxDetail.Checked Then
                    Dim rpt As BaoCaoGNTDetail
                    rpt = New BaoCaoGNTDetail
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()
                Else
                    Dim rpt As BaoCaoGNT
                    rpt = New BaoCaoGNT
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()
                End If

            Case "UNC"
                

                strQuery = "SELECT Tbl_Receipts.Receipt_Date, Tbl_Receipts.Ordinal_No, Tbl_Receipts.Service_Code, Tbl_Receipts.Descriptions, Tbl_Receipts.List_Quantity, Tbl_Receipts.Invoice_Quantity, Tbl_Receipts.Charge_Cycle, Tbl_Receipts.Total_Money, Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Receipts.SLUNC " 'FROM Tbl_Receipts "
                strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                strQuery = strQuery & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "'  AND Tbl_Receipts.Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND Tbl_Receipts.MaLoaiThu='" & cmbHTthu.SelectedValue & "' "
                'AND  val(LEFT(Tbl_Receipts.Charge_Cycle,2)) BETWEEN " & dtptuchuky.Value.Month & " AND " & dtpdenchuky.Value.Month & " AND val(RIGHT(Tbl_Receipts.Charge_Cycle,4)) BETWEEN " & dtptuchuky.Value.Year & " AND " & dtpdenchuky.Value.Year & " AND
                If Trim(CboEmploy_code.Text) <> "" Then
                    strQuery = strQuery & " AND Tbl_Receipts.Employ_Code='" & CboEmploy_code.Text & "'"
                End If

                If Trim(Cbolydo.Text) <> "" Then
                    strQuery = strQuery & " AND Tbl_Receipts.Service_Code='" & Cbolydo.Text & "'"
                End If

                FillReports(strQuery, "QryReportsUNC")
                If CheckBoxDetail.Checked Then
                    Dim rpt As BaoCaoUNCDetail
                    rpt = New BaoCaoUNCDetail
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()
                Else
                    Dim rpt As BaoCaoUNC
                    rpt = New BaoCaoUNC
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()
                End If
        End Select

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
    Private Sub AddNewRow()
        Dim newrow As DataRow
        Dim dsNew As New DataSet
        Dim str As String
        Dim i
        Dim count As Long
        Dim oleda As OleDbDataAdapter

        str = "SELECT Expense_Date AS Recei_Expen_Date,  Ordinal_No AS Expen_No, Charge_Cycle, Descriptions, Total_Money AS Expen_Money, Tbl_Expenses.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Expenses.Service_Code,List_Quantity,List_Detail "
        str = str & " FROM Tbl_Expenses INNER JOIN Tbl_Employee ON Tbl_Expenses.Employ_Code = Tbl_Employee.Employ_Code "
        str = str & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "'  AND Expense_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "#"
        'AND  val(LEFT(Tbl_Expenses.Charge_Cycle,2)) BETWEEN " & dtptuchuky.Value.Month & " AND " & dtpdenchuky.Value.Month & " AND val(RIGHT(Tbl_Expenses.Charge_Cycle,4)) BETWEEN " & dtptuchuky.Value.Year & " AND " & dtpdenchuky.Value.Year

        If Trim(Cbolydo.Text) <> "" Then
            str = str & "  AND Service_Code ='" & Cbolydo.Text & "'"
        End If

        If Trim(CboEmploy_code.Text) <> "" Then
            str = str & "  AND Tbl_Expenses.Employ_Code ='" & CboEmploy_code.Text & "'"
        End If

        Try
            oleda = New OleDbDataAdapter(str, oledbcon)
            oleda.Fill(dsNew)
            count = Dsrpt.Tables("qryBaoCaoThuChi").Rows.Count
            For i = 0 To dsNew.Tables(0).Rows.Count - 1
                newrow = Dsrpt.Tables("qryBaoCaoThuChi").NewRow
                Dsrpt.Tables("qryBaoCaoThuChi").Rows.Add(newrow)
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("Recei_Expen_Date") = dsNew.Tables(0).Rows(i)("Recei_Expen_Date")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("Expen_No") = dsNew.Tables(0).Rows(i)("Expen_No")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("Charge_Cycle") = dsNew.Tables(0).Rows(i)("Charge_Cycle")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("Descriptions") = dsNew.Tables(0).Rows(i)("Descriptions")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("Expen_Money") = dsNew.Tables(0).Rows(i)("Expen_Money")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("Employ_Code") = dsNew.Tables(0).Rows(i)("Employ_Code")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("List_Quantity") = dsNew.Tables(0).Rows(i)("List_Quantity")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("List_Detail") = dsNew.Tables(0).Rows(i)("List_Detail")
                Dsrpt.Tables("qryBaoCaoThuChi").Rows(count).Item("Service_Code") = dsNew.Tables(0).Rows(i)("Service_Code")
                count += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Cbolydo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbolydo.SelectedIndexChanged
        If (start) Then
            txttendichvu.Text = Cbolydo.SelectedValue
        End If
    End Sub

    Private Sub cbostations_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbostations.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmbHTthu.Focus()
        End If
    End Sub

    Private Sub cmbHTthu_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbHTthu.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            dtptungay.Focus()
        End If
    End Sub

    Private Sub dtptungay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtptungay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            dtpdenngay.Focus()
        End If
    End Sub

    'Private Sub dtpdenngay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtpdenngay.KeyPress
    '    Dim KeyAscii As Short = Asc(e.KeyChar)
    '    If (KeyAscii = 13) Then
    '        dtptuchuky.Focus()
    '    End If
    'End Sub

    'Private Sub dtptuchuky_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '    Dim KeyAscii As Short = Asc(e.KeyChar)
    '    If (KeyAscii = 13) Then
    '        dtpdenchuky.Focus()
    '    End If
    'End Sub

    Private Sub dtpdenchuky_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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
            cmdxem.Focus()
        End If
    End Sub

End Class
