Imports System.Data.OleDb
Public Class frmBaoCaoTonghopthu
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private table As New DataTable
    Dim start As Boolean = False
    Dim Dsrpt As New DsReceiptsDetail

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

        If (cmbHTthu.Items.Count > 0) Then
            If (cmbHTthu.FindString("TIỀN MẶT") > 0) Then
                cmbHTthu.SelectedIndex = cmbHTthu.FindString("TIỀN MẶT")
            End If
        End If

        If (Cbolydo.Items.Count > 0) Then
            txttendichvu.Text = Cbolydo.SelectedValue
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
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtpdenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtptungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbHTthu As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxDetail As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxPSTNDL As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxBaocaonhanh As System.Windows.Forms.CheckBox
    Friend WithEvents txttendichvu As System.Windows.Forms.TextBox
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBaoCaoTonghopthu))
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmdclose = New System.Windows.Forms.Button
        Me.cmdxem = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txttendichvu = New System.Windows.Forms.TextBox
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.CheckBoxBaocaonhanh = New System.Windows.Forms.CheckBox
        Me.CheckBoxPSTNDL = New System.Windows.Forms.CheckBox
        Me.dtpdenngay = New System.Windows.Forms.DateTimePicker
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtptungay = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.CheckBoxDetail = New System.Windows.Forms.CheckBox
        Me.cmbHTthu = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(85, 8)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(219, 27)
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
        Me.cmdclose.Location = New System.Drawing.Point(306, 198)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(80, 27)
        Me.cmdclose.TabIndex = 84
        Me.cmdclose.Text = "Đóng"
        '
        'cmdxem
        '
        Me.cmdxem.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdxem.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdxem.Location = New System.Drawing.Point(144, 198)
        Me.cmdxem.Name = "cmdxem"
        Me.cmdxem.Size = New System.Drawing.Size(80, 27)
        Me.cmdxem.TabIndex = 9
        Me.cmdxem.Text = "Xem BC"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txttendichvu)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.CheckBoxBaocaonhanh)
        Me.GroupBox1.Controls.Add(Me.CheckBoxPSTNDL)
        Me.GroupBox1.Controls.Add(Me.dtpdenngay)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtptungay)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.CheckBoxDetail)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(2, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(526, 160)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'txttendichvu
        '
        Me.txttendichvu.BackColor = System.Drawing.Color.White
        Me.txttendichvu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendichvu.ForeColor = System.Drawing.Color.Blue
        Me.txttendichvu.Location = New System.Drawing.Point(248, 45)
        Me.txttendichvu.Name = "txttendichvu"
        Me.txttendichvu.Size = New System.Drawing.Size(272, 26)
        Me.txttendichvu.TabIndex = 92
        Me.txttendichvu.Text = ""
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(79, 45)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(169, 27)
        Me.Cbolydo.TabIndex = 90
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(5, 45)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 91
        Me.Label5.Text = "Dịch vụ"
        '
        'CheckBoxBaocaonhanh
        '
        Me.CheckBoxBaocaonhanh.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxBaocaonhanh.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBoxBaocaonhanh.Location = New System.Drawing.Point(10, 133)
        Me.CheckBoxBaocaonhanh.Name = "CheckBoxBaocaonhanh"
        Me.CheckBoxBaocaonhanh.Size = New System.Drawing.Size(286, 24)
        Me.CheckBoxBaocaonhanh.TabIndex = 89
        Me.CheckBoxBaocaonhanh.Text = "Báo cáo nhanh tình hình thu - nộp cước"
        '
        'CheckBoxPSTNDL
        '
        Me.CheckBoxPSTNDL.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxPSTNDL.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBoxPSTNDL.Location = New System.Drawing.Point(280, 109)
        Me.CheckBoxPSTNDL.Name = "CheckBoxPSTNDL"
        Me.CheckBoxPSTNDL.Size = New System.Drawing.Size(128, 24)
        Me.CheckBoxPSTNDL.TabIndex = 88
        Me.CheckBoxPSTNDL.Text = "Đại Lý PSTN"
        '
        'dtpdenngay
        '
        Me.dtpdenngay.CustomFormat = "dd/MM/yyyy"
        Me.dtpdenngay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpdenngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpdenngay.Location = New System.Drawing.Point(296, 16)
        Me.dtpdenngay.Name = "dtpdenngay"
        Me.dtpdenngay.Size = New System.Drawing.Size(136, 26)
        Me.dtpdenngay.TabIndex = 4
        Me.dtpdenngay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboEmploy_code.Location = New System.Drawing.Point(79, 75)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(169, 27)
        Me.CboEmploy_code.TabIndex = 8
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(248, 75)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(272, 26)
        Me.txtEmployeeName.TabIndex = 72
        Me.txtEmployeeName.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(5, 75)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 22)
        Me.Label4.TabIndex = 71
        Me.Label4.Text = "Người nộp"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(5, 16)
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
        Me.dtptungay.Location = New System.Drawing.Point(80, 16)
        Me.dtptungay.Name = "dtptungay"
        Me.dtptungay.Size = New System.Drawing.Size(128, 26)
        Me.dtptungay.TabIndex = 3
        Me.dtptungay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(216, 19)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(69, 22)
        Me.Label8.TabIndex = 76
        Me.Label8.Text = "Đến ngày"
        '
        'CheckBoxDetail
        '
        Me.CheckBoxDetail.Checked = True
        Me.CheckBoxDetail.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxDetail.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxDetail.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBoxDetail.Location = New System.Drawing.Point(10, 109)
        Me.CheckBoxDetail.Name = "CheckBoxDetail"
        Me.CheckBoxDetail.Size = New System.Drawing.Size(270, 24)
        Me.CheckBoxDetail.TabIndex = 87
        Me.CheckBoxDetail.Text = "Xem chi tiết theo nhân viên, dịch vụ"
        '
        'cmbHTthu
        '
        Me.cmbHTthu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHTthu.Location = New System.Drawing.Point(376, 8)
        Me.cmbHTthu.Name = "cmbHTthu"
        Me.cmbHTthu.Size = New System.Drawing.Size(152, 27)
        Me.cmbHTthu.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(312, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 22)
        Me.Label2.TabIndex = 86
        Me.Label2.Text = "Loại thu"
        '
        'frmBaoCaoTonghopthu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(530, 231)
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
        Me.Name = "frmBaoCaoTonghopthu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Báo cáo tổng hợp thu cước"
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

        strSQL = "SELECT MaLoaithu,TenLoaiThu FROM Tbl_LoaiThu"
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_LoaiThu")
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

    End Sub

    Private Sub FillCombo()
        Dim dt As DataTable
        Dim dr As DataRow
        dt = mydataset.Tables("Tbl_Stations")

        cbostations.DataSource = dt
        cbostations.DisplayMember = "Station_Name"
        cbostations.ValueMember = "StationID"

        dt = mydataset.Tables("Tbl_LoaiThu")
        cmbHTthu.DataSource = dt
        cmbHTthu.DisplayMember = "TenLoaiThu"
        cmbHTthu.ValueMember = "MaLoaithu"

        Cbolydo.DataSource = mydataset.Tables("Tbl_Services")
        Cbolydo.DisplayMember = "Service_Code"
        Cbolydo.ValueMember = "Service_Name"

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
        Dsrpt.Tables("qryNgay").Rows(0).Item("Service_Code") = Cbolydo.Text

        'Fill gia tri ten khu vuc va don vi thu
        strQuery = "SELECT Tbl_Countries.CountryName, Tbl_Stations.Station_Name as StationName  FROM Tbl_Stations INNER JOIN Tbl_Countries ON Tbl_Stations.CountryCode = Tbl_Countries.CountryCode WHERE Tbl_Stations.StationID='" & cbostations.SelectedValue & "'"
        FillReports(strQuery, "qryCountry")
        If (cmbHTthu.Text <> "") Then
            'Lay DV 098
            strQuery = "SELECT Receipt_Date as Recei_Date , Ordinal_No AS Recei_No, Charge_Cycle,Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Receipts.Service_Code ,Invoice_Quantity AS Invoice_Quantity098 , Tbl_Receipts.Total_Money AS Recei_Money098,Account_Code AS Account_No, Maloaithu "
            strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
            strQuery = strQuery & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "'  AND Service_Code = '098' AND Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND MaLoaiThu='" & cmbHTthu.SelectedValue & "' "

            If Trim(CboEmploy_code.Text) <> "" Then
                strQuery = strQuery & "  AND Tbl_Receipts.Employ_Code ='" & CboEmploy_code.Text & "'"
            End If

            FillReports(strQuery, "qryReceiptsDetail")
            'Lay DV 178
            AddNewRow("178", "1")

            'Lay DV PSTN
            AddNewRow("PSTN", "1")

            'Lay DV ADSL
            AddNewRow("ADSL", "1")

            Select Case cmbHTthu.SelectedValue
                Case "TM"

                    If CheckBoxDetail.Checked Then

                        If (CheckBoxPSTNDL.Checked) Then

                            'Fill PSTNDL 
                            Dsrpt.Tables("qryPSTNDLReceiptsDetail").Clear()
                            strQuery = "SELECT Tbl_TNDN.Receipt_Date AS Recei_Date, Tbl_TNDN.Ordinal_No AS Recei_NoTNDN, Tbl_Receipts.Ordinal_No AS Recei_NoDS, Tbl_Expenses.Ordinal_No AS Expen_No, Tbl_Daily.TenDL AS AgentName, Tbl_Daily.MaDL AS ISDN, Tbl_Receipts.Charge_Cycle, Tbl_Receipts.Total_Money AS ReceiMoneyDS, Tbl_TNDN.Total_Money AS ReceiMoneyTNDN, Tbl_Expenses.Total_Money AS ExpenMoneyHH, Tbl_Receipts.Total_Money+Tbl_TNDN.Total_Money-Tbl_Expenses.Total_Money AS RealRevenue, Tbl_TNDN.Employ_Code+'-'+Tbl_Employee.Employ_Name AS Employ_Code "
                            strQuery = strQuery & " FROM (((Tbl_Receipts AS Tbl_TNDN INNER JOIN Tbl_Receipts ON Tbl_TNDN.STTDL = Tbl_Receipts.STTDL) INNER JOIN Tbl_Expenses ON Tbl_Receipts.STTDL = Tbl_Expenses.STTDL) INNER JOIN Tbl_Employee ON Tbl_TNDN.Employ_Code = Tbl_Employee.Employ_Code) INNER JOIN Tbl_Daily ON Tbl_TNDN.TenDaily = Tbl_Daily.MaDL "
                            strQuery = strQuery & " WHERE (((Tbl_Receipts.STTDL)>0) AND ((Tbl_TNDN.Service_Code)='PSTNTNDN') AND ((Tbl_Receipts.Service_Code)='PSTNDS') AND (Tbl_TNDN.Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "#  ))"

                            If Trim(CboEmploy_code.Text) <> "" Then
                                strQuery = strQuery & "  AND Tbl_TNDN.Employ_Code ='" & CboEmploy_code.Text & "'"
                            End If
                            FillReports(strQuery, "qryPSTNDLReceiptsDetail")

                            Dim rptDL As CrystalReportPSTNDLReceiptDetail
                            rptDL = New CrystalReportPSTNDLReceiptDetail
                            rptDL.SetDataSource(Dsrpt)

                            Dim frmDL As New frmPreview
                            frmDL.CrystalReportViewerReceipts.ReportSource = rptDL
                            frmDL.ShowDialog()
                            Exit Select
                        End If

                        Dim rpt As CrystalReportReceiptDetail
                        rpt = New CrystalReportReceiptDetail
                        rpt.SetDataSource(Dsrpt)

                        Dim frm As New frmPreview
                        frm.CrystalReportViewerReceipts.ReportSource = rpt
                        frm.ShowDialog()
                    Else
                        AddNewRowPSTNDL()
                        Dim rpt As CrystalReportSumReceipts
                        rpt = New CrystalReportSumReceipts
                        rpt.SetDataSource(Dsrpt)
                        Dim frm As New frmPreview
                        frm.CrystalReportViewerReceipts.ReportSource = rpt
                        frm.ShowDialog()

                    End If

                Case "GNT"

                    'Lay DV PSTNDL
                    AddNewRow("PSTNDL", "1")

                    If CheckBoxDetail.Checked Then
                        Dim rpt As CrystalReportTongHopGNTDetail
                        rpt = New CrystalReportTongHopGNTDetail
                        rpt.SetDataSource(Dsrpt)

                        Dim frm As New frmPreview
                        frm.CrystalReportViewerReceipts.ReportSource = rpt
                        frm.ShowDialog()
                    Else
                        Dim rpt As CrystalReportSumGNT
                        rpt = New CrystalReportSumGNT
                        rpt.SetDataSource(Dsrpt)
                        Dim frm As New frmPreview
                        frm.CrystalReportViewerReceipts.ReportSource = rpt
                        frm.ShowDialog()
                    End If

                Case "UNC"

                    'Lay DV ADSL
                    AddNewRow("PSTNDL", "1")

                    If CheckBoxDetail.Checked Then
                        Dim rpt As CrystalReportUNCReceiptDetail
                        rpt = New CrystalReportUNCReceiptDetail
                        rpt.SetDataSource(Dsrpt)

                        Dim frm As New frmPreview
                        frm.CrystalReportViewerReceipts.ReportSource = rpt
                        frm.ShowDialog()
                    Else
                        Dim rpt As CrystalReportSumUNC
                        rpt = New CrystalReportSumUNC
                        rpt.SetDataSource(Dsrpt)
                        Dim frm As New frmPreview
                        frm.CrystalReportViewerReceipts.ReportSource = rpt
                        frm.ShowDialog()
                    End If
            End Select
        Else
            If (Cbolydo.Text <> "") Then

                Dim rpt As CrystalReportReceiptByService
                rpt = New CrystalReportReceiptByService
                'strQuery = " SELECT Ordinal_No AS Recei_No, Receipt_Date AS Recei_Date, Charge_Cycle, Employ_Code , Round([Total_Money]/1.1,0) AS Recei_Money, Round([Total_Money]-[Total_Money]/1.1,0) AS Vat, Total_Money AS SumofMoney,MaLoaiThu FROM Tbl_Receipts "

                strQuery = "  SELECT Ordinal_No AS Recei_No, Receipt_Date AS Recei_Date, Charge_Cycle, Employ_Name AS Employ_Code, Round([Total_Money]/1.1,0) AS Recei_Money, Round([Total_Money]-[Total_Money]/1.1,0) AS Vat, Total_Money AS SumofMoney, MaLoaiThu, Invoice_Quantity FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code "
                strQuery = strQuery & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "'  AND Service_Code = '" & Cbolydo.Text & "' AND (Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# )"

                If Trim(CboEmploy_code.Text) <> "" Then
                    strQuery = strQuery & "  AND Tbl_Receipts.Employ_Code ='" & CboEmploy_code.Text & "'"
                End If

                strQuery = strQuery & "  ORDER BY Charge_Cycle"
                FillReports(strQuery, "qryReceiptsDetailByService")

                rpt.SetDataSource(Dsrpt)
                Dim frm As New frmPreview
                frm.CrystalReportViewerReceipts.ReportSource = rpt
                frm.ShowDialog()
                Exit Sub
            End If

            SumTongHop()

            If CheckBoxDetail.Checked Then
                Dim rpt As CrystalReportReceiptDetailOfEmployee
                rpt = New CrystalReportReceiptDetailOfEmployee
                rpt.SetDataSource(Dsrpt)

                Dim frm As New frmPreview
                frm.CrystalReportViewerReceipts.ReportSource = rpt
                frm.ShowDialog()
                Exit Sub
            Else

                If CheckBoxBaocaonhanh.Checked Then
                    Dim rpt As CrystalReportBaoCaoNhanh
                    rpt = New CrystalReportBaoCaoNhanh
                    rpt.SetDataSource(Dsrpt)

                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()
                    Exit Sub
                Else
                    Dim rpt As CrystalReportSumTMGNTUNC
                    rpt = New CrystalReportSumTMGNTUNC
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()
                End If
            End If
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

    Private Sub SumTongHop()
        Dim strQuery As String
        Dim dtRow As DataRow
        'Lay DV 098
        strQuery = "SELECT Receipt_Date as Recei_Date , Ordinal_No AS Recei_No, Charge_Cycle,Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Receipts.Service_Code ,Invoice_Quantity AS Invoice_Quantity098 , Tbl_Receipts.Total_Money AS Recei_Money098,Account_Code AS Account_No, Maloaithu "
        strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
        strQuery = strQuery & " WHERE Service_Code = '098' AND Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# "

        If Trim(CboEmploy_code.Text) <> "" Then
            strQuery = strQuery & "  AND Tbl_Receipts.Employ_Code ='" & CboEmploy_code.Text & "'"
        End If

        FillReports(strQuery, "qryReceiptsDetail")
        'Lay DV 178
        AddNewRow("178", "0")

        'Lay DV PSTN
        AddNewRow("PSTN", "0")

        'Lay DV ADSL
        AddNewRow("ADSL", "0")

        'Lay DV PSTNDL
        AddNewRow("PSTNDL", "0")

        'Lay DV PSTNDL LOAI TM
        AddNewRowPSTNDL()

    End Sub
    Private Sub AddNewRow(ByVal strService As String, ByVal strType As String)
        Dim newrow As DataRow
        Dim dsNew As New DataSet
        Dim str As String
        Dim i
        Dim count As Long
        Dim oleda As OleDbDataAdapter
        Select Case strType
            Case "0"
                str = "SELECT Receipt_Date AS Recei_Date , Ordinal_No AS Recei_No, Charge_Cycle,Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Receipts.Service_Code,Invoice_Quantity AS Invoice_Quantity" & strService & " , Tbl_Receipts.Total_Money AS Recei_Money" & strService & " ,Account_Code AS Account_No,MaLoaithu "
                str = str & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                str = str & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "' AND Service_Code = '" & strService & "' AND Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# "
            Case "1"
                str = "SELECT Receipt_Date AS Recei_Date , Ordinal_No AS Recei_No, Charge_Cycle,Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Receipts.Service_Code,Invoice_Quantity AS Invoice_Quantity" & strService & " , Tbl_Receipts.Total_Money AS Recei_Money" & strService & " ,Account_Code AS Account_No,MaLoaithu "
                str = str & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                str = str & " WHERE Tbl_Employee.StationID = '" & cbostations.SelectedValue & "' AND Service_Code = '" & strService & "' AND Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND MaLoaiThu='" & cmbHTthu.SelectedValue & "' "
        End Select

        If Trim(CboEmploy_code.Text) <> "" Then
            str = str & "  AND Tbl_Receipts.Employ_Code ='" & CboEmploy_code.Text & "'"
        End If

        Try
            oleda = New OleDbDataAdapter(str, oledbcon)
            oleda.Fill(dsNew)
            count = Dsrpt.Tables("qryReceiptsDetail").Rows.Count
            For i = 0 To dsNew.Tables(0).Rows.Count - 1
                newrow = Dsrpt.Tables("qryReceiptsDetail").NewRow
                Dsrpt.Tables("qryReceiptsDetail").Rows.Add(newrow)
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Recei_Date") = dsNew.Tables(0).Rows(i)("Recei_Date")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Recei_No") = dsNew.Tables(0).Rows(i)("Recei_No")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Account_No") = dsNew.Tables(0).Rows(i)("Account_No")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Service_Code") = dsNew.Tables(0).Rows(i)("Service_Code")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Charge_Cycle") = dsNew.Tables(0).Rows(i)("Charge_Cycle")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Invoice_Quantity" & strService) = dsNew.Tables(0).Rows(i)("Invoice_Quantity" & strService)
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Recei_Money" & strService) = dsNew.Tables(0).Rows(i)("Recei_Money" & strService)
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Employ_Code") = dsNew.Tables(0).Rows(i)("Employ_Code")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("MaLoaiThu") = dsNew.Tables(0).Rows(i)("MaLoaiThu")
                count += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AddNewRowPSTNDL()
        Dim newrow As DataRow
        Dim dsNew As New DataSet
        Dim str As String
        Dim i
        Dim count As Long
        Dim oleda As OleDbDataAdapter

        str = "SELECT Tbl_TNDN.Receipt_Date AS Recei_Date, Tbl_Receipts.Charge_Cycle,Tbl_Receipts.Invoice_Quantity AS Invoice_QuantityPSTNDL,Tbl_Receipts.Total_Money + Tbl_TNDN.Total_Money - Tbl_Expenses.Total_Money AS Recei_MoneyPSTNDL,Tbl_TNDN.MaLoaithu "
        str = str & " FROM (((Tbl_Receipts AS Tbl_TNDN INNER JOIN Tbl_Receipts ON Tbl_TNDN.STTDL = Tbl_Receipts.STTDL) INNER JOIN Tbl_Expenses ON Tbl_Receipts.STTDL = Tbl_Expenses.STTDL) INNER JOIN Tbl_Employee ON Tbl_TNDN.Employ_Code = Tbl_Employee.Employ_Code) INNER JOIN Tbl_Daily ON Tbl_TNDN.TenDaily = Tbl_Daily.MaDL "
        str = str & " WHERE (((Tbl_Receipts.STTDL)>0) AND ((Tbl_TNDN.Service_Code)='PSTNTNDN') AND ((Tbl_Receipts.Service_Code)='PSTNDS') AND (Tbl_TNDN.Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "#  ))"

        'str = str & " FROM (((Tbl_Daily INNER JOIN Tbl_Receipts AS Tbl_TNDN ON Tbl_Daily.TenDL = Tbl_TNDN.TenDaily) INNER JOIN Tbl_Receipts ON Tbl_TNDN.STTDL = Tbl_Receipts.STTDL) INNER JOIN Tbl_Expenses ON Tbl_Receipts.STTDL = Tbl_Expenses.STTDL) INNER JOIN Tbl_Employee ON Tbl_TNDN.Employ_Code = Tbl_Employee.Employ_Code "
        'str = str & " WHERE (((Tbl_Receipts.STTDL)>0) AND ((Tbl_TNDN.Service_Code)='PSTNTNDN') AND ((Tbl_Receipts.Service_Code)='PSTNDS') AND (Tbl_TNDN.Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "#  ))"

        If Trim(CboEmploy_code.Text) <> "" Then
            str = str & "  AND Tbl_TNDN.Employ_Code ='" & CboEmploy_code.Text & "'"
        End If

        Try
            oleda = New OleDbDataAdapter(str, oledbcon)
            oleda.Fill(dsNew)
            count = Dsrpt.Tables("qryReceiptsDetail").Rows.Count
            For i = 0 To dsNew.Tables(0).Rows.Count - 1
                newrow = Dsrpt.Tables("qryReceiptsDetail").NewRow
                Dsrpt.Tables("qryReceiptsDetail").Rows.Add(newrow)
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Recei_Date") = dsNew.Tables(0).Rows(i)("Recei_Date")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Invoice_QuantityPSTNDL") = dsNew.Tables(0).Rows(i)("Invoice_QuantityPSTNDL")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Recei_MoneyPSTNDL") = dsNew.Tables(0).Rows(i)("Recei_MoneyPSTNDL")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("Charge_Cycle") = dsNew.Tables(0).Rows(i)("Charge_Cycle")
                Dsrpt.Tables("qryReceiptsDetail").Rows(count).Item("MaLoaiThu") = dsNew.Tables(0).Rows(i)("MaLoaiThu")
                count += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
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
    Private Sub Cbolydo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbolydo.SelectedIndexChanged

        If (start) Then
            txttendichvu.Text = Cbolydo.SelectedValue
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

    Private Sub CheckBoxBaocaonhanh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxBaocaonhanh.CheckedChanged
        If (CheckBoxBaocaonhanh.Checked) Then
            CheckBoxDetail.Checked = False
            CheckBoxPSTNDL.Checked = False
        End If
    End Sub

    Private Sub CheckBoxDetail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxDetail.CheckedChanged
        If (CheckBoxDetail.Checked) Then
            CheckBoxBaocaonhanh.Checked = False
        End If
    End Sub

    Private Sub CheckBoxPSTNDL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPSTNDL.CheckedChanged
        If (CheckBoxPSTNDL.Checked) Then
            CheckBoxBaocaonhanh.Checked = False
        End If
    End Sub

End Class
