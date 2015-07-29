Imports System.Data.OleDb
Public Class frmBaoCaoNgayThang
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private table As New DataTable
    Dim start As Boolean = False
    Dim Dsrpt As New DsBaoCaoNgay

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        FillDataSet()
        FillCombo()

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
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpdenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtptungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents txttendichvu As System.Windows.Forms.TextBox
    Friend WithEvents cmbHTthu As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmdclose = New System.Windows.Forms.Button
        Me.cmdxem = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txttendichvu = New System.Windows.Forms.TextBox
        Me.dtpdenngay = New System.Windows.Forms.DateTimePicker
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtptungay = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cmbHTthu = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(72, 8)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(179, 27)
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
        Me.cmdclose.Location = New System.Drawing.Point(240, 120)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(80, 27)
        Me.cmdclose.TabIndex = 84
        Me.cmdclose.Text = "Đóng"
        '
        'cmdxem
        '
        Me.cmdxem.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdxem.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdxem.Location = New System.Drawing.Point(96, 120)
        Me.cmdxem.Name = "cmdxem"
        Me.cmdxem.Size = New System.Drawing.Size(80, 27)
        Me.cmdxem.TabIndex = 9
        Me.cmdxem.Text = "Xem BC"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txttendichvu)
        Me.GroupBox1.Controls.Add(Me.dtpdenngay)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtptungay)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(2, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(438, 80)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'txttendichvu
        '
        Me.txttendichvu.BackColor = System.Drawing.Color.White
        Me.txttendichvu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendichvu.ForeColor = System.Drawing.Color.Blue
        Me.txttendichvu.Location = New System.Drawing.Point(200, 48)
        Me.txttendichvu.Name = "txttendichvu"
        Me.txttendichvu.Size = New System.Drawing.Size(232, 26)
        Me.txttendichvu.TabIndex = 80
        Me.txttendichvu.Text = ""
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
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(72, 48)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(128, 27)
        Me.Cbolydo.TabIndex = 7
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
        Me.dtptungay.Location = New System.Drawing.Point(72, 16)
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
        Me.Label8.Location = New System.Drawing.Point(208, 19)
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
        Me.Label5.Location = New System.Drawing.Point(5, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 69
        Me.Label5.Text = "Dịch vụ"
        '
        'cmbHTthu
        '
        Me.cmbHTthu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHTthu.Location = New System.Drawing.Point(312, 8)
        Me.cmbHTthu.Name = "cmbHTthu"
        Me.cmbHTthu.Size = New System.Drawing.Size(128, 27)
        Me.cmbHTthu.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(267, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 22)
        Me.Label2.TabIndex = 86
        Me.Label2.Text = "Loại"
        '
        'frmBaoCaoNgayThang
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(442, 160)
        Me.Controls.Add(Me.cmbHTthu)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cmdclose)
        Me.Controls.Add(Me.cmdxem)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmBaoCaoNgayThang"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Báo cáo ngày - tháng"
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
        
        cbostations.DataSource = dt
        cbostations.DisplayMember = "Station_Name"
        cbostations.ValueMember = "StationID"

        dt = mydataset.Tables("Tbl_Services")
        Cbolydo.DataSource = dt
        Cbolydo.DisplayMember = "Service_Code"
        Cbolydo.ValueMember = "Service_Name"

        dt = mydataset.Tables("Tbl_LoaiThu")
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

                strQuery = "SELECT Receipt_Date , Ordinal_No ,  Charge_Cycle, Descriptions , Total_Money ,Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Service_Code,Invoice_Quantity,List_Quantity,List_Detail"
                strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                strQuery = strQuery & " WHERE Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND MaLoaiThu='" & cmbHTthu.SelectedValue & "' "

                If Trim(Cbolydo.Text) <> "" Then
                    strQuery = strQuery & "  AND Service_Code ='" & Cbolydo.Text & "'"
                End If
                FillReports(strQuery, "QryBaoCaoNgay")
                Dim rpt As CrystalReportNgayThang
                rpt = New CrystalReportNgayThang
                rpt.SetDataSource(Dsrpt)
                Dim frm As New frmPreview
                frm.CrystalReportViewerReceipts.ReportSource = rpt
                frm.ShowDialog()
            Case "GNT"

                strQuery = "SELECT Tbl_Receipts.Receipt_Date as NgayNop, Tbl_Receipts.Ordinal_No as SoPT, Tbl_Receipts.Pay_No as SoGNT, Tbl_Receipts.Charge_Cycle as KyCuoc, Tbl_Receipts.Descriptions as MoTa, Tbl_Receipts.Account_Code as SoTK, Tbl_Receipts.Total_Money as SoTT, Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS NguoiNop, Tbl_Receipts.Service_Code as DichVu"
                strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                strQuery = strQuery & " WHERE Tbl_Receipts.Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND Tbl_Receipts.MaLoaiThu='" & cmbHTthu.SelectedValue & "' "

                If Trim(Cbolydo.Text) <> "" Then
                    strQuery = strQuery & "  AND Service_Code ='" & Cbolydo.Text & "'"
                End If


                FillReports(strQuery, "qryBaoCaoGNT")
                Dim rpt As BaoCaoGNT
                    rpt = New BaoCaoGNT
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()

            Case "UNC"


                strQuery = "SELECT Tbl_Receipts.Receipt_Date, Tbl_Receipts.Ordinal_No, Tbl_Receipts.Service_Code, Tbl_Receipts.Descriptions, Tbl_Receipts.List_Quantity, Tbl_Receipts.Invoice_Quantity, Tbl_Receipts.Charge_Cycle, Tbl_Receipts.Total_Money, Tbl_Receipts.Employ_Code + ' - ' +  Tbl_Employee.Employ_Name AS Employ_Code , Tbl_Receipts.SLUNC " 'FROM Tbl_Receipts "
                strQuery = strQuery & " FROM Tbl_Receipts INNER JOIN Tbl_Employee ON Tbl_Receipts.Employ_Code = Tbl_Employee.Employ_Code"
                strQuery = strQuery & " WHERE Tbl_Receipts.Receipt_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "# AND Tbl_Receipts.MaLoaiThu='" & cmbHTthu.SelectedValue & "' "

                If Trim(Cbolydo.Text) <> "" Then
                    strQuery = strQuery & " AND Tbl_Receipts.Service_Code='" & Cbolydo.Text & "'"
                End If

                FillReports(strQuery, "QryReportsUNC")
                Dim rpt As BaoCaoUNC
                    rpt = New BaoCaoUNC
                    rpt.SetDataSource(Dsrpt)
                    Dim frm As New frmPreview
                    frm.CrystalReportViewerReceipts.ReportSource = rpt
                    frm.ShowDialog()

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
        str = str & " WHERE Expense_Date BETWEEN #" & dtptungay.Value.ToShortDateString & "# AND #" & dtpdenngay.Value.ToShortDateString & "#"

        If Trim(Cbolydo.Text) <> "" Then
            str = str & "  AND Service_Code ='" & Cbolydo.Text & "'"
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

    Private Sub dtpdenngay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtpdenngay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then

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

        End If
    End Sub

End Class
