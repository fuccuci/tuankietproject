Imports System.Data.OleDb
Public Class frmDaily
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private ds As DataSet
    Dim flag, varMaDL As String

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
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdEditEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdNewEmployee As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtdiachitthuongtru As System.Windows.Forms.TextBox
    Friend WithEvents txttendaily As System.Windows.Forms.TextBox
    Friend WithEvents txtmadaily As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgridDaily As System.Windows.Forms.DataGrid
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdSaveEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdClosed As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cmdSaveEmployee = New System.Windows.Forms.Button
        Me.cmdClosed = New System.Windows.Forms.Button
        Me.cmdEditEmployee = New System.Windows.Forms.Button
        Me.cmdNewEmployee = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtdiachitthuongtru = New System.Windows.Forms.TextBox
        Me.txttendaily = New System.Windows.Forms.TextBox
        Me.txtmadaily = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgridDaily = New System.Windows.Forms.DataGrid
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgridDaily, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.Label9.Location = New System.Drawing.Point(8, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(656, 48)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "THÔNG TIN ĐẠI LÝ"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdDelete)
        Me.GroupBox2.Controls.Add(Me.cmdSaveEmployee)
        Me.GroupBox2.Controls.Add(Me.cmdClosed)
        Me.GroupBox2.Controls.Add(Me.cmdEditEmployee)
        Me.GroupBox2.Controls.Add(Me.cmdNewEmployee)
        Me.GroupBox2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox2.Location = New System.Drawing.Point(8, 376)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(656, 54)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(320, 15)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(88, 32)
        Me.cmdDelete.TabIndex = 18
        Me.cmdDelete.Text = "Xóa"
        '
        'cmdSaveEmployee
        '
        Me.cmdSaveEmployee.Location = New System.Drawing.Point(216, 14)
        Me.cmdSaveEmployee.Name = "cmdSaveEmployee"
        Me.cmdSaveEmployee.Size = New System.Drawing.Size(88, 32)
        Me.cmdSaveEmployee.TabIndex = 5
        Me.cmdSaveEmployee.Text = "Lưu"
        '
        'cmdClosed
        '
        Me.cmdClosed.Location = New System.Drawing.Point(560, 15)
        Me.cmdClosed.Name = "cmdClosed"
        Me.cmdClosed.Size = New System.Drawing.Size(88, 32)
        Me.cmdClosed.TabIndex = 16
        Me.cmdClosed.Text = "Đóng"
        '
        'cmdEditEmployee
        '
        Me.cmdEditEmployee.Location = New System.Drawing.Point(112, 14)
        Me.cmdEditEmployee.Name = "cmdEditEmployee"
        Me.cmdEditEmployee.Size = New System.Drawing.Size(88, 32)
        Me.cmdEditEmployee.TabIndex = 13
        Me.cmdEditEmployee.Text = "Thay Đổi"
        '
        'cmdNewEmployee
        '
        Me.cmdNewEmployee.Location = New System.Drawing.Point(8, 14)
        Me.cmdNewEmployee.Name = "cmdNewEmployee"
        Me.cmdNewEmployee.Size = New System.Drawing.Size(88, 32)
        Me.cmdNewEmployee.TabIndex = 17
        Me.cmdNewEmployee.Text = "Nhập Mới"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtdiachitthuongtru)
        Me.GroupBox1.Controls.Add(Me.txttendaily)
        Me.GroupBox1.Controls.Add(Me.txtmadaily)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 56)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(656, 80)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        '
        'txtdiachitthuongtru
        '
        Me.txtdiachitthuongtru.BackColor = System.Drawing.Color.White
        Me.txtdiachitthuongtru.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdiachitthuongtru.Location = New System.Drawing.Point(86, 43)
        Me.txtdiachitthuongtru.Name = "txtdiachitthuongtru"
        Me.txtdiachitthuongtru.Size = New System.Drawing.Size(562, 26)
        Me.txtdiachitthuongtru.TabIndex = 3
        Me.txtdiachitthuongtru.Text = ""
        '
        'txttendaily
        '
        Me.txttendaily.BackColor = System.Drawing.Color.White
        Me.txttendaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendaily.Location = New System.Drawing.Point(312, 14)
        Me.txttendaily.Name = "txttendaily"
        Me.txttendaily.Size = New System.Drawing.Size(336, 26)
        Me.txttendaily.TabIndex = 2
        Me.txttendaily.Text = ""
        '
        'txtmadaily
        '
        Me.txtmadaily.BackColor = System.Drawing.Color.White
        Me.txtmadaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmadaily.ForeColor = System.Drawing.Color.Blue
        Me.txtmadaily.Location = New System.Drawing.Point(86, 14)
        Me.txtmadaily.Name = "txtmadaily"
        Me.txtmadaily.Size = New System.Drawing.Size(154, 26)
        Me.txtmadaily.TabIndex = 1
        Me.txtmadaily.Text = ""
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(13, 43)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 24)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Địa chỉ"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(243, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Tên đại lý"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(10, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Mã đại lý"
        '
        'dgridDaily
        '
        Me.dgridDaily.DataMember = ""
        Me.dgridDaily.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgridDaily.Location = New System.Drawing.Point(8, 144)
        Me.dgridDaily.Name = "dgridDaily"
        Me.dgridDaily.ReadOnly = True
        Me.dgridDaily.Size = New System.Drawing.Size(656, 232)
        Me.dgridDaily.TabIndex = 68
        '
        'frmDaily
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(672, 434)
        Me.Controls.Add(Me.dgridDaily)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmDaily"
        Me.Text = "Nhập đại lý"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgridDaily, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FormatDataGridDaily()

        With dgridDaily
            .AllowNavigation = False
            .DataMember = "Tbl_Daily"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách các đại lý"
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "Tbl_Daily"
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

                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "MaDL"
                    .HeaderText = "Mã đại lý   "
                    .Width = 100
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "TenDL"
                    .HeaderText = "Tên đại lý"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
                '.Add(New DataGridDateTimePicker)
                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Diachi"
                    .HeaderText = "Địa chỉ"
                    .Width = 320
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
            End With
        End With
        dgridDaily.TableStyles.Add(TblStyle)
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
        'strAccountCode = cboAccounts_Banks.Text
        'strNguoiNop = cbonguoinhan.Text
        'strNgayNop = DateTimePickerNgayNop.Value
        'strBankname = txtBank_Name.Text
        'dt = DataGridListExpenes.DataSource
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
    Private Sub FillDataset(ByVal strQuery As String)
        Try
            'ds = Nothing
            ds = New DataSet
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds, "Tbl_Daily")
            dgridDaily.DataSource = ds.Tables("Tbl_Daily")
        Catch ex As Exception
            MsgBox("Error Fill DataSet :" & ex.ToString)
        End Try
    End Sub

    Private Sub frmDaily_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormatDataGridDaily()
        strSQL = "SELECT * FROM Tbl_Daily ORDER BY MaDL"
        FillDataset(strSQL)
        BindingTextBox()
        cmdNewEmployee.Focus()
        flag = "_Add"
    End Sub

    Private Sub cmdNewEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewEmployee.Click
        txtmadaily.DataBindings.Clear()
        txtmadaily.Text = ""
        txttendaily.Text = ""
        txtdiachitthuongtru.Text = ""
        txtmadaily.Focus()
        flag = "_Add"
        cmdSaveEmployee.Text = "Lưu"
        cmdEditEmployee.Enabled = False
    End Sub

    Private Sub txtmadaily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmadaily.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttendaily.Focus()
        End If
    End Sub

    Private Sub txttendaily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttendaily.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtdiachitthuongtru.Focus()
        End If
    End Sub

    Private Sub txtdiachitthuongtru_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdiachitthuongtru.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdSaveEmployee.Focus()
        End If
    End Sub
    Private Sub SaveDataSource(ByVal strMaDL As String, ByVal strTenDL As String, ByVal strDiachi As String)
        Dim sql As String
        If flag = "_Add" Then
            sql = "INSERT INTO Tbl_Daily(MaDL,TenDL,DiaChi) VALUES('" & strMaDL & "','" & strTenDL & "','" & strDiachi & "')"
        Else
            sql = "UPDATE Tbl_Daily SET MaDL='" & strMaDL & "' , TenDL='" & strTenDL & "' , DiaChi='" & strDiachi & "' WHERE MaDL='" & varMaDL & "'"
        End If
        Try

            oledbcon.Open()
            Dim cmd As New OleDbCommand(sql, oledbcon)
            cmd.ExecuteNonQuery()
            MsgBox("Dữ liệu đã được cập nhật thành công!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Thông báo")

        Catch ex As Exception
            MsgBox("Error Fill DataSet :" & ex.ToString)
        End Try
        FillDataset(strSQL)
        BindingTextBox()
        oledbcon.Close()
    End Sub
    Private Sub cmdSaveEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveEmployee.Click
        If txtmadaily.Text = "" Then
            MsgBox("Mã đại lý không được rỗng!!!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Chú ý")
            Exit Sub
        End If
        If flag = "_Add" Then
            If (checkDaily()) Then
                SaveDataSource(Trim(txtmadaily.Text), txttendaily.Text, txtdiachitthuongtru.Text)
            Else
                MsgBox("Đại lý này đã tồn tại. Vui lòng kiểm tra lại!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Thông báo")
                Exit Sub
            End If
        Else
            SaveDataSource(txtmadaily.Text, txttendaily.Text, txtdiachitthuongtru.Text)
        End If

        'txtmadaily.Text = vbNullString
        'xttendaily.Text = vbNullString
        'txtdiachitthuongtru.Text = vbNullString
        cmdNewEmployee.Enabled = True
        cmdEditEmployee.Enabled = True
        cmdNewEmployee.Focus()
        flag = ""
    End Sub

    Private Sub cmdEditEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditEmployee.Click
        cmdSaveEmployee.Text = "Cập nhật"
        flag = "_Edit"
        varMaDL = txtmadaily.Text
        txtmadaily.Focus()
    End Sub
    Private Sub BindingTextBox()
        Try
            txtmadaily.DataBindings.Clear()
            txttendaily.DataBindings.Clear()
            txtdiachitthuongtru.DataBindings.Clear()
            txtmadaily.DataBindings.Add("Text", ds.Tables("Tbl_Daily"), "MaDL")
            txttendaily.DataBindings.Add("Text", ds.Tables("Tbl_Daily"), "TenDL")
            txtdiachitthuongtru.DataBindings.Add("Text", ds.Tables("Tbl_Daily"), "DiaChi")
            'dgridDaily.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
            'Exit Sub
        End Try
    End Sub

    Private Sub frmDaily_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        flag = ""
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim olecom As OleDbCommand
        Dim sqlDelete, ask As String
        oledbcon.Open()
        If txtmadaily.Text = "" Then
            MsgBox("Bạn phải chọn mẫu tin cần xoá.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Note")
            Exit Sub
        Else
            ask = MsgBox("Bạn có chắc muốn xoá mẫu tin này?", MsgBoxStyle.YesNo, "Debt Management")
            If ask = vbYes Then
                Try
                    sqlDelete = "DELETE * FROM Tbl_Daily WHERE MaDL='" & txtmadaily.Text & "'"
                    olecom = New OleDbCommand(sqlDelete, oledbcon)
                    olecom.ExecuteNonQuery()
                    MsgBox("Mẫu tin đã được xoá!!!")
                Catch ex As Exception
                    MsgBox("Error occured when you delete product. Please check again!!!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Note")
                    Exit Sub
                End Try
            End If
        End If
        FillDataset(strSQL)
        BindingTextBox()
        oledbcon.Close()
    End Sub


    Private Sub cmdClosed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClosed.Click
        Me.Close()
    End Sub
    Private Function checkDaily() As Boolean
        Dim dr As DataRow
        For Each dr In ds.Tables("Tbl_Daily").Rows
            If dr("MaDL") = txtmadaily.Text Then
                Return False
            End If
        Next dr
        Return True
    End Function

End Class
