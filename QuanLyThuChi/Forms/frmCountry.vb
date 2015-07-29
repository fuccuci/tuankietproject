Imports System.Data.OleDb
Public Class frmCountry
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
    Friend WithEvents dgridDaily As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txttendaily As System.Windows.Forms.TextBox
    Friend WithEvents txtmadaily As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdSaveEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdClosed As System.Windows.Forms.Button
    Friend WithEvents cmdEditEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdNewEmployee As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCountry))
        Me.dgridDaily = New System.Windows.Forms.DataGrid
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txttendaily = New System.Windows.Forms.TextBox
        Me.txtmadaily = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cmdSaveEmployee = New System.Windows.Forms.Button
        Me.cmdClosed = New System.Windows.Forms.Button
        Me.cmdEditEmployee = New System.Windows.Forms.Button
        Me.cmdNewEmployee = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        CType(Me.dgridDaily, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgridDaily
        '
        Me.dgridDaily.DataMember = ""
        Me.dgridDaily.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgridDaily.Location = New System.Drawing.Point(6, 132)
        Me.dgridDaily.Name = "dgridDaily"
        Me.dgridDaily.ReadOnly = True
        Me.dgridDaily.Size = New System.Drawing.Size(434, 116)
        Me.dgridDaily.TabIndex = 72
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txttendaily)
        Me.GroupBox1.Controls.Add(Me.txtmadaily)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 48)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(424, 80)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txttendaily
        '
        Me.txttendaily.BackColor = System.Drawing.Color.White
        Me.txttendaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendaily.Location = New System.Drawing.Point(86, 45)
        Me.txttendaily.Name = "txttendaily"
        Me.txttendaily.Size = New System.Drawing.Size(330, 26)
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
        Me.txtmadaily.Size = New System.Drawing.Size(122, 26)
        Me.txtmadaily.TabIndex = 1
        Me.txtmadaily.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(10, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Tên KV"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(10, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Mã KV"
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
        Me.GroupBox2.Location = New System.Drawing.Point(8, 242)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(432, 48)
        Me.GroupBox2.TabIndex = 71
        Me.GroupBox2.TabStop = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Font = New System.Drawing.Font("Times New Roman", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.Location = New System.Drawing.Point(266, 15)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(72, 26)
        Me.cmdDelete.TabIndex = 18
        Me.cmdDelete.Text = "Xóa"
        '
        'cmdSaveEmployee
        '
        Me.cmdSaveEmployee.Font = New System.Drawing.Font("Times New Roman", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSaveEmployee.Location = New System.Drawing.Point(180, 14)
        Me.cmdSaveEmployee.Name = "cmdSaveEmployee"
        Me.cmdSaveEmployee.Size = New System.Drawing.Size(72, 26)
        Me.cmdSaveEmployee.TabIndex = 13
        Me.cmdSaveEmployee.Text = "Lưu"
        '
        'cmdClosed
        '
        Me.cmdClosed.Font = New System.Drawing.Font("Times New Roman", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClosed.Location = New System.Drawing.Point(352, 15)
        Me.cmdClosed.Name = "cmdClosed"
        Me.cmdClosed.Size = New System.Drawing.Size(72, 26)
        Me.cmdClosed.TabIndex = 16
        Me.cmdClosed.Text = "Đóng"
        '
        'cmdEditEmployee
        '
        Me.cmdEditEmployee.Font = New System.Drawing.Font("Times New Roman", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEditEmployee.Location = New System.Drawing.Point(94, 14)
        Me.cmdEditEmployee.Name = "cmdEditEmployee"
        Me.cmdEditEmployee.Size = New System.Drawing.Size(72, 26)
        Me.cmdEditEmployee.TabIndex = 13
        Me.cmdEditEmployee.Text = "Thay Đổi"
        '
        'cmdNewEmployee
        '
        Me.cmdNewEmployee.Font = New System.Drawing.Font("Times New Roman", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNewEmployee.Location = New System.Drawing.Point(8, 14)
        Me.cmdNewEmployee.Name = "cmdNewEmployee"
        Me.cmdNewEmployee.Size = New System.Drawing.Size(72, 26)
        Me.cmdNewEmployee.TabIndex = 17
        Me.cmdNewEmployee.Text = "Nhập Mới"
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
        Me.Label9.Location = New System.Drawing.Point(-20, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(462, 48)
        Me.Label9.TabIndex = 70
        Me.Label9.Text = "TỈNH - THÀNH PHỐ"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'frmCountry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(442, 296)
        Me.Controls.Add(Me.dgridDaily)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label9)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmCountry"
        Me.Text = "Khai báo Tỉnh - TP"
        CType(Me.dgridDaily, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FormatDataGridDaily()

        With dgridDaily
            .AllowNavigation = False
            .DataMember = "Tbl_Countries"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách các khu vực"
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "Tbl_Countries"
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
                    .MappingName = "CountryCode"
                    .HeaderText = "Mã khu vực  "
                    .Width = 100
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "CountryName"
                    .HeaderText = "Tên khu vực"
                    .Width = 300
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
            da.Fill(ds, "Tbl_Countries")
            dgridDaily.DataSource = ds.Tables("Tbl_Countries")
        Catch ex As Exception
            MsgBox("Error Fill DataSet :" & ex.ToString)
        End Try
    End Sub

    Private Sub frmDaily_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormatDataGridDaily()
        strSQL = "SELECT * FROM Tbl_Countries ORDER BY CountryCode"
        FillDataset(strSQL)
        BindingTextBox()
        cmdNewEmployee.Focus()
        txtmadaily.ReadOnly = True
        txttendaily.ReadOnly = True
        'flag = "_Add"
    End Sub

    Private Sub cmdNewEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewEmployee.Click
        txtmadaily.ReadOnly = False
        txttendaily.ReadOnly = False
        txtmadaily.Text = ""
        txttendaily.Text = ""

        txtmadaily.Focus()
        flag = "_Add"
        cmdNewEmployee.Enabled = False
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
            cmdSaveEmployee.Focus()
        End If
    End Sub


    Private Sub SaveDataSource(ByVal strMaDL As String, ByVal strTenDL As String)
        Dim sql As String
        If flag = "_Add" Then
            sql = "INSERT INTO Tbl_Countries(CountryCode,CountryName) VALUES('" & strMaDL & "','" & strTenDL & "')"
        Else
            sql = "UPDATE Tbl_Countries SET CountryCode='" & strMaDL & "' , CountryName='" & strTenDL & "'  WHERE CountryCode='" & varMaDL & "'"
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
            MsgBox("Mã khu vực không được rỗng!!!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Chú ý")
            Exit Sub
        End If
        If flag = "_Add" Then
            SaveDataSource(Trim(txtmadaily.Text), txttendaily.Text)
        Else
            SaveDataSource(txtmadaily.Text, txttendaily.Text)
        End If

        cmdNewEmployee.Enabled = True
        cmdEditEmployee.Enabled = True
        cmdNewEmployee.Focus()
        flag = ""
    End Sub

    Private Sub cmdEditEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditEmployee.Click
        flag = "_Edit"
        varMaDL = txtmadaily.Text
        txttendaily.ReadOnly = False
        txtmadaily.Focus()
    End Sub
    Private Sub BindingTextBox()
        Try
            txtmadaily.DataBindings.Clear()
            txttendaily.DataBindings.Clear()

            txtmadaily.DataBindings.Add("Text", ds.Tables("Tbl_Countries"), "CountryCode")
            txttendaily.DataBindings.Add("Text", ds.Tables("Tbl_Countries"), "CountryName")

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
        Dim dsStation As New DataSet
        Dim daStation As OleDbDataAdapter
        Dim i
        oledbcon.Open()
        daStation = New OleDbDataAdapter("SELECT * FROM Tbl_Stations", oledbcon)
        daStation.Fill(dsStation)
        If txtmadaily.Text = "" Then
            MsgBox("Bạn phải chọn mẫu tin cần xoá.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Note")
            Exit Sub
        Else
            ask = MsgBox("Bạn có chắc muốn xoá mẫu tin này? Nó có thể ảnh hưởng đến các bảng khác.", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Debt Management")
            If ask = vbYes Then
                Try
                    For i = 0 To dsStation.Tables(0).Rows.Count - 1
                        sqlDelete = "DELETE * FROM Tbl_Stations WHERE CountryCode='" & txtmadaily.Text & "'"
                        olecom = New OleDbCommand(sqlDelete, oledbcon)
                        olecom.ExecuteNonQuery()
                    Next
                    sqlDelete = "DELETE * FROM Tbl_Countries WHERE CountryCode='" & txtmadaily.Text & "'"
                    olecom = New OleDbCommand(sqlDelete, oledbcon)
                    olecom.ExecuteNonQuery()
                    MsgBox("Mẫu tin đã được xoá!!!")
                Catch ex As Exception
                    MsgBox("Error occured when you delete Country. Please check again!!!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Note")
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

End Class
