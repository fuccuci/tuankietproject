Imports System.Data.OleDb
Public Class frmTram
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private ds As DataSet
    Dim flag, varMaDL, strQuery As String
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
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdSaveEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdClosed As System.Windows.Forms.Button
    Friend WithEvents cmdEditEmployee As System.Windows.Forms.Button
    Friend WithEvents cmdNewEmployee As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtdiachitthuongtru As System.Windows.Forms.TextBox
    Friend WithEvents txttendaily As System.Windows.Forms.TextBox
    Friend WithEvents txtmadaily As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbokhuvuc As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmTram))
        Me.dgridDaily = New System.Windows.Forms.DataGrid
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cmdSaveEmployee = New System.Windows.Forms.Button
        Me.cmdClosed = New System.Windows.Forms.Button
        Me.cmdEditEmployee = New System.Windows.Forms.Button
        Me.cmdNewEmployee = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbokhuvuc = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtdiachitthuongtru = New System.Windows.Forms.TextBox
        Me.txttendaily = New System.Windows.Forms.TextBox
        Me.txtmadaily = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.dgridDaily, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgridDaily
        '
        Me.dgridDaily.DataMember = ""
        Me.dgridDaily.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgridDaily.Location = New System.Drawing.Point(8, 152)
        Me.dgridDaily.Name = "dgridDaily"
        Me.dgridDaily.ReadOnly = True
        Me.dgridDaily.Size = New System.Drawing.Size(528, 184)
        Me.dgridDaily.TabIndex = 72
        '
        'Label9
        '
        Me.Label9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label9.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(544, 40)
        Me.Label9.TabIndex = 70
        Me.Label9.Text = "THÔNG TIN TỔ THU CƯỚC"
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
        Me.GroupBox2.Location = New System.Drawing.Point(8, 336)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(528, 48)
        Me.GroupBox2.TabIndex = 71
        Me.GroupBox2.TabStop = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.Location = New System.Drawing.Point(256, 13)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(70, 27)
        Me.cmdDelete.TabIndex = 18
        Me.cmdDelete.Text = "Xóa"
        '
        'cmdSaveEmployee
        '
        Me.cmdSaveEmployee.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSaveEmployee.Location = New System.Drawing.Point(176, 13)
        Me.cmdSaveEmployee.Name = "cmdSaveEmployee"
        Me.cmdSaveEmployee.Size = New System.Drawing.Size(70, 27)
        Me.cmdSaveEmployee.TabIndex = 5
        Me.cmdSaveEmployee.Text = "Lưu"
        '
        'cmdClosed
        '
        Me.cmdClosed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClosed.Location = New System.Drawing.Point(449, 13)
        Me.cmdClosed.Name = "cmdClosed"
        Me.cmdClosed.Size = New System.Drawing.Size(70, 27)
        Me.cmdClosed.TabIndex = 16
        Me.cmdClosed.Text = "Ðóng"
        '
        'cmdEditEmployee
        '
        Me.cmdEditEmployee.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEditEmployee.Location = New System.Drawing.Point(96, 13)
        Me.cmdEditEmployee.Name = "cmdEditEmployee"
        Me.cmdEditEmployee.Size = New System.Drawing.Size(70, 27)
        Me.cmdEditEmployee.TabIndex = 13
        Me.cmdEditEmployee.Text = "Thay đổi"
        '
        'cmdNewEmployee
        '
        Me.cmdNewEmployee.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNewEmployee.Location = New System.Drawing.Point(16, 13)
        Me.cmdNewEmployee.Name = "cmdNewEmployee"
        Me.cmdNewEmployee.Size = New System.Drawing.Size(70, 27)
        Me.cmdNewEmployee.TabIndex = 17
        Me.cmdNewEmployee.Text = "Nhập mới"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbokhuvuc)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtdiachitthuongtru)
        Me.GroupBox1.Controls.Add(Me.txttendaily)
        Me.GroupBox1.Controls.Add(Me.txtmadaily)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(528, 106)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'cbokhuvuc
        '
        Me.cbokhuvuc.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbokhuvuc.Location = New System.Drawing.Point(368, 12)
        Me.cbokhuvuc.Name = "cbokhuvuc"
        Me.cbokhuvuc.Size = New System.Drawing.Size(152, 27)
        Me.cbokhuvuc.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(280, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 24)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Tình/TP"
        '
        'txtdiachitthuongtru
        '
        Me.txtdiachitthuongtru.BackColor = System.Drawing.Color.White
        Me.txtdiachitthuongtru.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdiachitthuongtru.Location = New System.Drawing.Point(86, 75)
        Me.txtdiachitthuongtru.Name = "txtdiachitthuongtru"
        Me.txtdiachitthuongtru.Size = New System.Drawing.Size(434, 26)
        Me.txtdiachitthuongtru.TabIndex = 4
        Me.txtdiachitthuongtru.Text = ""
        '
        'txttendaily
        '
        Me.txttendaily.BackColor = System.Drawing.Color.White
        Me.txttendaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendaily.Location = New System.Drawing.Point(86, 44)
        Me.txttendaily.Name = "txttendaily"
        Me.txttendaily.Size = New System.Drawing.Size(434, 26)
        Me.txttendaily.TabIndex = 3
        Me.txttendaily.Text = ""
        '
        'txtmadaily
        '
        Me.txtmadaily.BackColor = System.Drawing.Color.White
        Me.txtmadaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmadaily.ForeColor = System.Drawing.Color.Blue
        Me.txtmadaily.Location = New System.Drawing.Point(86, 12)
        Me.txtmadaily.Name = "txtmadaily"
        Me.txtmadaily.Size = New System.Drawing.Size(170, 26)
        Me.txtmadaily.TabIndex = 1
        Me.txtmadaily.Text = ""
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(11, 75)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 24)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Ðịa chỉ"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(10, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Tên tổ thu"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(10, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Mã tổ thu"
        '
        'frmTram
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(544, 392)
        Me.Controls.Add(Me.dgridDaily)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmTram"
        Me.Text = "Tổ thu"
        CType(Me.dgridDaily, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FillCombo(ByRef cbo As ComboBox, ByVal strQuery As String, ByVal strTablename As String, ByVal strdislaymember As String, ByVal strValuemember As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            mydataset = New DataSet
            da.Fill(mydataset, strTablename)
            cbo.DataSource = mydataset.Tables(strTablename).DefaultView
            cbo.DisplayMember = strdislaymember
            cbo.ValueMember = strValuemember
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub FormatDataGridDaily()

        With dgridDaily
            .AllowNavigation = False
            .DataMember = "Tbl_Stations"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách các tram thu cước"
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "Tbl_Stations"
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
                    .MappingName = "StationID"
                    .HeaderText = "Mã trạm   "
                    .Width = 100
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "Station_Name"
                    .HeaderText = "Tên trạm"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Station_Address"
                    .HeaderText = "Ðịa chỉ"
                    .Width = 250
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "CountryCode"
                    .HeaderText = "Mã KV"
                    .Width = 70
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
            End With
        End With
        dgridDaily.TableStyles.Add(TblStyle)
    End Sub

    Private Sub FillDataset(ByVal strQuery As String)
        Try
            'ds = Nothing
            ds = New DataSet
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds, "Tbl_Stations")
            dgridDaily.DataSource = ds.Tables("Tbl_Stations")
        Catch ex As Exception
            MsgBox("Error Fill DataSet :" & ex.ToString)
        End Try
    End Sub

    Private Sub frmDaily_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormatDataGridDaily()
        strSQL = "SELECT * FROM Tbl_Stations ORDER BY StationID"
        FillDataset(strSQL)
        BindingTextBox()
        strQuery = "SELECT CountryCode,CountryName FROM Tbl_Countries "
        FillCombo(cbokhuvuc, strQuery, "Tbl_Country", "CountryCode", "CountryName")
        cmdNewEmployee.Focus()
        txttendaily.ReadOnly = True
        txtmadaily.ReadOnly = True
        txtdiachitthuongtru.ReadOnly = True
        cmdSaveEmployee.Enabled = False
        'flag = "_Add"
    End Sub

    Private Sub cmdNewEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewEmployee.Click
        txttendaily.ReadOnly = False
        txtmadaily.ReadOnly = False
        txtdiachitthuongtru.ReadOnly = False
        cmdSaveEmployee.Text = "Lưu"
        cmdSaveEmployee.Enabled = True
        txtmadaily.Text = ""
        txttendaily.Text = ""
        txtdiachitthuongtru.Text = ""
        txtmadaily.Focus()
        flag = "_Add"
        'cmdNewEmployee.Enabled = False
        cmdEditEmployee.Enabled = False
        strQuery = ""
    End Sub

    Private Sub txtmadaily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmadaily.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cbokhuvuc.Focus()
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
    Private Sub SaveDataSource(ByVal strMaDL As String, ByVal strTenDL As String, ByVal strDiachi As String, ByVal strKhuvuc As String)
        Dim sql As String
        If flag = "_Add" Then
            sql = "INSERT INTO Tbl_Stations(StationID,Station_Name,Station_Address,CountryCode) VALUES('" & strMaDL & "','" & strTenDL & "','" & strDiachi & "','" & strKhuvuc & "')"
        Else
            sql = "UPDATE Tbl_Stations SET StationID='" & strMaDL & "' , Station_Name='" & strTenDL & "' , Station_Address='" & strDiachi & "', CountryCode='" & cbokhuvuc.Text & "' WHERE StationID='" & varMaDL & "'"
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
            MsgBox("Mã trạm không được rỗng!!!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Chú ý")
            Exit Sub
        End If
        If flag = "_Add" Then
            SaveDataSource(Trim(txtmadaily.Text), Trim(txttendaily.Text), Trim(txtdiachitthuongtru.Text), Trim(cbokhuvuc.Text))
        Else
            SaveDataSource(Trim(txtmadaily.Text), Trim(txttendaily.Text), Trim(txtdiachitthuongtru.Text), Trim(cbokhuvuc.Text))
        End If
        cmdNewEmployee.Enabled = True
        cmdEditEmployee.Enabled = True
        cmdNewEmployee.Focus()
        flag = ""
    End Sub

    Private Sub cmdEditEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditEmployee.Click
        flag = "_Edit"
        cmdSaveEmployee.Text = "Cập nhật"
        cmdSaveEmployee.Enabled = True
        txttendaily.ReadOnly = False
        txtdiachitthuongtru.ReadOnly = False
        varMaDL = txtmadaily.Text
        txtmadaily.Focus()
    End Sub
    Private Sub BindingTextBox()
        Try
            txtmadaily.DataBindings.Clear()
            txttendaily.DataBindings.Clear()
            txtdiachitthuongtru.DataBindings.Clear()
            cbokhuvuc.DataBindings.Clear()
            txtmadaily.DataBindings.Add("Text", ds.Tables("Tbl_Stations"), "StationID")
            txttendaily.DataBindings.Add("Text", ds.Tables("Tbl_Stations"), "Station_Name")
            txtdiachitthuongtru.DataBindings.Add("Text", ds.Tables("Tbl_Stations"), "Station_Address")
            cbokhuvuc.DataBindings.Add("Text", ds.Tables("Tbl_Stations"), "CountryCode")
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
            MsgBox("Bạn phải chọn mẫu tin cần xóa.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Note")
            Exit Sub
        Else
            ask = MsgBox("Bạn có chắc muốn xóa thông tin này?", MsgBoxStyle.YesNo, "Debt Management")
            If ask = vbYes Then
                Try
                    sqlDelete = "DELETE * FROM Tbl_Stations WHERE StationID='" & txtmadaily.Text & "'"
                    olecom = New OleDbCommand(sqlDelete, oledbcon)
                    olecom.ExecuteNonQuery()
                    MsgBox("Mẫu tin đã được xoá!!!")
                Catch ex As Exception
                    MsgBox("Error occured when you delete Station. Please check again!!!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Note")
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


    Private Sub cbokhuvuc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbokhuvuc.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txttendaily.Focus()
        End If
    End Sub

    
    Private Sub cbokhuvuc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbokhuvuc.SelectedIndexChanged
        strSQL = "SELECT * FROM Tbl_Stations WHERE CountryCode ='" & cbokhuvuc.Text & "' ORDER BY StationID"
        FillDataset(strSQL)
        BindingTextBox()
    End Sub
End Class
