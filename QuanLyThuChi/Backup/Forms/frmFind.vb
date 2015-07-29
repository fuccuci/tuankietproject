Imports System.Data.OleDb
Public Class frmFind
    Inherits System.Windows.Forms.Form
    Dim strFind As String
    Dim mydataset As DataSet
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        FormatGridWithBothTableAndColumnStyles()
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
    Friend WithEvents cmdthoat As System.Windows.Forms.Button
    Friend WithEvents cmdtim As System.Windows.Forms.Button
    Friend WithEvents lbloai As System.Windows.Forms.Label
    Friend WithEvents grpTimKiem As System.Windows.Forms.GroupBox
    Friend WithEvents rdtenduong As System.Windows.Forms.RadioButton
    Friend WithEvents rdsdt As System.Windows.Forms.RadioButton
    Friend WithEvents RdoMa As System.Windows.Forms.RadioButton
    Friend WithEvents rdten As System.Windows.Forms.RadioButton
    Friend WithEvents DataGridDskh As System.Windows.Forms.DataGrid
    Friend WithEvents txtloai As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboQuan As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblsoKH As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdthoat = New System.Windows.Forms.Button
        Me.cmdtim = New System.Windows.Forms.Button
        Me.lbloai = New System.Windows.Forms.Label
        Me.txtloai = New System.Windows.Forms.TextBox
        Me.grpTimKiem = New System.Windows.Forms.GroupBox
        Me.rdtenduong = New System.Windows.Forms.RadioButton
        Me.rdsdt = New System.Windows.Forms.RadioButton
        Me.RdoMa = New System.Windows.Forms.RadioButton
        Me.rdten = New System.Windows.Forms.RadioButton
        Me.DataGridDskh = New System.Windows.Forms.DataGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboQuan = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblsoKH = New System.Windows.Forms.Label
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.grpTimKiem.SuspendLayout()
        CType(Me.DataGridDskh, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdthoat
        '
        Me.cmdthoat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdthoat.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdthoat.Location = New System.Drawing.Point(736, 520)
        Me.cmdthoat.Name = "cmdthoat"
        Me.cmdthoat.Size = New System.Drawing.Size(72, 27)
        Me.cmdthoat.TabIndex = 11
        Me.cmdthoat.Text = "&Đóng"
        '
        'cmdtim
        '
        Me.cmdtim.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdtim.Location = New System.Drawing.Point(612, 156)
        Me.cmdtim.Name = "cmdtim"
        Me.cmdtim.Size = New System.Drawing.Size(72, 27)
        Me.cmdtim.TabIndex = 3
        Me.cmdtim.Text = "&Tìm ..."
        '
        'lbloai
        '
        Me.lbloai.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbloai.Location = New System.Drawing.Point(266, 129)
        Me.lbloai.Name = "lbloai"
        Me.lbloai.Size = New System.Drawing.Size(272, 24)
        Me.lbloai.TabIndex = 8
        Me.lbloai.Text = "hh"
        '
        'txtloai
        '
        Me.txtloai.Font = New System.Drawing.Font(".VnArialH", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtloai.Location = New System.Drawing.Point(261, 156)
        Me.txtloai.Name = "txtloai"
        Me.txtloai.Size = New System.Drawing.Size(346, 25)
        Me.txtloai.TabIndex = 2
        Me.txtloai.Text = ""
        '
        'grpTimKiem
        '
        Me.grpTimKiem.Controls.Add(Me.rdtenduong)
        Me.grpTimKiem.Controls.Add(Me.rdsdt)
        Me.grpTimKiem.Controls.Add(Me.RdoMa)
        Me.grpTimKiem.Controls.Add(Me.rdten)
        Me.grpTimKiem.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpTimKiem.Location = New System.Drawing.Point(4, 9)
        Me.grpTimKiem.Name = "grpTimKiem"
        Me.grpTimKiem.Size = New System.Drawing.Size(252, 175)
        Me.grpTimKiem.TabIndex = 7
        Me.grpTimKiem.TabStop = False
        Me.grpTimKiem.Text = "Tìm Kiếm Khách Hàng Theo :"
        '
        'rdtenduong
        '
        Me.rdtenduong.Location = New System.Drawing.Point(16, 127)
        Me.rdtenduong.Name = "rdtenduong"
        Me.rdtenduong.Size = New System.Drawing.Size(181, 32)
        Me.rdtenduong.TabIndex = 3
        Me.rdtenduong.Text = "Điạ Chỉ Khách Hàng"
        '
        'rdsdt
        '
        Me.rdsdt.Location = New System.Drawing.Point(19, 91)
        Me.rdsdt.Name = "rdsdt"
        Me.rdsdt.Size = New System.Drawing.Size(136, 32)
        Me.rdsdt.TabIndex = 2
        Me.rdsdt.Text = "Số Điện Thoại"
        '
        'RdoMa
        '
        Me.RdoMa.Location = New System.Drawing.Point(19, 55)
        Me.RdoMa.Name = "RdoMa"
        Me.RdoMa.Size = New System.Drawing.Size(136, 32)
        Me.RdoMa.TabIndex = 1
        Me.RdoMa.Text = "Mã Khách Hàng"
        '
        'rdten
        '
        Me.rdten.Checked = True
        Me.rdten.Location = New System.Drawing.Point(19, 19)
        Me.rdten.Name = "rdten"
        Me.rdten.Size = New System.Drawing.Size(136, 32)
        Me.rdten.TabIndex = 0
        Me.rdten.TabStop = True
        Me.rdten.Text = "Tên Khách Hàng"
        '
        'DataGridDskh
        '
        Me.DataGridDskh.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridDskh.CaptionText = "Danh sách khách hàng"
        Me.DataGridDskh.DataMember = ""
        Me.DataGridDskh.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridDskh.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridDskh.Location = New System.Drawing.Point(8, 186)
        Me.DataGridDskh.Name = "DataGridDskh"
        Me.DataGridDskh.Size = New System.Drawing.Size(800, 326)
        Me.DataGridDskh.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.SystemColors.Info
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.Location = New System.Drawing.Point(267, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(512, 80)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "CÔNG TY VIỄN THÔNG QUÂN ĐỘI  BỘ PHẬN CHĂM SÓC KHÁCH HÀNG"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(267, 103)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 24)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Quận - Huyện"
        '
        'cboQuan
        '
        Me.cboQuan.Location = New System.Drawing.Point(368, 102)
        Me.cboQuan.Name = "cboQuan"
        Me.cboQuan.Size = New System.Drawing.Size(121, 23)
        Me.cboQuan.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label3.Location = New System.Drawing.Point(24, 520)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(176, 24)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Số khách hàng tìm thấy : "
        '
        'lblsoKH
        '
        Me.lblsoKH.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblsoKH.Font = New System.Drawing.Font("Times New Roman", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblsoKH.ForeColor = System.Drawing.Color.IndianRed
        Me.lblsoKH.Location = New System.Drawing.Point(192, 520)
        Me.lblsoKH.Name = "lblsoKH"
        Me.lblsoKH.Size = New System.Drawing.Size(53, 24)
        Me.lblsoKH.TabIndex = 25
        Me.lblsoKH.Text = "0"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.BackColor = System.Drawing.SystemColors.Info
        Me.LinkLabel1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.LinkLabel1.ForeColor = System.Drawing.Color.Transparent
        Me.LinkLabel1.Location = New System.Drawing.Point(667, 80)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(112, 16)
        Me.LinkLabel1.TabIndex = 27
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Contact to Author"
        '
        'frmFind
        '
        Me.AcceptButton = Me.cmdtim
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(810, 552)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.lblsoKH)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboQuan)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridDskh)
        Me.Controls.Add(Me.cmdthoat)
        Me.Controls.Add(Me.cmdtim)
        Me.Controls.Add(Me.lbloai)
        Me.Controls.Add(Me.txtloai)
        Me.Controls.Add(Me.grpTimKiem)
        Me.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmFind"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tìm Kiếm Khách Hàng"
        Me.grpTimKiem.ResumeLayout(False)
        CType(Me.DataGridDskh, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdthoat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdthoat.Click
        Me.Close()
        'Application.Exit()
    End Sub

    Private Sub rdten_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdten.CheckedChanged
        lbloai.Text = "Nhập Vào Tên Khách Hàng:"
    End Sub

    Private Sub RdoMa_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdoMa.CheckedChanged
        lbloai.Text = "Nhập Vào Mã Khách Hàng:"
    End Sub

    Private Sub txtloai_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtloai.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub rdsdt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdsdt.CheckedChanged
        lbloai.Text = "Nhập Vào SĐT Khách Hàng:"
    End Sub

    Private Sub rdtenduong_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdtenduong.CheckedChanged
        lbloai.Text = "Nhập Vào Địa Chỉ Khách Hàng:"
    End Sub

    Private Sub cmdtim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdtim.Click
        strSQL = "SELECT CUST_ID,ISDN,Cust_NAME,ADDRESS,CONTRACT_NO,TYPE_NAME FROM Tbl_Customers,Cust_Type WHERE Tbl_Customers.SUB_TYPE= Cust_Type.BUS_TYPE  AND "
        strFind = Trim$(txtloai.Text)
        If (rdten.Checked) Then
            If (txtloai.Text = "") Then
                MsgBox("Bạn Chưa Nhập Vào Tên Khách Hàng", MsgBoxStyle.Critical, "Tìm Kiếm.")
                txtloai.Focus()
                Exit Sub
            Else
                strSQL += " Cust_NAME LIKE '%" + strFind + "%'"
                If (Trim$(cboQuan.Text) <> "") Then
                    strSQL += " AND DISTRICT = '" + Trim$(cboQuan.Text) + "'"
                End If
                Load_Info()
                Exit Sub
            End If
        End If

        If (RdoMa.Checked) Then
            If (txtloai.Text = "") Then
                MsgBox("Bạn Chưa Nhập Vào Mã Khách Hàng" & vbCrLf & "Hoặc Mã KH Bạn Nhập Vào Không Hợp Lệ.", MsgBoxStyle.Critical, "Tìm Kiếm.")
                txtloai.Focus()
                Exit Sub
            Else
                strSQL += " Cust_ID =" + strFind
                Load_Info()
                Exit Sub
            End If
        End If

        If (rdsdt.Checked) Then
            If ((txtloai.Text = "") OrElse Not IsNumeric(Trim$(txtloai.Text))) Then
                MsgBox("Bạn Chưa Nhập Vào Số Điện Thoại" & vbCrLf & "Hoặc Số Điện Thoại Bạn Nhập Vào Không Hợp Lệ.", MsgBoxStyle.Critical, "Tìm Kiếm.")
                txtloai.Focus()
                Exit Sub
            Else
                strSQL += " ISDN = '" + strFind + "'"
                Load_Info()
                Exit Sub
            End If
        End If

        If (rdtenduong.Checked) Then
            If (txtloai.Text = "") Then
                MsgBox("Bạn Chưa Nhập Vào Địa Chỉ Khách Hàng", MsgBoxStyle.Critical, "Tìm Kiếm.")
                txtloai.Focus()
                Exit Sub
            Else
                strSQL += " ADDRESS LIKE '%" + strFind + "%'"
                If (Trim$(cboQuan.Text) <> "") Then
                    strSQL += " AND DISTRICT = '" + Trim$(cboQuan.Text) + "'"
                End If
                Load_Info()
            End If
        End If

    End Sub
    Private Sub Load_Info()
        Try
            DataGridDskh.DataSource = Nothing
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            mydataset = New DataSet
            da.Fill(mydataset, "Customers")
            If (mydataset.Tables(0).Rows.Count > 0) Then
                DataGridDskh.DataSource = mydataset.Tables(0)
                lblsoKH.Text = mydataset.Tables(0).Rows.Count
            Else
                MsgBox("Khách Hàng Này Chưa Được Cập Nhật Vào Hệ Thống!", MsgBoxStyle.Critical, "Tìm Kiếm ...")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        
    End Sub
    'DISTRICT
    Private Sub LoadInToCombo()
        strSQL = "SELECT DISTINCT DISTRICT FROM Tbl_Customers "
        Dim cmd As New OleDbCommand(strSQL, oledbcon)
        Dim da1 As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da1.Fill(ds, "Customers")
        Dim i As Integer
        For i = 0 To ds.Tables(0).Rows.Count - 1
            cboQuan.Items.Add(Trim$(ds.Tables(0).Rows(i).Item(0)))
        Next
    End Sub
    Private Sub FormatGridWithBothTableAndColumnStyles()

        With DataGridDskh
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh Sách Khách Hàng "
            .Font = New System.Drawing.Font(".VnArialH", 10.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "Customers"
            .BackColor = System.Drawing.Color.MintCream
            .ForeColor = System.Drawing.Color.Navy
            .GridLineColor = System.Drawing.Color.MediumBlue
            .HeaderBackColor = System.Drawing.Color.Lavender
            .HeaderForeColor = System.Drawing.Color.Navy
            .AlternatingBackColor = Color.LightGray
            .HeaderFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .RowHeaderWidth = 10
            .ReadOnly = True
            ' Set column styles
            With .GridColumnStyles
                ' Set datagrid ColumnStyle for ID field
                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "CUST_ID"
                    .HeaderText = "Mã KH"
                    .Width = 90
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "ISDN"
                    .HeaderText = "Số ĐT"
                    .Width = 80
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Cust_NAME"
                    .HeaderText = "Họ Tên Khách Hàng/Công Ty"
                    .Width = 250
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "ADDRESS"
                    .HeaderText = "Địa Chỉ Khách Hàng"
                    .Width = 300
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "CONTRACT_NO"
                    .HeaderText = "Mã HĐ"
                    .Width = 70
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "TYPE_NAME"
                    .HeaderText = "Loại KH"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With
            End With
        End With
        DataGridDskh.TableStyles.Add(TblStyle)
    End Sub

    Private Sub frmFind_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadInToCombo()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        MsgBox("Chương trình Quản Lý Cộng Tác Viên" & vbCrLf & vbCrLf & "Programmed By : Nguyen Van An." & vbCrLf & vbCrLf & "Email : vannguyenan@yahoo.com" & vbCrLf & "           vanan@vietel.com.vn" & vbCrLf & "Mobile: 0988000891", MsgBoxStyle.Information, "Quản Lý CTV")
    End Sub

    Private Sub cboQuan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboQuan.SelectedIndexChanged

    End Sub
End Class
