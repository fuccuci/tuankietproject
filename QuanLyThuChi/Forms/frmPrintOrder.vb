Imports ConvertNumberToChar
Public Class frmPrintOrder
    Inherits System.Windows.Forms.Form
    Dim ds As New DataSet
    Dim rpt As New CrystalReportOrder
    Dim Arrvar(11) As String
    Dim Arrval(11) As String
    Dim numbers As New ConvertNumbersToString
    Dim splitn As New SplitNumbers
    Friend WithEvents ThePrintDocument As System.Drawing.Printing.PrintDocument
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtHoTenKH As System.Windows.Forms.TextBox
    Friend WithEvents lblTenKH As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents cmdPreView As System.Windows.Forms.Button
    Friend WithEvents cmdInHD As System.Windows.Forms.Button
    Friend WithEvents txtdiachi As System.Windows.Forms.TextBox
    Friend WithEvents txtmst As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNoiDung As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerchuky As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtMaKH As System.Windows.Forms.TextBox
    Friend WithEvents txtTien As System.Windows.Forms.TextBox
    Friend WithEvents txtSoMay As System.Windows.Forms.TextBox
    Friend WithEvents txtVAT As System.Windows.Forms.TextBox
    Friend WithEvents txtTongTien As System.Windows.Forms.TextBox
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DateTimePickerchuky = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtMaKH = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtNoiDung = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtmst = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtTien = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtdiachi = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtHoTenKH = New System.Windows.Forms.TextBox
        Me.lblTenKH = New System.Windows.Forms.Label
        Me.txtSoMay = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.cmdInHD = New System.Windows.Forms.Button
        Me.cmdPreView = New System.Windows.Forms.Button
        Me.cmddong = New System.Windows.Forms.Button
        Me.txtVAT = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtTongTien = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DateTimePickerchuky)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtMaKH)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtNoiDung)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtmst)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtTien)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtdiachi)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtHoTenKH)
        Me.GroupBox1.Controls.Add(Me.lblTenKH)
        Me.GroupBox1.Controls.Add(Me.txtSoMay)
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Controls.Add(Me.cmdInHD)
        Me.GroupBox1.Controls.Add(Me.cmdPreView)
        Me.GroupBox1.Controls.Add(Me.cmddong)
        Me.GroupBox1.Controls.Add(Me.txtVAT)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtTongTien)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(522, 296)
        Me.GroupBox1.TabIndex = 40
        Me.GroupBox1.TabStop = False
        '
        'DateTimePickerchuky
        '
        Me.DateTimePickerchuky.CustomFormat = "MM/yyyy"
        Me.DateTimePickerchuky.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerchuky.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerchuky.Location = New System.Drawing.Point(342, 47)
        Me.DateTimePickerchuky.Name = "DateTimePickerchuky"
        Me.DateTimePickerchuky.Size = New System.Drawing.Size(99, 26)
        Me.DateTimePickerchuky.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(240, 45)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 23)
        Me.Label3.TabIndex = 65
        Me.Label3.Text = "Kỳ Cước"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMaKH
        '
        Me.txtMaKH.BackColor = System.Drawing.Color.White
        Me.txtMaKH.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaKH.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.txtMaKH.Location = New System.Drawing.Point(112, 47)
        Me.txtMaKH.Name = "txtMaKH"
        Me.txtMaKH.Size = New System.Drawing.Size(120, 26)
        Me.txtMaKH.TabIndex = 3
        Me.txtMaKH.Text = "HCM229800"
        Me.txtMaKH.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(8, 45)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 23)
        Me.Label4.TabIndex = 63
        Me.Label4.Text = "Mã Số KH"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtNoiDung
        '
        Me.txtNoiDung.BackColor = System.Drawing.Color.White
        Me.txtNoiDung.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNoiDung.ForeColor = System.Drawing.Color.Black
        Me.txtNoiDung.Location = New System.Drawing.Point(112, 146)
        Me.txtNoiDung.Name = "txtNoiDung"
        Me.txtNoiDung.Size = New System.Drawing.Size(400, 26)
        Me.txtNoiDung.TabIndex = 7
        Me.txtNoiDung.Text = "Dịch vụ viễn thông Viettel"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(8, 145)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 23)
        Me.Label6.TabIndex = 61
        Me.Label6.Text = "Nội Dung "
        '
        'txtmst
        '
        Me.txtmst.BackColor = System.Drawing.Color.White
        Me.txtmst.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmst.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.txtmst.Location = New System.Drawing.Point(341, 16)
        Me.txtmst.Name = "txtmst"
        Me.txtmst.Size = New System.Drawing.Size(171, 26)
        Me.txtmst.TabIndex = 2
        Me.txtmst.Text = "0100109106"
        Me.txtmst.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(240, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 23)
        Me.Label5.TabIndex = 59
        Me.Label5.Text = "Mã Số Thuế "
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTien
        '
        Me.txtTien.AutoSize = False
        Me.txtTien.BackColor = System.Drawing.Color.White
        Me.txtTien.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTien.ForeColor = System.Drawing.Color.Chocolate
        Me.txtTien.Location = New System.Drawing.Point(112, 179)
        Me.txtTien.Name = "txtTien"
        Me.txtTien.Size = New System.Drawing.Size(120, 26)
        Me.txtTien.TabIndex = 8
        Me.txtTien.Text = ""
        Me.txtTien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(8, 179)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 23)
        Me.Label2.TabIndex = 48
        Me.Label2.Text = "Tiền trước thuế :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtdiachi
        '
        Me.txtdiachi.BackColor = System.Drawing.Color.White
        Me.txtdiachi.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdiachi.ForeColor = System.Drawing.Color.Black
        Me.txtdiachi.Location = New System.Drawing.Point(112, 113)
        Me.txtdiachi.Name = "txtdiachi"
        Me.txtdiachi.Size = New System.Drawing.Size(400, 26)
        Me.txtdiachi.TabIndex = 6
        Me.txtdiachi.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(8, 113)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 23)
        Me.Label1.TabIndex = 46
        Me.Label1.Text = "Địa Chỉ "
        '
        'txtHoTenKH
        '
        Me.txtHoTenKH.BackColor = System.Drawing.Color.White
        Me.txtHoTenKH.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHoTenKH.ForeColor = System.Drawing.Color.Black
        Me.txtHoTenKH.Location = New System.Drawing.Point(112, 80)
        Me.txtHoTenKH.Name = "txtHoTenKH"
        Me.txtHoTenKH.Size = New System.Drawing.Size(400, 26)
        Me.txtHoTenKH.TabIndex = 5
        Me.txtHoTenKH.Text = ""
        '
        'lblTenKH
        '
        Me.lblTenKH.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTenKH.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lblTenKH.Location = New System.Drawing.Point(8, 81)
        Me.lblTenKH.Name = "lblTenKH"
        Me.lblTenKH.Size = New System.Drawing.Size(96, 23)
        Me.lblTenKH.TabIndex = 44
        Me.lblTenKH.Text = "Họ và Tên "
        '
        'txtSoMay
        '
        Me.txtSoMay.BackColor = System.Drawing.Color.White
        Me.txtSoMay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoMay.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.txtSoMay.Location = New System.Drawing.Point(112, 14)
        Me.txtSoMay.Name = "txtSoMay"
        Me.txtSoMay.Size = New System.Drawing.Size(120, 26)
        Me.txtSoMay.TabIndex = 1
        Me.txtSoMay.Text = ""
        Me.txtSoMay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label21.Location = New System.Drawing.Point(8, 15)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(105, 23)
        Me.Label21.TabIndex = 42
        Me.Label21.Text = "Số Thuê Bao "
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdInHD
        '
        Me.cmdInHD.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdInHD.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdInHD.Location = New System.Drawing.Point(240, 256)
        Me.cmdInHD.Name = "cmdInHD"
        Me.cmdInHD.Size = New System.Drawing.Size(75, 25)
        Me.cmdInHD.TabIndex = 11
        Me.cmdInHD.Text = "In"
        '
        'cmdPreView
        '
        Me.cmdPreView.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPreView.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdPreView.Location = New System.Drawing.Point(56, 256)
        Me.cmdPreView.Name = "cmdPreView"
        Me.cmdPreView.Size = New System.Drawing.Size(80, 25)
        Me.cmdPreView.TabIndex = 12
        Me.cmdPreView.Text = "Preview"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(433, 256)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(75, 25)
        Me.cmddong.TabIndex = 13
        Me.cmddong.Text = "Đóng"
        '
        'txtVAT
        '
        Me.txtVAT.AutoSize = False
        Me.txtVAT.BackColor = System.Drawing.Color.White
        Me.txtVAT.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVAT.ForeColor = System.Drawing.Color.Chocolate
        Me.txtVAT.Location = New System.Drawing.Point(349, 176)
        Me.txtVAT.Name = "txtVAT"
        Me.txtVAT.Size = New System.Drawing.Size(120, 26)
        Me.txtVAT.TabIndex = 9
        Me.txtVAT.Text = ""
        Me.txtVAT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label7.Location = New System.Drawing.Point(247, 180)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 23)
        Me.Label7.TabIndex = 48
        Me.Label7.Text = "Tiền  VAT :"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTongTien
        '
        Me.txtTongTien.AutoSize = False
        Me.txtTongTien.BackColor = System.Drawing.Color.White
        Me.txtTongTien.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTongTien.ForeColor = System.Drawing.Color.Chocolate
        Me.txtTongTien.Location = New System.Drawing.Point(112, 210)
        Me.txtTongTien.Name = "txtTongTien"
        Me.txtTongTien.Size = New System.Drawing.Size(120, 26)
        Me.txtTongTien.TabIndex = 10
        Me.txtTongTien.Text = ""
        Me.txtTongTien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(8, 210)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(88, 23)
        Me.Label8.TabIndex = 48
        Me.Label8.Text = "Tổng tiền :"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmPrintOrder
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(530, 304)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmPrintOrder"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "In Hoá Đơn ..."
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmPrintOder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Arrvar(0) = "txtHotenKH"
        Arrvar(1) = "txtDiaChi"
        Arrvar(2) = "txtMaKH"
        Arrvar(3) = "txtMST"
        Arrvar(4) = "txtSoMay"
        Arrvar(5) = "txtChukyCuoc"
        Arrvar(6) = "txtTenDV"
        Arrvar(7) = "txtTienDV"
        Arrvar(8) = "txtTienVAT"
        Arrvar(9) = "txtCongTien"
        Arrvar(10) = "txtTongTien"
        Arrvar(11) = "txtTienBangChu"
    End Sub

    Private Sub cmddong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddong.Click
        Me.Close()
    End Sub

    Private Sub cmdPreView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreView.Click
        If (Trim$(txtHoTenKH.Text) = "") Then
            MsgBox("Tên Khách Hàng chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtHoTenKH.Focus()
            Exit Sub
        End If

        If (Trim$(txtdiachi.Text) = "") Then
            MsgBox("Địa chỉ chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtdiachi.Focus()
            Exit Sub
        End If

        If (Trim$(txtTien.Text) = "") Then
            MsgBox("Số tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTien.Focus()
            Exit Sub
        End If

        If (Not IsNumeric(Trim$(txtTien.Text))) Then
            MsgBox("Số tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTien.Focus()
            txtTien.SelectAll()
            Exit Sub
        End If

        If (Trim$(txtVAT.Text) = "") Then
            MsgBox("Tiền VAT chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtVAT.Focus()
            Exit Sub
        End If

        If (Not IsNumeric(Trim$(txtVAT.Text))) Then
            MsgBox("Tiền VAT phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtVAT.Focus()
            txtVAT.SelectAll()
            Exit Sub
        End If


        If (Trim$(txtTongTien.Text) = "") Then
            MsgBox("Tổng Tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTongTien.Focus()
            Exit Sub
        End If

        If (Not IsNumeric(Trim$(txtTongTien.Text))) Then
            MsgBox("Số hóa đơn phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTongTien.Focus()
            txtTongTien.SelectAll()
            Exit Sub
        End If
        AssignVal()
        Dim frm As New frmPreview
        frm.CrystalReportViewerReceipts.ReportSource = rpt
        frm.ShowDialog()
    End Sub
    Private Sub AssignVal()
        Dim strmst As String
        If (Trim$(txtmst.Text) <> "") Then
            If (Len(Trim$(txtmst.Text)) < 10) Then
                MsgBox("MST nhập vào thiếu!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtmst.Focus()
                Exit Sub
            End If
            strmst = txtmst.Text.Chars(0) & " " & txtmst.Text.Chars(1) & "  "
            strmst += txtmst.Text.Chars(2) & " "
            strmst += txtmst.Text.Chars(3) & " "
            strmst += txtmst.Text.Chars(4) & " "
            strmst += txtmst.Text.Chars(5) & " "
            strmst += txtmst.Text.Chars(6) & " "
            strmst += txtmst.Text.Chars(7) & " "
            strmst += txtmst.Text.Chars(8) & "  "
            strmst += txtmst.Text.Chars(9)
            If (Len(Trim$(txtmst.Text)) > 10) Then
                strmst += "            " & txtmst.Text.Chars(10)
            End If
        End If
        
        Arrval(0) = txtHoTenKH.Text
        Arrval(1) = txtdiachi.Text
        Arrval(2) = txtMaKH.Text
        Arrval(3) = strmst
        Arrval(4) = txtSoMay.Text
        Arrval(5) = DateTimePickerchuky.Text
        Arrval(6) = "1 " & txtNoiDung.Text
        Arrval(7) = txtTien.Text
        Arrval(8) = txtVAT.Text
        Arrval(9) = txtTongTien.Text
        Arrval(10) = txtTongTien.Text
        numbers.Number = CLng(txtTongTien.Text)
        Dim str As String
        str = numbers.NumbersToString
        If (Len(str) > 0) Then
            str = str.Substring(0, 1).ToUpper + str.Remove(0, 1)
        Else
            str = "Không"
        End If
        Arrval(11) = str + " đồng."
        SetFieldTextOjectReports(rpt, Arrvar, Arrval)

    End Sub
    Private Sub cmdInHD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInHD.Click

        If (Trim$(txtHoTenKH.Text) = "") Then
            MsgBox("Tên Khách Hàng chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtHoTenKH.Focus()
            Exit Sub
        End If

        If (Trim$(txtmst.Text) <> "") Then
            If (Len(Trim$(txtmst.Text)) < 10) Then
                MsgBox("MST nhập vào thiếu!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtmst.Focus()
                Exit Sub
            End If
        End If

        If (Trim$(txtdiachi.Text) = "") Then
            MsgBox("Địa chỉ chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtdiachi.Focus()
            Exit Sub
        End If

        If (Trim$(txtTien.Text) = "") Then
            MsgBox("Số tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTien.Focus()
            Exit Sub
        End If

        If (Not IsNumeric(Trim$(txtTien.Text))) Then
            MsgBox("Số tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTien.Focus()
            txtTien.SelectAll()
            Exit Sub
        End If

        If (Trim$(txtVAT.Text) = "") Then
            MsgBox("Tiền VAT chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtVAT.Focus()
            Exit Sub
        End If

        If (Not IsNumeric(Trim$(txtVAT.Text))) Then
            MsgBox("Tiền VAT phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtVAT.Focus()
            txtVAT.SelectAll()
            Exit Sub
        End If


        If (Trim$(txtTongTien.Text) = "") Then
            MsgBox("Tổng Tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTongTien.Focus()
            Exit Sub
        End If

        If (Not IsNumeric(Trim$(txtTongTien.Text))) Then
            MsgBox("Số hóa đơn phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtTongTien.Focus()
            txtTongTien.SelectAll()
            Exit Sub
        End If

        AssignVal()

        Dim strprinter
        Dim value = MsgBox("Bạn Có Thật Sự Muốn in Hóa Đơn Không?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "In")
        If (value = vbYes) Then
            Try
                Dim val
                PrintDialog1.Document = ThePrintDocument
                val = PrintDialog1.ShowDialog()
                If (val = vbOK) Then
                    strprinter = PrintDialog1.PrinterSettings.PrinterName
                    rpt.PrintOptions.PrinterName = strprinter
                    rpt.PrintToPrinter(1, True, 1, 1)
                End If
            Catch ex As Exception
                MsgBox("Không Tìm Thấy Máy In " & strprinter, MsgBoxStyle.Critical, "Lỗi In Hoá Đơn")
            End Try
        End If
    End Sub

    Private Sub SetFieldTextOjectReports(ByRef rpt As CrystalReportOrder, ByVal ArrVar As String(), ByVal ArrVal As String())
        Dim ReportTextObject As CrystalDecisions.CrystalReports.Engine.TextObject
        Dim i As Integer
        Try
            For i = 0 To ArrVar.Length - 1
                ReportTextObject = CType(rpt.ReportDefinition.ReportObjects.Item(ArrVar(i)), CrystalDecisions.CrystalReports.Engine.TextObject)
                ReportTextObject.Text = ArrVal(i)
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTien_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTien.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (Trim$(txtTien.Text) = "") Then
                MsgBox("Số tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtTien.Focus()
                Exit Sub
            End If

            If (Not IsNumeric(Trim$(txtTien.Text))) Then
                MsgBox("Số tiền phải là kiểu số!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                txtTien.Focus()
                txtTien.SelectAll()
                Exit Sub
            End If
            splitn.strnumbers = CStr(CLng(txtTien.Text))
            txtTien.Text = splitn.Splitnumer(",")

            txtVAT.Text = CLng(txtTien.Text) * 0.1
            splitn.strnumbers = CStr(CLng(txtVAT.Text))
            txtVAT.Text = splitn.Splitnumer(",")

            txtTongTien.Text = CLng(txtTien.Text) + CLng(txtVAT.Text)
            splitn.strnumbers = CStr(CLng(txtTongTien.Text))
            txtTongTien.Text = splitn.Splitnumer(",")
            cmdInHD.Focus()
        End If
    End Sub

    Private Sub txtSoMay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSoMay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtdiachi_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdiachi.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtHoTenKH_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHoTenKH.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtMaKH_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaKH.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtmst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmst.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtNoiDung_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNoiDung.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub DateTimePickerchuky_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerchuky.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub
End Class
