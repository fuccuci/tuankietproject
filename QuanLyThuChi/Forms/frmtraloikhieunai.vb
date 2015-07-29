Imports System.Data.OleDb
Public Class frmtraloikhieunai
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DateTimePickerNgaytraloi As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents txtnoidungtraloi As System.Windows.Forms.TextBox
    Friend WithEvents cmdChitiet As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdclose As System.Windows.Forms.Button
    Friend WithEvents cmdXoa As System.Windows.Forms.Button
    Friend WithEvents cmdluu As System.Windows.Forms.Button
    Friend WithEvents cbotinhtrang As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtnoidungkhieunai As System.Windows.Forms.TextBox
    Friend WithEvents cboloaikhieunai As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerngaybaonhan As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePickerNgaynhan As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDiachi As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtHotenkh As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSomay As System.Windows.Forms.TextBox
    Friend WithEvents txtSoCV As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdnew As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.DateTimePickerNgaytraloi = New System.Windows.Forms.DateTimePicker
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.txtnoidungtraloi = New System.Windows.Forms.TextBox
        Me.cmdChitiet = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.cmdnew = New System.Windows.Forms.Button
        Me.cmdclose = New System.Windows.Forms.Button
        Me.cmdXoa = New System.Windows.Forms.Button
        Me.cmdluu = New System.Windows.Forms.Button
        Me.cbotinhtrang = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtnoidungkhieunai = New System.Windows.Forms.TextBox
        Me.cboloaikhieunai = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.DateTimePickerngaybaonhan = New System.Windows.Forms.DateTimePicker
        Me.DateTimePickerNgaynhan = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDiachi = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtHotenkh = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSomay = New System.Windows.Forms.TextBox
        Me.txtSoCV = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.cmdChitiet)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.cbotinhtrang)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.cboloaikhieunai)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerngaybaonhan)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerNgaynhan)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtDiachi)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtHotenkh)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtSomay)
        Me.GroupBox1.Controls.Add(Me.txtSoCV)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(613, 405)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DateTimePickerNgaytraloi)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.GroupBox5)
        Me.GroupBox3.Location = New System.Drawing.Point(2, 240)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(608, 120)
        Me.GroupBox3.TabIndex = 11
        Me.GroupBox3.TabStop = False
        '
        'DateTimePickerNgaytraloi
        '
        Me.DateTimePickerNgaytraloi.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerNgaytraloi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerNgaytraloi.Location = New System.Drawing.Point(142, 11)
        Me.DateTimePickerNgaytraloi.Name = "DateTimePickerNgaytraloi"
        Me.DateTimePickerNgaytraloi.Size = New System.Drawing.Size(119, 20)
        Me.DateTimePickerNgaytraloi.TabIndex = 12
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label9.Location = New System.Drawing.Point(7, 7)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(121, 25)
        Me.Label9.TabIndex = 23
        Me.Label9.Text = "Ngày giải quyết:"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.txtnoidungtraloi)
        Me.GroupBox5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox5.Location = New System.Drawing.Point(4, 32)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(600, 80)
        Me.GroupBox5.TabIndex = 13
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Nội dung trả lời"
        '
        'txtnoidungtraloi
        '
        Me.txtnoidungtraloi.AutoSize = False
        Me.txtnoidungtraloi.BackColor = System.Drawing.Color.White
        Me.txtnoidungtraloi.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnoidungtraloi.Location = New System.Drawing.Point(8, 24)
        Me.txtnoidungtraloi.Multiline = True
        Me.txtnoidungtraloi.Name = "txtnoidungtraloi"
        Me.txtnoidungtraloi.Size = New System.Drawing.Size(584, 48)
        Me.txtnoidungtraloi.TabIndex = 14
        Me.txtnoidungtraloi.Text = ""
        '
        'cmdChitiet
        '
        Me.cmdChitiet.Location = New System.Drawing.Point(448, 16)
        Me.cmdChitiet.Name = "cmdChitiet"
        Me.cmdChitiet.Size = New System.Drawing.Size(32, 24)
        Me.cmdChitiet.TabIndex = 23
        Me.cmdChitiet.Text = "..."
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cmdnew)
        Me.GroupBox4.Controls.Add(Me.cmdclose)
        Me.GroupBox4.Controls.Add(Me.cmdXoa)
        Me.GroupBox4.Controls.Add(Me.cmdluu)
        Me.GroupBox4.Location = New System.Drawing.Point(5, 356)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(603, 44)
        Me.GroupBox4.TabIndex = 15
        Me.GroupBox4.TabStop = False
        '
        'cmdnew
        '
        Me.cmdnew.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdnew.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdnew.Location = New System.Drawing.Point(64, 12)
        Me.cmdnew.Name = "cmdnew"
        Me.cmdnew.Size = New System.Drawing.Size(88, 24)
        Me.cmdnew.TabIndex = 21
        Me.cmdnew.Text = "&Tạo Mới"
        '
        'cmdclose
        '
        Me.cmdclose.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdclose.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdclose.Location = New System.Drawing.Point(504, 12)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(88, 24)
        Me.cmdclose.TabIndex = 20
        Me.cmdclose.Text = "&Đóng"
        '
        'cmdXoa
        '
        Me.cmdXoa.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdXoa.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdXoa.Location = New System.Drawing.Point(256, 12)
        Me.cmdXoa.Name = "cmdXoa"
        Me.cmdXoa.Size = New System.Drawing.Size(88, 24)
        Me.cmdXoa.TabIndex = 19
        Me.cmdXoa.Text = "&Xóa"
        '
        'cmdluu
        '
        Me.cmdluu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdluu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdluu.Location = New System.Drawing.Point(160, 12)
        Me.cmdluu.Name = "cmdluu"
        Me.cmdluu.Size = New System.Drawing.Size(88, 24)
        Me.cmdluu.TabIndex = 16
        Me.cmdluu.Text = "&Lưu"
        '
        'cbotinhtrang
        '
        Me.cbotinhtrang.Location = New System.Drawing.Point(354, 126)
        Me.cbotinhtrang.Name = "cbotinhtrang"
        Me.cbotinhtrang.Size = New System.Drawing.Size(128, 21)
        Me.cbotinhtrang.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(248, 126)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 24)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Tình trạng:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtnoidungkhieunai)
        Me.GroupBox2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox2.Location = New System.Drawing.Point(7, 152)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(600, 88)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Nội dung khiếu nại"
        '
        'txtnoidungkhieunai
        '
        Me.txtnoidungkhieunai.AutoSize = False
        Me.txtnoidungkhieunai.BackColor = System.Drawing.Color.White
        Me.txtnoidungkhieunai.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnoidungkhieunai.Location = New System.Drawing.Point(8, 18)
        Me.txtnoidungkhieunai.Multiline = True
        Me.txtnoidungkhieunai.Name = "txtnoidungkhieunai"
        Me.txtnoidungkhieunai.Size = New System.Drawing.Size(584, 61)
        Me.txtnoidungkhieunai.TabIndex = 10
        Me.txtnoidungkhieunai.Text = ""
        '
        'cboloaikhieunai
        '
        Me.cboloaikhieunai.Location = New System.Drawing.Point(120, 125)
        Me.cboloaikhieunai.Name = "cboloaikhieunai"
        Me.cboloaikhieunai.Size = New System.Drawing.Size(128, 21)
        Me.cboloaikhieunai.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label7.Location = New System.Drawing.Point(9, 123)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 24)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Loại khiếu nại:"
        '
        'DateTimePickerngaybaonhan
        '
        Me.DateTimePickerngaybaonhan.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngaybaonhan.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngaybaonhan.Location = New System.Drawing.Point(328, 44)
        Me.DateTimePickerngaybaonhan.Name = "DateTimePickerngaybaonhan"
        Me.DateTimePickerngaybaonhan.Size = New System.Drawing.Size(152, 20)
        Me.DateTimePickerngaybaonhan.TabIndex = 4
        '
        'DateTimePickerNgaynhan
        '
        Me.DateTimePickerNgaynhan.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerNgaynhan.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerNgaynhan.Location = New System.Drawing.Point(120, 44)
        Me.DateTimePickerNgaynhan.Name = "DateTimePickerNgaynhan"
        Me.DateTimePickerNgaynhan.Size = New System.Drawing.Size(96, 20)
        Me.DateTimePickerNgaynhan.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(222, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(114, 24)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Ngày báo nhận:"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(9, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(93, 24)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Ngày Nhận:"
        '
        'txtDiachi
        '
        Me.txtDiachi.AutoSize = False
        Me.txtDiachi.BackColor = System.Drawing.Color.White
        Me.txtDiachi.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDiachi.Location = New System.Drawing.Point(120, 96)
        Me.txtDiachi.Name = "txtDiachi"
        Me.txtDiachi.Size = New System.Drawing.Size(360, 24)
        Me.txtDiachi.TabIndex = 6
        Me.txtDiachi.Text = ""
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(9, 95)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 24)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Địa chỉ:"
        '
        'txtHotenkh
        '
        Me.txtHotenkh.AutoSize = False
        Me.txtHotenkh.BackColor = System.Drawing.Color.White
        Me.txtHotenkh.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHotenkh.Location = New System.Drawing.Point(120, 68)
        Me.txtHotenkh.Name = "txtHotenkh"
        Me.txtHotenkh.Size = New System.Drawing.Size(360, 24)
        Me.txtHotenkh.TabIndex = 4
        Me.txtHotenkh.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(9, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(125, 24)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Tên Khách hàng:"
        '
        'txtSomay
        '
        Me.txtSomay.AutoSize = False
        Me.txtSomay.BackColor = System.Drawing.Color.White
        Me.txtSomay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSomay.ForeColor = System.Drawing.Color.IndianRed
        Me.txtSomay.Location = New System.Drawing.Point(328, 16)
        Me.txtSomay.Name = "txtSomay"
        Me.txtSomay.Size = New System.Drawing.Size(120, 24)
        Me.txtSomay.TabIndex = 2
        Me.txtSomay.Text = ""
        '
        'txtSoCV
        '
        Me.txtSoCV.AutoSize = False
        Me.txtSoCV.BackColor = System.Drawing.Color.White
        Me.txtSoCV.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoCV.ForeColor = System.Drawing.Color.IndianRed
        Me.txtSoCV.Location = New System.Drawing.Point(120, 16)
        Me.txtSoCV.Name = "txtSoCV"
        Me.txtSoCV.Size = New System.Drawing.Size(96, 24)
        Me.txtSoCV.TabIndex = 1
        Me.txtSoCV.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(224, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Số máy:"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(9, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Số công văn:"
        '
        'frmtraloikhieunai
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(624, 414)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmtraloikhieunai"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Trả lời khiếu nại khách hàng"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub LoadDataSet()

        strSQL = "SELECT TypeName,TypeID FROM Tbl_Type_Complain "
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "Tbl_Type_Complain")
            cboloaikhieunai.DataSource = mydataset.Tables("Tbl_Type_Complain").DefaultView
            cboloaikhieunai.DisplayMember = "TypeName"
            cboloaikhieunai.ValueMember = "TypeID"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        strSQL = "SELECT TypeName,TypeID FROM Tbl_TypeStatics "
        Try
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, "StaticsType")
            cbotinhtrang.DataSource = mydataset.Tables("StaticsType").DefaultView
            cbotinhtrang.DisplayMember = "TypeName"
            cbotinhtrang.ValueMember = "TypeID"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Private Sub Delele(ByVal strSoCv As String)
        Dim olecommand As OleDbCommand
        strSQL = " DELETE tbl_Solve WHERE SoCV ='" & strSoCv & "'"
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.CommandText = strSQL
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
            MsgBox("Đã xóa rồi xong!!")
        Catch ex As Exception
            MsgBox("Lổi rồi người ơi !!" & ex.Message)
        End Try
        oledbcon.Close()
    End Sub

    Private Sub SaveSolve()
        Dim strMaLoai As String
        strMaLoai = cboloaikhieunai.SelectedValue
        Dim strMatinhtrang As String
        strMatinhtrang = cbotinhtrang.SelectedValue
        strSQL = " INSERT INTO tbl_Solve(SoCV,ISDN,Customer_Name,Customer_Address,Reci_Date1,Reci_Date2,Content_ComPlain,Explain_Date,Content_Explain,TypeCom,States) " & _
            " VALUES('" & Trim$(txtSoCV.Text) & _
            "','" & Trim$(txtSomay.Text) & "','" & _
            Trim$(txtHotenkh.Text) & "','" & _
            txtDiachi.Text & "','" & _
            DateTimePickerNgaynhan.Value.ToShortDateString & _
            "','" & DateTimePickerngaybaonhan.Value.ToShortDateString & _
            "','" & txtnoidungkhieunai.Text & _
            "','" & DateTimePickerNgaytraloi.Value.ToShortDateString & _
            "','" & txtnoidungtraloi.Text & _
            "','" & strMaLoai & "','" & strMatinhtrang & "')"
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strSQL
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
            MsgBox("Đã lưu vào hệ thống!!")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        oledbcon.Close()
    End Sub

    Private Sub UpdateSolve()
        Dim strMaLoai As String
        strMaLoai = cboloaikhieunai.SelectedValue
        Dim strMatinhtrang As String
        strMatinhtrang = cbotinhtrang.SelectedValue
        strSQL = " UPDATE tbl_Solve SET ISDN ='" & Trim$(txtSomay.Text) & _
         "',Customer_Name ='" & Trim$(txtHotenkh.Text) & _
         "',Customer_Address = '" & txtDiachi.Text & _
         "',Reci_Date1 ='" & DateTimePickerNgaynhan.Value.ToShortDateString & _
         "',Reci_Date2 ='" & DateTimePickerngaybaonhan.Value.ToShortDateString & _
         "',Content_ComPlain ='" & txtnoidungkhieunai.Text & _
         "',Content_Explain ='" & txtnoidungtraloi.Text & _
         "',Explain_Date ='" & DateTimePickerNgaytraloi.Value.ToShortDateString & _
         "',TypeCom ='" & strMaLoai & _
         "',States ='" & strMatinhtrang & _
         "' WHERE SoCV = '" & Trim$(txtSoCV.Text) & "'"
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strSQL
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
            MsgBox("Đã cập nhật vào hệ thống!!")
        Catch ex As Exception
            MsgBox(" Lổi rồi người ơi!!!" & ex.ToString)
        End Try
        oledbcon.Close()
    End Sub
    Private Sub LoadInfo(ByVal strSoCV As String)
        Dim result As Boolean = False
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = " SELECT ISDN,Customer_Name,Customer_Address,Reci_Date1,Reci_Date2,Content_ComPlain, Content_Explain,Explain_Date,TypeCom,States FROM tbl_Solve WHERE SoCV = '" & strSoCV & "'"
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then
                If Not oleread.IsDBNull(0) Then
                    txtSomay.Text = oleread.GetString(0)
                End If

                If Not oleread.IsDBNull(1) Then
                    txtHotenkh.Text = oleread.GetString(1)
                End If

                If Not oleread.IsDBNull(2) Then
                    txtDiachi.Text = oleread.GetString(2)
                End If

                If Not oleread.IsDBNull(3) Then
                    DateTimePickerNgaynhan.Value = CDate(oleread.GetString(3))
                End If

                If Not oleread.IsDBNull(4) Then
                    DateTimePickerngaybaonhan.Value = CDate(oleread.GetString(4))
                End If

                If Not oleread.IsDBNull(5) Then
                    txtnoidungkhieunai.Text = oleread.GetString(5)
                End If

                If Not oleread.IsDBNull(6) Then
                    txtnoidungtraloi.Text = oleread.GetString(6)
                End If

                If Not oleread.IsDBNull(7) Then
                    DateTimePickerNgaytraloi.Value = CDate(oleread.GetString(7))
                End If

                Dim dt As DataTable
                If Not oleread.IsDBNull(8) Then
                    dt = mydataset.Tables("Tbl_Type_Complain")
                    cboloaikhieunai.SelectedIndex = cboloaikhieunai.FindString(GetStringTypeName(oleread.GetString(8), dt))
                End If

                If Not oleread.IsDBNull(9) Then
                    dt = mydataset.Tables("StaticsType")
                    cbotinhtrang.SelectedIndex = cbotinhtrang.FindString(GetStringTypeName(oleread.GetString(9), dt))
                End If
            Else
                MsgBox("Không tìm thấy số công văn này!", MsgBoxStyle.Critical)
                txtSoCV.Focus()
            End If
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :" & ex.ToString)
        End Try
        oledbcon.Close()

    End Sub
    Private Function GetStringTypeName(ByVal strTypeID As String, ByVal table As DataTable) As String
        Dim strResult As String
        Dim i As Integer
        For i = 0 To table.Rows.Count - 1
            If (strTypeID = table.Rows(i).Item("TypeID")) Then
                strResult = table.Rows(i).Item("TypeName")
                Exit For
            End If
        Next
        Return strResult
    End Function
    Private Sub txtSoCV_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSoCV.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Dim value
        If (KeyAscii = 13) Then
            If (Trim$(txtSoCV.Text) <> "") Then
                If (CheckSoCV(Trim$(txtSoCV.Text))) Then
                    value = MsgBox("Số công văn này đã tồn tại. Bạn có muốn sửa lại thông tin của công văn này không?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Có rồi người ơi !!!")
                    If (value = vbYes) Then
                        'LoadInfo(Trim$(txtSoCV.Text))
                        'UnLocktextbox()
                        LoadInfo(Trim$(txtSoCV.Text))
                        cmdluu.Text = "&Cập nhật"
                    Else
                        Exit Sub
                    End If
                End If
            Else
                MsgBox("Số công  văn chưa nhập vào!", MsgBoxStyle.Critical, "Lổi người ơi")
                Exit Sub
            End If
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Function CheckSoCV(ByVal strSoCV As String) As Boolean
        Dim result As Boolean = False
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = " SELECT SoCV FROM tbl_Solve WHERE SoCV = '" & strSoCV & "'"
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then
                result = True
            End If
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :")
        End Try
        oledbcon.Close()
        Return result
    End Function
    Private Sub frmtraloikhieunai_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            mydataset = New DataSet
            LoadDataSet()
        Catch eLoad As System.Exception
            System.Windows.Forms.MessageBox.Show(eLoad.Message)
        End Try
    End Sub
    Private Sub DateTimePickerngaybaonhan_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePickerngaybaonhan.ValueChanged
        If DateTimePickerngaybaonhan.Value < DateTimePickerNgaynhan.Value Then
            DateTimePickerngaybaonhan.Value = DateTimePickerNgaynhan.Value
            MsgBox("Ngày báo nhận không thể nhỏ hơn ngày nhận")
        End If
    End Sub

    Private Sub cmdclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdclose.Click
        Me.Close()
    End Sub

    Private Sub cmdXoa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdXoa.Click
        If (Trim$(txtSomay.Text) <> "") Then
            Dim value = MsgBox("Bạn có thật sự muốn xóa thông tin của công văn này không?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Thông báo!")
            If (value = vbYes) Then
                Delele(Trim$(txtSoCV.Text))
            End If
        Else
            MsgBox("Không có số CV để xóa!", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub cmdChitiet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChitiet.Click
        If (Trim$(txtSomay.Text) <> "") Then
            Dim frm As New frmchitiet(Trim$(txtSomay.Text))
            frm.ShowDialog()
        Else
            MsgBox("Không có số máy để kiểm tra chi tiết")
        End If
    End Sub

    Private Sub DeleteTexbox()
        txtSoCV.Text = vbNullString
        txtSomay.Text = vbNullString
        txtHotenkh.Text = vbNullString
        txtDiachi.Text = vbNullString
        txtnoidungkhieunai.Text = vbNullString
        txtnoidungtraloi.Text = vbNullString
    End Sub

    Private Sub cmdluu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdluu.Click
        If (Trim$(txtSoCV.Text) = "") Then
            MsgBox("Số CV chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtSoCV.Focus()
            Exit Sub
        End If

        If (Trim$(txtSomay.Text) = "") Then
            MsgBox("Số Máy chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtSomay.Focus()
            Exit Sub
        End If

        If (Trim$(txtHotenkh.Text) = "") Then
            MsgBox("Họ tên khách hàng chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtHotenkh.Focus()
            Exit Sub
        End If

        If (Trim$(txtnoidungkhieunai.Text) = "") Then
            MsgBox("Nội dung khiếu nại chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtnoidungkhieunai.Focus()
            Exit Sub
        End If

        If (Trim$(txtnoidungtraloi.Text) = "") Then
            MsgBox("Nội dung trả lời chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
            txtnoidungtraloi.Focus()
            Exit Sub
        End If
        If (cmdluu.Text = "&Lưu") Then
            SaveSolve()
        Else
            UpdateSolve()
        End If
        Locktextbox()
        cmdluu.Enabled = False
    End Sub

    Private Sub DateTimePickerNgaytraloi_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePickerNgaytraloi.ValueChanged
        If DateTimePickerNgaytraloi.Value < DateTimePickerngaybaonhan.Value Then
            MsgBox("Ngày trả lời không thể nhỏ hơn ngày báo nhận!", MsgBoxStyle.Critical, "Thông báo!!!")
        End If
    End Sub

    Private Sub cmdnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdnew.Click
        cmdluu.Enabled = True
        cmdluu.Text = "&Lưu"
        DeleteTexbox()
        UnLocktextbox()
    End Sub
    Private Sub Locktextbox()
        txtSoCV.ReadOnly = True
        txtSomay.ReadOnly = True
        txtHotenkh.ReadOnly = True
        txtDiachi.ReadOnly = True
        txtnoidungkhieunai.ReadOnly = True
        txtnoidungtraloi.ReadOnly = True
    End Sub
    Private Sub UnLocktextbox()
        txtSoCV.ReadOnly = False
        txtSomay.ReadOnly = False
        txtHotenkh.ReadOnly = False
        txtDiachi.ReadOnly = False
        txtnoidungkhieunai.ReadOnly = False
        txtnoidungtraloi.ReadOnly = False
    End Sub

    Private Sub txtSomay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSomay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtHotenkh_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHotenkh.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtDiachi_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDiachi.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub cboloaikhieunai_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboloaikhieunai.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub cbotinhtrang_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbotinhtrang.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub DateTimePickerngaybaonhan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerngaybaonhan.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub DateTimePickerNgaynhan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerNgaynhan.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub DateTimePickerNgaytraloi_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    
End Class
