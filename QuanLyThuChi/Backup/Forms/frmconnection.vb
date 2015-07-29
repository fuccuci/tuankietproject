Imports System.Net.Dns
Imports System.Data.OleDb
Public Class frmconnection
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private count As Short = 0
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
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
    Public WithEvents frmConnectionInfo As System.Windows.Forms.GroupBox
    Public WithEvents txtDatabaseName As System.Windows.Forms.TextBox
    Public WithEvents txtServerName As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents cmdConnect As System.Windows.Forms.Button
    Public WithEvents cmdDisconnect As System.Windows.Forms.Button
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents txtUserName As System.Windows.Forms.TextBox
    Public WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxLAN As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmconnection))
        Me.frmConnectionInfo = New System.Windows.Forms.GroupBox
        Me.CheckBoxLAN = New System.Windows.Forms.CheckBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmdConnect = New System.Windows.Forms.Button
        Me.cmdDisconnect = New System.Windows.Forms.Button
        Me.txtServerName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.frmConnectionInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'frmConnectionInfo
        '
        Me.frmConnectionInfo.BackColor = System.Drawing.SystemColors.ControlLight
        Me.frmConnectionInfo.Controls.Add(Me.CheckBoxLAN)
        Me.frmConnectionInfo.Controls.Add(Me.txtPassword)
        Me.frmConnectionInfo.Controls.Add(Me.Label4)
        Me.frmConnectionInfo.Controls.Add(Me.txtUserName)
        Me.frmConnectionInfo.Controls.Add(Me.Label3)
        Me.frmConnectionInfo.Controls.Add(Me.cmdConnect)
        Me.frmConnectionInfo.Controls.Add(Me.cmdDisconnect)
        Me.frmConnectionInfo.Controls.Add(Me.txtServerName)
        Me.frmConnectionInfo.Controls.Add(Me.Label1)
        Me.frmConnectionInfo.Controls.Add(Me.PictureBox1)
        Me.frmConnectionInfo.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmConnectionInfo.ForeColor = System.Drawing.SystemColors.Desktop
        Me.frmConnectionInfo.Location = New System.Drawing.Point(6, 0)
        Me.frmConnectionInfo.Name = "frmConnectionInfo"
        Me.frmConnectionInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmConnectionInfo.Size = New System.Drawing.Size(322, 168)
        Me.frmConnectionInfo.TabIndex = 4
        Me.frmConnectionInfo.TabStop = False
        Me.frmConnectionInfo.Text = "System login"
        '
        'CheckBoxLAN
        '
        Me.CheckBoxLAN.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxLAN.ForeColor = System.Drawing.SystemColors.Desktop
        Me.CheckBoxLAN.Location = New System.Drawing.Point(113, 76)
        Me.CheckBoxLAN.Name = "CheckBoxLAN"
        Me.CheckBoxLAN.Size = New System.Drawing.Size(200, 25)
        Me.CheckBoxLAN.TabIndex = 18
        Me.CheckBoxLAN.Text = "Qua mạng LAN/WAN"
        '
        'txtPassword
        '
        Me.txtPassword.AcceptsReturn = True
        Me.txtPassword.AutoSize = False
        Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPassword.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPassword.Location = New System.Drawing.Point(114, 50)
        Me.txtPassword.MaxLength = 0
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPassword.Size = New System.Drawing.Size(200, 26)
        Me.txtPassword.TabIndex = 16
        Me.txtPassword.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(33, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(80, 24)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Mật khẩu"
        '
        'txtUserName
        '
        Me.txtUserName.AcceptsReturn = True
        Me.txtUserName.AutoSize = False
        Me.txtUserName.BackColor = System.Drawing.SystemColors.Window
        Me.txtUserName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUserName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUserName.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUserName.Location = New System.Drawing.Point(114, 20)
        Me.txtUserName.MaxLength = 0
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUserName.Size = New System.Drawing.Size(200, 26)
        Me.txtUserName.TabIndex = 14
        Me.txtUserName.Text = ""
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(4, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(112, 24)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Tên truy nhập"
        '
        'cmdConnect
        '
        Me.cmdConnect.BackColor = System.Drawing.SystemColors.Control
        Me.cmdConnect.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdConnect.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdConnect.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdConnect.Location = New System.Drawing.Point(112, 136)
        Me.cmdConnect.Name = "cmdConnect"
        Me.cmdConnect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdConnect.Size = New System.Drawing.Size(80, 25)
        Me.cmdConnect.TabIndex = 11
        Me.cmdConnect.Text = "Chấp nhận"
        '
        'cmdDisconnect
        '
        Me.cmdDisconnect.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDisconnect.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDisconnect.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdDisconnect.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDisconnect.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdDisconnect.Location = New System.Drawing.Point(232, 136)
        Me.cmdDisconnect.Name = "cmdDisconnect"
        Me.cmdDisconnect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDisconnect.Size = New System.Drawing.Size(80, 25)
        Me.cmdDisconnect.TabIndex = 12
        Me.cmdDisconnect.Text = "Bỏ qua"
        '
        'txtServerName
        '
        Me.txtServerName.AcceptsReturn = True
        Me.txtServerName.AutoSize = False
        Me.txtServerName.BackColor = System.Drawing.SystemColors.Window
        Me.txtServerName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtServerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtServerName.Enabled = False
        Me.txtServerName.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtServerName.Location = New System.Drawing.Point(112, 102)
        Me.txtServerName.MaxLength = 0
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtServerName.Size = New System.Drawing.Size(200, 26)
        Me.txtServerName.TabIndex = 6
        Me.txtServerName.Text = ""
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(37, 102)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(72, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Tên máy"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(7, 114)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(80, 48)
        Me.PictureBox1.TabIndex = 19
        Me.PictureBox1.TabStop = False
        '
        'frmconnection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ClientSize = New System.Drawing.Size(330, 175)
        Me.Controls.Add(Me.frmConnectionInfo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmconnection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.TransparencyKey = System.Drawing.SystemColors.ControlDark
        Me.frmConnectionInfo.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConnect.Click
        Dim Tenmay As String

        If (CheckBoxLAN.Checked) Then
            Tenmay = txtServerName.Text
            If (Trim$(Tenmay) = "") Then
                MsgBox("Chưa nhập vào tên máy cần nối", MsgBoxStyle.Critical)
                txtServerName.Focus()
                Exit Sub
            End If

            con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "\\" & Tenmay & "\Data\QLCTV.mdb ;Jet OLEDB:Database Password='##^^&&**~`!!$-+/%%an';Persist Security Info=False;"
            strinfor = "KẾT NỐI QUA MÁY: " & Tenmay.ToUpper
            Try
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                oledbcon = New OleDbConnection(con)
                oledbcon.Open()
            Catch ex As Exception
                MsgBox("Không tìm thấy máy: " & Tenmay & "!. Vui lòng kiểm tra lại")
                txtServerName.Focus()
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Exit Sub
            End Try
        Else
            con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Data\QLCTV.mdb ;Jet OLEDB:Database Password='##^^&&**~`!!$-+/%%an';Persist Security Info=False;"
            strinfor = "NHẬP LIỆU TẠI MÁY ĐƠN"
            Try
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                oledbcon = New OleDbConnection(con)
                oledbcon.Open()
            Catch ex As Exception
                MsgBox("Không kết nối được CSDL!. Vui lòng kiểm tra lại")
                txtServerName.Focus()
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Exit Sub
            End Try
        End If
        LoadUsername()
        If (CheckUser()) Then
            'MsgBox("Kết nối thành công!")
            Me.Close()
        End If
    End Sub

    Private Sub frmconnection_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If

    End Sub

    Public Sub cmdDisconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDisconnect.Click
        If (oledbcon.State = ConnectionState.Open) Then
            oledbcon.Close()
        End If
        Me.Close()
    End Sub

    Private Sub CheckBoxLAN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLAN.CheckedChanged
        Dim cl As System.Drawing.Color
        If (CheckBoxLAN.Checked) Then
            txtServerName.Enabled = True
            txtServerName.BackColor = cl.White
        Else
            txtServerName.Enabled = False
            txtServerName.BackColor = cl.Gray
        End If
    End Sub

    Private Sub frmconnection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cl As System.Drawing.Color
        txtServerName.BackColor = cl.Gray
        txtUserName.Focus()
    End Sub

    Private Function CheckUser() As Boolean
        Dim bol As Boolean
        Dim dr As DataRow
        bol = False

        If txtUserName.Text = "" Then
            MsgBox("Bạn chưa nhập tên truy cập !", MsgBoxStyle.Critical, "Quản Lý Cộng Tác Viên")
            txtUserName.Focus()
            GoTo kt
        End If

        username = txtUserName.Text

        If txtPassword.Text = "" Then
            MsgBox("Bạn chưa nhập Password !", MsgBoxStyle.Critical, "Quản Lý Cộng Tác Viên")
            txtPassword.Focus()
            GoTo kt
        End If

        password = txtPassword.Text

        For Each dr In mydataset.Tables(0).Rows
            If dr("UserName") = username AndAlso dr("Passwrd") = password Then
                bol = True
                Exit For
            End If
        Next dr

        If bol = False Then
            count += 1
            If count >= 3 Then
                MsgBox("Bạn không có quyền truy cập chương trình này.Bye bye !!!", MsgBoxStyle.Critical, "Quản Lý Cộng Tác Viên")
                Application.Exit()
            Else
                MsgBox("Username - Password không đúng!", MsgBoxStyle.Critical, "Quản Lý Cộng Tác Viên")
                txtUserName.Clear()
                txtPassword.Clear()
                txtUserName.Focus()
            End If
        End If
kt:
        Return bol
    End Function
    Private Sub LoadUsername()
        Try
            strSQL = "SELECT * FROM Tbl_UserPass"
            Dim cmd As New OleDbCommand(strSQL, oledbcon)
            da = New OleDbDataAdapter(cmd)
            mydataset = New DataSet
            da.Fill(mydataset, "Pass")
        Catch ex As Exception
            MsgBox("Có lổi khi tìm tên truy cập")
        End Try
        
    End Sub
End Class
