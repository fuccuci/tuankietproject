Imports System.Data.OleDb
Public Class frmNewUser
    Inherits System.Windows.Forms.Form
    Dim CreateUserPass As ClassCreateUser
    Dim namelogin As String
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
    Public WithEvents frmConnectionInfo As System.Windows.Forms.GroupBox
    Public WithEvents txtusername As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblUserName As System.Windows.Forms.Label
    Public WithEvents txtpassword As System.Windows.Forms.TextBox
    Public WithEvents lblPassword As System.Windows.Forms.Label
    Public WithEvents txtconfirmPassword As System.Windows.Forms.TextBox
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdok As System.Windows.Forms.Button
    Friend WithEvents PanelUsers As System.Windows.Forms.Panel
    Friend WithEvents lblusers As System.Windows.Forms.Label
    Friend WithEvents Lvusername As System.Windows.Forms.ListView
    Friend WithEvents ContextMenuManager As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItemXoa As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemsetpass As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNewUser))
        Me.frmConnectionInfo = New System.Windows.Forms.GroupBox
        Me.txtusername = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblUserName = New System.Windows.Forms.Label
        Me.txtpassword = New System.Windows.Forms.TextBox
        Me.lblPassword = New System.Windows.Forms.Label
        Me.txtconfirmPassword = New System.Windows.Forms.TextBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdok = New System.Windows.Forms.Button
        Me.PanelUsers = New System.Windows.Forms.Panel
        Me.Lvusername = New System.Windows.Forms.ListView
        Me.ContextMenuManager = New System.Windows.Forms.ContextMenu
        Me.MenuItemXoa = New System.Windows.Forms.MenuItem
        Me.MenuItemsetpass = New System.Windows.Forms.MenuItem
        Me.lblusers = New System.Windows.Forms.Label
        Me.frmConnectionInfo.SuspendLayout()
        Me.PanelUsers.SuspendLayout()
        Me.SuspendLayout()
        '
        'frmConnectionInfo
        '
        Me.frmConnectionInfo.BackColor = System.Drawing.SystemColors.Control
        Me.frmConnectionInfo.Controls.Add(Me.txtusername)
        Me.frmConnectionInfo.Controls.Add(Me.Label1)
        Me.frmConnectionInfo.Controls.Add(Me.lblUserName)
        Me.frmConnectionInfo.Controls.Add(Me.txtpassword)
        Me.frmConnectionInfo.Controls.Add(Me.lblPassword)
        Me.frmConnectionInfo.Controls.Add(Me.txtconfirmPassword)
        Me.frmConnectionInfo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmConnectionInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmConnectionInfo.Location = New System.Drawing.Point(32, 3)
        Me.frmConnectionInfo.Name = "frmConnectionInfo"
        Me.frmConnectionInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmConnectionInfo.Size = New System.Drawing.Size(368, 117)
        Me.frmConnectionInfo.TabIndex = 6
        Me.frmConnectionInfo.TabStop = False
        Me.frmConnectionInfo.Text = "Thông Tin Người Dùng"
        '
        'txtusername
        '
        Me.txtusername.AcceptsReturn = True
        Me.txtusername.AutoSize = False
        Me.txtusername.BackColor = System.Drawing.SystemColors.Window
        Me.txtusername.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtusername.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtusername.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtusername.Location = New System.Drawing.Point(161, 17)
        Me.txtusername.MaxLength = 0
        Me.txtusername.Name = "txtusername"
        Me.txtusername.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtusername.Size = New System.Drawing.Size(203, 26)
        Me.txtusername.TabIndex = 6
        Me.txtusername.Text = ""
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(142, 20)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Tên Người Dùng:"
        '
        'lblUserName
        '
        Me.lblUserName.BackColor = System.Drawing.SystemColors.Control
        Me.lblUserName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUserName.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUserName.Location = New System.Drawing.Point(12, 51)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUserName.Size = New System.Drawing.Size(126, 21)
        Me.lblUserName.TabIndex = 10
        Me.lblUserName.Text = "Mật Khẩu :"
        '
        'txtpassword
        '
        Me.txtpassword.AcceptsReturn = True
        Me.txtpassword.AutoSize = False
        Me.txtpassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtpassword.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtpassword.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtpassword.Location = New System.Drawing.Point(161, 47)
        Me.txtpassword.MaxLength = 0
        Me.txtpassword.Name = "txtpassword"
        Me.txtpassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtpassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtpassword.Size = New System.Drawing.Size(203, 26)
        Me.txtpassword.TabIndex = 12
        Me.txtpassword.Text = ""
        '
        'lblPassword
        '
        Me.lblPassword.BackColor = System.Drawing.SystemColors.Control
        Me.lblPassword.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPassword.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPassword.Location = New System.Drawing.Point(12, 79)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPassword.Size = New System.Drawing.Size(124, 20)
        Me.lblPassword.TabIndex = 11
        Me.lblPassword.Text = "Xác Nhận Lại :"
        '
        'txtconfirmPassword
        '
        Me.txtconfirmPassword.AcceptsReturn = True
        Me.txtconfirmPassword.AutoSize = False
        Me.txtconfirmPassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtconfirmPassword.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtconfirmPassword.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtconfirmPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtconfirmPassword.Location = New System.Drawing.Point(161, 78)
        Me.txtconfirmPassword.MaxLength = 0
        Me.txtconfirmPassword.Name = "txtconfirmPassword"
        Me.txtconfirmPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtconfirmPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtconfirmPassword.Size = New System.Drawing.Size(203, 26)
        Me.txtconfirmPassword.TabIndex = 13
        Me.txtconfirmPassword.Text = ""
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(232, 129)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(80, 25)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "&Cancel"
        '
        'cmdok
        '
        Me.cmdok.BackColor = System.Drawing.SystemColors.Control
        Me.cmdok.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdok.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdok.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdok.Location = New System.Drawing.Point(88, 128)
        Me.cmdok.Name = "cmdok"
        Me.cmdok.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdok.Size = New System.Drawing.Size(80, 25)
        Me.cmdok.TabIndex = 15
        Me.cmdok.Text = "&OK"
        '
        'PanelUsers
        '
        Me.PanelUsers.Controls.Add(Me.Lvusername)
        Me.PanelUsers.Location = New System.Drawing.Point(30, 5)
        Me.PanelUsers.Name = "PanelUsers"
        Me.PanelUsers.Size = New System.Drawing.Size(26, 156)
        Me.PanelUsers.TabIndex = 20
        '
        'Lvusername
        '
        Me.Lvusername.ContextMenu = Me.ContextMenuManager
        Me.Lvusername.FullRowSelect = True
        Me.Lvusername.GridLines = True
        Me.Lvusername.Location = New System.Drawing.Point(4, 4)
        Me.Lvusername.MultiSelect = False
        Me.Lvusername.Name = "Lvusername"
        Me.Lvusername.Size = New System.Drawing.Size(361, 148)
        Me.Lvusername.TabIndex = 0
        Me.Lvusername.View = System.Windows.Forms.View.Details
        '
        'ContextMenuManager
        '
        Me.ContextMenuManager.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemXoa, Me.MenuItemsetpass})
        '
        'MenuItemXoa
        '
        Me.MenuItemXoa.Index = 0
        Me.MenuItemXoa.Text = "&Xóa "
        '
        'MenuItemsetpass
        '
        Me.MenuItemsetpass.Index = 1
        Me.MenuItemsetpass.Text = "&Set Password"
        '
        'lblusers
        '
        Me.lblusers.Location = New System.Drawing.Point(1, 10)
        Me.lblusers.Name = "lblusers"
        Me.lblusers.Size = New System.Drawing.Size(24, 118)
        Me.lblusers.TabIndex = 19
        Me.lblusers.Text = ">>"
        Me.lblusers.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmNewUser
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 18)
        Me.ClientSize = New System.Drawing.Size(406, 164)
        Me.Controls.Add(Me.PanelUsers)
        Me.Controls.Add(Me.lblusers)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdok)
        Me.Controls.Add(Me.frmConnectionInfo)
        Me.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmNewUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Thêm Người Dùng"
        Me.frmConnectionInfo.ResumeLayout(False)
        Me.PanelUsers.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Function CheckInput() As Short
        Dim usertype As String
        If (Trim$(txtusername.Text) = "") Then
            Return 1
        End If

        If (Len(Trim$(txtusername.Text)) < 6) Then
            Return 2
        End If

        If (Trim$(txtpassword.Text) <> Trim$(txtconfirmPassword.Text)) Then
            Return 3
        End If
        CreateUserPass.UserName = Trim$(txtusername.Text)
        If (Not CreateUserPass.CheckUserPass) Then
            Return 4
        End If
        Return 6
    End Function

    Private Sub cmdok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdok.Click
        Dim result As Short
        result = CheckInput()
        Select Case result
            Case 1
                MsgBox("Chưa Nhập UserLogin!", MsgBoxStyle.Critical, "UserLogin Lổi")
                txtusername.Focus()
                Exit Sub
            Case 2
                MsgBox("UserLogin Ít Nhất Phải 6 Ký Tự!", MsgBoxStyle.Critical, "UserLogin Lổi")
                txtusername.Focus()
                Exit Sub
            Case 3
                MsgBox("Mật Khẩu Không Đúng!", MsgBoxStyle.Critical, "Sai Mật Khẩu")
                txtpassword.Focus()
                Exit Sub
            Case 4
                MsgBox("Username đã tồn tại!", MsgBoxStyle.Critical, "Trùng tên truy cập")
                txtusername.Focus()
                Exit Sub
            Case 5
                MsgBox("UserLogin Đặt Không Theo Qui Ước!", MsgBoxStyle.Critical, "UserLogin Lổi")
                txtusername.Focus()
                Exit Sub
            Case 6
                CreateUserPass.CreateAccount(Trim$(txtusername.Text), Trim$(txtpassword.Text))
                txtusername.Clear()
                txtpassword.Clear()
                txtconfirmPassword.Clear()
                FormatListview()
        End Select
    End Sub

    Private Sub frmNewUser_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PanelUsers.Width = 0
        CreateUserPass = New ClassCreateUser
        FormatListview()
    End Sub

    Private Sub lblusers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblusers.Click
        Dim i As Long
        Dim Pwidth As Long
        If (lblusers.Text = ">>") Then
            lblusers.Text = "<<"
            For i = 0 To 370000
                Pwidth = i / 1000
                PanelUsers.Width = Pwidth
            Next
        Else
            lblusers.Text = ">>"
            For i = 370000 To 0 Step -1
                Pwidth = i / 1000
                PanelUsers.Width = Pwidth
            Next
        End If
    End Sub

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
    Private Sub FillListview()
        LoadUsername()
        Dim i As Integer
        Dim count As Integer
        Dim strpass As String
        Dim strlen As String
        For i = 0 To mydataset.Tables("Pass").Rows.Count - 1
            Lvusername.Items.Add(CStr(i + 1))
            Lvusername.Items(i).SubItems.Add(mydataset.Tables("Pass").Rows(i).Item("UserName"))
            strlen = mydataset.Tables("Pass").Rows(i).Item("Passwrd")
            count = 0
            strpass = ""
            While (count < strlen.Length)
                strpass += "*"
                count += 1
            End While
            Lvusername.Items(i).SubItems.Add(strpass)

            Lvusername.Items(i).SubItems.Add(CStr(mydataset.Tables("Pass").Rows(i).Item("createdate")))
        Next
    End Sub

    Private Sub FormatListview()
        Lvusername.Items.Clear()
        Lvusername.Columns.Clear()
        Lvusername.Columns.Add("TT", 40, HorizontalAlignment.Center)
        Lvusername.Columns.Add("Tên Login", 100, HorizontalAlignment.Left)
        Lvusername.Columns.Add("Password", 100, HorizontalAlignment.Left)
        Lvusername.Columns.Add("Ngày Cấp", 120, HorizontalAlignment.Left)
        FillListview()
    End Sub

    Private Sub MenuItemXoa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemXoa.Click
        Dim lvitem As ListViewItem
        Dim ulogin As String
        lvitem = Lvusername.FocusedItem
        ulogin = lvitem.SubItems(1).Text
        Dim value = MsgBox("Bạn Có Thật Sự Muốn Xóa Account Này Khỏi Hệ Thống Không?" & vbCrLf & "Vui Lòng Kiểm Tra Lại Xem User Này Có Đang Login ..? Trước Khi Delete", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Xóa Account!")
        If (value = vbYes) Then
            CreateUserPass.UserName = ulogin
            CreateUserPass.DeleteAccount()
            FormatListview()
        End If
    End Sub

    Private Sub ContextMenuManager_Popup(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextMenuManager.Popup
        Try
            If (Lvusername.FocusedItem.Selected) Then
                MenuItemXoa.Enabled = True
                MenuItemsetpass.Enabled = True
            Else
                MenuItemXoa.Enabled = False
                MenuItemsetpass.Enabled = False
            End If
        Catch ex As Exception
            MenuItemXoa.Enabled = False
            MenuItemsetpass.Enabled = False
        End Try
    End Sub

    Private Sub MenuItemsetpass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemsetpass.Click
        Dim lvitem As ListViewItem
        Dim ulogin As String
        Dim pass As String
        lvitem = Lvusername.FocusedItem
        ulogin = lvitem.SubItems(1).Text
        Dim frm As New frmSetpassword(ulogin)
        frm.ShowDialog()
        FormatListview()

    End Sub

    Private Sub txtusername_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtusername.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            txtusername.Text = txtusername.Text.ToUpper()
            KeyAscii = 0
            txtpassword.Focus()
        End If
    End Sub

    Private Sub txtpassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtpassword.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            txtconfirmPassword.Focus()
        End If
    End Sub

    Private Sub txtconfirmPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtconfirmPassword.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdok_Click(sender, e)
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

   
End Class
