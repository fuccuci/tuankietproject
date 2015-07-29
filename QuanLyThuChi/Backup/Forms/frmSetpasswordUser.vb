Public Class frmSetpasswordUser
    Inherits System.Windows.Forms.Form
    Dim CreateUserPass As ClassCreateUser

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
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdok As System.Windows.Forms.Button
    Public WithEvents frmConnectionInfo As System.Windows.Forms.GroupBox
    Public WithEvents lblUserName As System.Windows.Forms.Label
    Public WithEvents txtpasswordnew As System.Windows.Forms.TextBox
    Public WithEvents lblPassword As System.Windows.Forms.Label
    Public WithEvents txtconfirmPassword As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents txtmatkhaucu As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdok = New System.Windows.Forms.Button
        Me.frmConnectionInfo = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtmatkhaucu = New System.Windows.Forms.TextBox
        Me.lblUserName = New System.Windows.Forms.Label
        Me.txtpasswordnew = New System.Windows.Forms.TextBox
        Me.lblPassword = New System.Windows.Forms.Label
        Me.txtconfirmPassword = New System.Windows.Forms.TextBox
        Me.frmConnectionInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(189, 113)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(80, 25)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "&Cancel"
        '
        'cmdok
        '
        Me.cmdok.BackColor = System.Drawing.SystemColors.Control
        Me.cmdok.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdok.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdok.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdok.Location = New System.Drawing.Point(77, 113)
        Me.cmdok.Name = "cmdok"
        Me.cmdok.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdok.Size = New System.Drawing.Size(80, 25)
        Me.cmdok.TabIndex = 4
        Me.cmdok.Text = "&OK"
        '
        'frmConnectionInfo
        '
        Me.frmConnectionInfo.BackColor = System.Drawing.SystemColors.Control
        Me.frmConnectionInfo.Controls.Add(Me.Label1)
        Me.frmConnectionInfo.Controls.Add(Me.txtmatkhaucu)
        Me.frmConnectionInfo.Controls.Add(Me.lblUserName)
        Me.frmConnectionInfo.Controls.Add(Me.txtpasswordnew)
        Me.frmConnectionInfo.Controls.Add(Me.lblPassword)
        Me.frmConnectionInfo.Controls.Add(Me.txtconfirmPassword)
        Me.frmConnectionInfo.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmConnectionInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmConnectionInfo.Location = New System.Drawing.Point(4, -4)
        Me.frmConnectionInfo.Name = "frmConnectionInfo"
        Me.frmConnectionInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmConnectionInfo.Size = New System.Drawing.Size(320, 108)
        Me.frmConnectionInfo.TabIndex = 0
        Me.frmConnectionInfo.TabStop = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(9, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(126, 21)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Mật Khẩu Cũ"
        '
        'txtmatkhaucu
        '
        Me.txtmatkhaucu.AcceptsReturn = True
        Me.txtmatkhaucu.AutoSize = False
        Me.txtmatkhaucu.BackColor = System.Drawing.SystemColors.Window
        Me.txtmatkhaucu.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtmatkhaucu.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmatkhaucu.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtmatkhaucu.Location = New System.Drawing.Point(144, 16)
        Me.txtmatkhaucu.MaxLength = 0
        Me.txtmatkhaucu.Name = "txtmatkhaucu"
        Me.txtmatkhaucu.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtmatkhaucu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtmatkhaucu.Size = New System.Drawing.Size(168, 26)
        Me.txtmatkhaucu.TabIndex = 1
        Me.txtmatkhaucu.Text = ""
        '
        'lblUserName
        '
        Me.lblUserName.BackColor = System.Drawing.SystemColors.Control
        Me.lblUserName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUserName.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUserName.Location = New System.Drawing.Point(10, 48)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUserName.Size = New System.Drawing.Size(126, 21)
        Me.lblUserName.TabIndex = 10
        Me.lblUserName.Text = "Mật Khẩu Mới "
        '
        'txtpasswordnew
        '
        Me.txtpasswordnew.AcceptsReturn = True
        Me.txtpasswordnew.AutoSize = False
        Me.txtpasswordnew.BackColor = System.Drawing.SystemColors.Window
        Me.txtpasswordnew.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtpasswordnew.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpasswordnew.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtpasswordnew.Location = New System.Drawing.Point(144, 45)
        Me.txtpasswordnew.MaxLength = 0
        Me.txtpasswordnew.Name = "txtpasswordnew"
        Me.txtpasswordnew.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtpasswordnew.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtpasswordnew.Size = New System.Drawing.Size(168, 26)
        Me.txtpasswordnew.TabIndex = 2
        Me.txtpasswordnew.Text = ""
        '
        'lblPassword
        '
        Me.lblPassword.BackColor = System.Drawing.SystemColors.Control
        Me.lblPassword.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPassword.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPassword.Location = New System.Drawing.Point(10, 75)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPassword.Size = New System.Drawing.Size(118, 20)
        Me.lblPassword.TabIndex = 11
        Me.lblPassword.Text = "Xác Nhận Lại "
        '
        'txtconfirmPassword
        '
        Me.txtconfirmPassword.AcceptsReturn = True
        Me.txtconfirmPassword.AutoSize = False
        Me.txtconfirmPassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtconfirmPassword.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtconfirmPassword.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtconfirmPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtconfirmPassword.Location = New System.Drawing.Point(144, 74)
        Me.txtconfirmPassword.MaxLength = 0
        Me.txtconfirmPassword.Name = "txtconfirmPassword"
        Me.txtconfirmPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtconfirmPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtconfirmPassword.Size = New System.Drawing.Size(168, 26)
        Me.txtconfirmPassword.TabIndex = 3
        Me.txtconfirmPassword.Text = ""
        '
        'frmSetpasswordUser
        '
        Me.AcceptButton = Me.cmdok
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(326, 147)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdok)
        Me.Controls.Add(Me.frmConnectionInfo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmSetpasswordUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Thay đổi password"
        Me.frmConnectionInfo.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdok.Click
        If (Trim$(txtmatkhaucu.Text) <> Password) Then
            MsgBox("Password Cũ Không Đúng!", MsgBoxStyle.Critical, "Set Password")
            txtmatkhaucu.Focus()
            Exit Sub
        End If

        If (Trim$(txtpasswordnew.Text) <> Trim$(txtconfirmPassword.Text)) Then
            MsgBox("Password Nhập Không Có Giá Trị!", MsgBoxStyle.Critical, "Set Password")
            txtpasswordnew.Focus()
            Exit Sub
        End If
        CreateUserPass.Password = Trim$(txtpasswordnew.Text)
        CreateUserPass.UserName = UserName
        CreateUserPass.SetPassword()
        Me.Close()
    End Sub

    Private Sub frmSetpassword_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CreateUserPass = New ClassCreateUser
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub
End Class
