Public Class frmBankAccounts
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents cbokhuvuc As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txttendaily As System.Windows.Forms.TextBox
    Friend WithEvents txtmadaily As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
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
        Me.txttendaily = New System.Windows.Forms.TextBox
        Me.txtmadaily = New System.Windows.Forms.TextBox
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
        Me.dgridDaily.Location = New System.Drawing.Point(6, 152)
        Me.dgridDaily.Name = "dgridDaily"
        Me.dgridDaily.ReadOnly = True
        Me.dgridDaily.Size = New System.Drawing.Size(528, 184)
        Me.dgridDaily.TabIndex = 76
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
        Me.Label9.TabIndex = 74
        Me.Label9.Text = "TÀI KHOẢN - NGÂN HÀNG"
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
        Me.GroupBox2.Location = New System.Drawing.Point(8, 332)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(528, 48)
        Me.GroupBox2.TabIndex = 75
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
        Me.GroupBox1.Controls.Add(Me.txttendaily)
        Me.GroupBox1.Controls.Add(Me.txtmadaily)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(528, 106)
        Me.GroupBox1.TabIndex = 73
        Me.GroupBox1.TabStop = False
        '
        'cbokhuvuc
        '
        Me.cbokhuvuc.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbokhuvuc.Location = New System.Drawing.Point(109, 72)
        Me.cbokhuvuc.Name = "cbokhuvuc"
        Me.cbokhuvuc.Size = New System.Drawing.Size(411, 27)
        Me.cbokhuvuc.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(7, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 24)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Tại Ngân Hàng"
        '
        'txttendaily
        '
        Me.txttendaily.BackColor = System.Drawing.Color.White
        Me.txttendaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttendaily.Location = New System.Drawing.Point(109, 42)
        Me.txttendaily.Name = "txttendaily"
        Me.txttendaily.Size = New System.Drawing.Size(411, 26)
        Me.txttendaily.TabIndex = 3
        Me.txttendaily.Text = ""
        '
        'txtmadaily
        '
        Me.txtmadaily.BackColor = System.Drawing.Color.White
        Me.txtmadaily.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmadaily.ForeColor = System.Drawing.Color.Blue
        Me.txtmadaily.Location = New System.Drawing.Point(109, 12)
        Me.txtmadaily.Name = "txtmadaily"
        Me.txtmadaily.Size = New System.Drawing.Size(411, 26)
        Me.txtmadaily.TabIndex = 1
        Me.txtmadaily.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(7, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(110, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Tên Tài Khoản"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(7, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(118, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Số Tài Khoản"
        '
        'frmBankAccounts
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(536, 381)
        Me.Controls.Add(Me.dgridDaily)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmBankAccounts"
        Me.Text = "frmBankAccounts"
        CType(Me.dgridDaily, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
