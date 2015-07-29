Imports ConvertNumberToChar
Imports System.Data.OleDb
Public Class frmEditReceipt
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private Indexlistview As Integer
    Dim start As Boolean = False
    Dim splitn As New SplitNumbers
    Dim numbers As New ConvertNumbersToString
    Dim rpt As CrystalReport_Receipts
    Dim SaveFlag As Boolean
    Dim ValueID As Long
    Dim DateValue As Date
    Dim soPT As Long
    Dim strMaloaiThu As String
    Dim Arrvar(15) As String
    Dim Arrval(15) As String
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
    Friend WithEvents cmbHTthu As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtslunc As System.Windows.Forms.TextBox
    Friend WithEvents CboEmploy_code As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtso As System.Windows.Forms.TextBox
    Friend WithEvents Cbolydo As System.Windows.Forms.ComboBox
    Friend WithEvents txtsoBK As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePickerchuky As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtquyen As System.Windows.Forms.TextBox
    Friend WithEvents txtsotien As System.Windows.Forms.TextBox
    Friend WithEvents txtchitietbk As System.Windows.Forms.TextBox
    Friend WithEvents txtsohd As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePickerngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtlydo As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtsoGNT As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cboAccounts As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerngaynop As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmddong As System.Windows.Forms.Button
    Friend WithEvents cmdLuu As System.Windows.Forms.Button
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents cbostations As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cmdPreview As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbHTthu = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmdPreview = New System.Windows.Forms.Button
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtslunc = New System.Windows.Forms.TextBox
        Me.CboEmploy_code = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtso = New System.Windows.Forms.TextBox
        Me.Cbolydo = New System.Windows.Forms.ComboBox
        Me.txtsoBK = New System.Windows.Forms.TextBox
        Me.DateTimePickerchuky = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtquyen = New System.Windows.Forms.TextBox
        Me.txtsotien = New System.Windows.Forms.TextBox
        Me.txtchitietbk = New System.Windows.Forms.TextBox
        Me.txtsohd = New System.Windows.Forms.TextBox
        Me.DateTimePickerngay = New System.Windows.Forms.DateTimePicker
        Me.txtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtlydo = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtsoGNT = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.cboAccounts = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.DateTimePickerngaynop = New System.Windows.Forms.DateTimePicker
        Me.Label18 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cmddong = New System.Windows.Forms.Button
        Me.cmdLuu = New System.Windows.Forms.Button
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.cbostations = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbHTthu
        '
        Me.cmbHTthu.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbHTthu.Location = New System.Drawing.Point(299, 16)
        Me.cmbHTthu.Name = "cmbHTthu"
        Me.cmbHTthu.Size = New System.Drawing.Size(272, 27)
        Me.cmbHTthu.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmbHTthu)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.cmdPreview)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.txtslunc)
        Me.GroupBox1.Controls.Add(Me.CboEmploy_code)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.txtso)
        Me.GroupBox1.Controls.Add(Me.Cbolydo)
        Me.GroupBox1.Controls.Add(Me.txtsoBK)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerchuky)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtquyen)
        Me.GroupBox1.Controls.Add(Me.txtsotien)
        Me.GroupBox1.Controls.Add(Me.txtchitietbk)
        Me.GroupBox1.Controls.Add(Me.txtsohd)
        Me.GroupBox1.Controls.Add(Me.DateTimePickerngay)
        Me.GroupBox1.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtlydo)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(0, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(577, 344)
        Me.GroupBox1.TabIndex = 53
        Me.GroupBox1.TabStop = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(192, 19)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(97, 22)
        Me.Label9.TabIndex = 57
        Me.Label9.Text = "Hình thức thu"
        '
        'cmdPreview
        '
        Me.cmdPreview.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPreview.Location = New System.Drawing.Point(440, 56)
        Me.cmdPreview.Name = "cmdPreview"
        Me.cmdPreview.Size = New System.Drawing.Size(80, 27)
        Me.cmdPreview.TabIndex = 5
        Me.cmdPreview.Text = "Xem"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(248, 219)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(64, 22)
        Me.Label19.TabIndex = 55
        Me.Label19.Text = "SL UNC"
        '
        'txtslunc
        '
        Me.txtslunc.BackColor = System.Drawing.Color.White
        Me.txtslunc.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtslunc.ForeColor = System.Drawing.Color.Black
        Me.txtslunc.Location = New System.Drawing.Point(328, 216)
        Me.txtslunc.Name = "txtslunc"
        Me.txtslunc.Size = New System.Drawing.Size(104, 26)
        Me.txtslunc.TabIndex = 13
        Me.txtslunc.Text = ""
        Me.txtslunc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CboEmploy_code
        '
        Me.CboEmploy_code.Location = New System.Drawing.Point(72, 88)
        Me.CboEmploy_code.Name = "CboEmploy_code"
        Me.CboEmploy_code.Size = New System.Drawing.Size(168, 27)
        Me.CboEmploy_code.TabIndex = 6
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label14.Location = New System.Drawing.Point(3, 48)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(573, 2)
        Me.Label14.TabIndex = 34
        '
        'txtso
        '
        Me.txtso.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtso.ForeColor = System.Drawing.Color.Black
        Me.txtso.Location = New System.Drawing.Point(71, 56)
        Me.txtso.Name = "txtso"
        Me.txtso.Size = New System.Drawing.Size(113, 26)
        Me.txtso.TabIndex = 3
        Me.txtso.Text = ""
        Me.txtso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Cbolydo
        '
        Me.Cbolydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbolydo.ItemHeight = 19
        Me.Cbolydo.Location = New System.Drawing.Point(71, 120)
        Me.Cbolydo.Name = "Cbolydo"
        Me.Cbolydo.Size = New System.Drawing.Size(169, 27)
        Me.Cbolydo.TabIndex = 7
        '
        'txtsoBK
        '
        Me.txtsoBK.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsoBK.ForeColor = System.Drawing.Color.Black
        Me.txtsoBK.Location = New System.Drawing.Point(72, 184)
        Me.txtsoBK.Name = "txtsoBK"
        Me.txtsoBK.Size = New System.Drawing.Size(137, 26)
        Me.txtsoBK.TabIndex = 10
        Me.txtsoBK.Text = ""
        Me.txtsoBK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DateTimePickerchuky
        '
        Me.DateTimePickerchuky.CustomFormat = "MM/yyyy"
        Me.DateTimePickerchuky.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerchuky.Location = New System.Drawing.Point(72, 152)
        Me.DateTimePickerchuky.Name = "DateTimePickerchuky"
        Me.DateTimePickerchuky.Size = New System.Drawing.Size(137, 26)
        Me.DateTimePickerchuky.TabIndex = 8
        Me.DateTimePickerchuky.Value = New Date(2005, 7, 1, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(192, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 22)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Quyển số"
        '
        'txtquyen
        '
        Me.txtquyen.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtquyen.ForeColor = System.Drawing.Color.Black
        Me.txtquyen.Location = New System.Drawing.Point(299, 56)
        Me.txtquyen.Name = "txtquyen"
        Me.txtquyen.Size = New System.Drawing.Size(113, 26)
        Me.txtquyen.TabIndex = 4
        Me.txtquyen.Text = ""
        Me.txtquyen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtsotien
        '
        Me.txtsotien.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsotien.ForeColor = System.Drawing.Color.Black
        Me.txtsotien.Location = New System.Drawing.Point(328, 152)
        Me.txtsotien.Name = "txtsotien"
        Me.txtsotien.Size = New System.Drawing.Size(144, 26)
        Me.txtsotien.TabIndex = 9
        Me.txtsotien.Text = ""
        Me.txtsotien.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtchitietbk
        '
        Me.txtchitietbk.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtchitietbk.ForeColor = System.Drawing.Color.Black
        Me.txtchitietbk.Location = New System.Drawing.Point(328, 184)
        Me.txtchitietbk.Name = "txtchitietbk"
        Me.txtchitietbk.Size = New System.Drawing.Size(240, 26)
        Me.txtchitietbk.TabIndex = 11
        Me.txtchitietbk.Text = ""
        '
        'txtsohd
        '
        Me.txtsohd.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsohd.ForeColor = System.Drawing.Color.Black
        Me.txtsohd.Location = New System.Drawing.Point(72, 216)
        Me.txtsohd.Name = "txtsohd"
        Me.txtsohd.Size = New System.Drawing.Size(136, 26)
        Me.txtsohd.TabIndex = 12
        Me.txtsohd.Text = ""
        Me.txtsohd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DateTimePickerngay
        '
        Me.DateTimePickerngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngay.Location = New System.Drawing.Point(71, 17)
        Me.DateTimePickerngay.Name = "DateTimePickerngay"
        Me.DateTimePickerngay.Size = New System.Drawing.Size(113, 26)
        Me.DateTimePickerngay.TabIndex = 1
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BackColor = System.Drawing.Color.White
        Me.txtEmployeeName.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeName.ForeColor = System.Drawing.Color.Blue
        Me.txtEmployeeName.Location = New System.Drawing.Point(240, 88)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(328, 26)
        Me.txtEmployeeName.TabIndex = 50
        Me.txtEmployeeName.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 22)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Số PT"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 22)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Ngày/T/N"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 22)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Mã CTV "
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 124)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 22)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Dịch vụ"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(248, 154)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(52, 22)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Số tiền"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 184)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(52, 22)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "SL BK"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 216)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 22)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "SL HĐ"
        '
        'txtlydo
        '
        Me.txtlydo.BackColor = System.Drawing.Color.White
        Me.txtlydo.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtlydo.ForeColor = System.Drawing.Color.Black
        Me.txtlydo.Location = New System.Drawing.Point(240, 120)
        Me.txtlydo.Name = "txtlydo"
        Me.txtlydo.Size = New System.Drawing.Size(328, 26)
        Me.txtlydo.TabIndex = 50
        Me.txtlydo.Text = ""
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(248, 186)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(82, 22)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "Chi tiết BK"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(8, 152)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(54, 22)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Chu kỳ"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtsoGNT)
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.Controls.Add(Me.cboAccounts)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.DateTimePickerngaynop)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Location = New System.Drawing.Point(5, 243)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(568, 48)
        Me.GroupBox3.TabIndex = 14
        Me.GroupBox3.TabStop = False
        '
        'txtsoGNT
        '
        Me.txtsoGNT.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsoGNT.ForeColor = System.Drawing.Color.Black
        Me.txtsoGNT.Location = New System.Drawing.Point(64, 14)
        Me.txtsoGNT.Name = "txtsoGNT"
        Me.txtsoGNT.Size = New System.Drawing.Size(64, 26)
        Me.txtsoGNT.TabIndex = 15
        Me.txtsoGNT.Text = ""
        Me.txtsoGNT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(2, 18)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(61, 22)
        Me.Label17.TabIndex = 55
        Me.Label17.Text = "Số GNT"
        '
        'cboAccounts
        '
        Me.cboAccounts.Location = New System.Drawing.Point(400, 14)
        Me.cboAccounts.Name = "cboAccounts"
        Me.cboAccounts.Size = New System.Drawing.Size(160, 27)
        Me.cboAccounts.TabIndex = 17
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(318, 18)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(88, 22)
        Me.Label16.TabIndex = 52
        Me.Label16.Text = "Số tài khỏan"
        '
        'DateTimePickerngaynop
        '
        Me.DateTimePickerngaynop.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerngaynop.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerngaynop.Location = New System.Drawing.Point(208, 14)
        Me.DateTimePickerngaynop.Name = "DateTimePickerngaynop"
        Me.DateTimePickerngaynop.Size = New System.Drawing.Size(104, 26)
        Me.DateTimePickerngaynop.TabIndex = 16
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(136, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(70, 22)
        Me.Label18.TabIndex = 25
        Me.Label18.Text = "Ngày nộp"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdDelete)
        Me.GroupBox2.Controls.Add(Me.cmddong)
        Me.GroupBox2.Controls.Add(Me.cmdLuu)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 291)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(566, 48)
        Me.GroupBox2.TabIndex = 18
        Me.GroupBox2.TabStop = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdDelete.Location = New System.Drawing.Point(24, 15)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(80, 27)
        Me.cmdDelete.TabIndex = 20
        Me.cmdDelete.Text = "Xoá"
        '
        'cmddong
        '
        Me.cmddong.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmddong.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmddong.Location = New System.Drawing.Point(480, 15)
        Me.cmddong.Name = "cmddong"
        Me.cmddong.Size = New System.Drawing.Size(80, 27)
        Me.cmddong.TabIndex = 31
        Me.cmddong.Text = "Đóng"
        '
        'cmdLuu
        '
        Me.cmdLuu.Font = New System.Drawing.Font("Times New Roman", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLuu.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdLuu.Location = New System.Drawing.Point(128, 15)
        Me.cmdLuu.Name = "cmdLuu"
        Me.cmdLuu.Size = New System.Drawing.Size(80, 27)
        Me.cmdLuu.TabIndex = 19
        Me.cmdLuu.Text = "Cập nhật"
        '
        'cbostations
        '
        Me.cbostations.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbostations.Location = New System.Drawing.Point(304, 15)
        Me.cbostations.Name = "cbostations"
        Me.cbostations.Size = New System.Drawing.Size(280, 27)
        Me.cbostations.TabIndex = 54
        '
        'Label15
        '
        Me.Label15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label15.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label15.Font = New System.Drawing.Font("Arial", 21.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label15.Location = New System.Drawing.Point(5, 3)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(579, 40)
        Me.Label15.TabIndex = 56
        Me.Label15.Text = "  Điều chỉnh phiếu thu"
        '
        'frmEditReceipt
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 392)
        Me.Controls.Add(Me.cbostations)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmEditReceipt"
        Me.Text = "Điều chỉnh phiếu thu"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub FillDataSet()

        strSQL = "SELECT MaLoaithu,TenLoaiThu FROM Tbl_LoaiThu"
        FillCombo(cmbHTthu, strSQL, "Tbl_LoaiThu", "TenLoaiThu", "MaLoaithu")

        strSQL = "SELECT StationID,Station_Name,Station_Address FROM Tbl_Stations "
        FillCombo(cbostations, strSQL, "Tbl_Stations", "Station_Name", "StationID")

        strSQL = "SELECT Service_Code,Service_Name FROM Tbl_Services "
        FillCombo(Cbolydo, strSQL, "Tbl_Services", "Service_Code", "Service_Name")

        strSQL = "SELECT Account_No,Bank_Name FROM Tbl_Accounts_Banks,Tbl_Banks WHERE Tbl_Banks.Bank_Code = Tbl_Accounts_Banks.Bank_Code"
        FillCombo(cboAccounts, strSQL, "Tbl_Accounts_Banks", "Account_No", "Bank_Name")

    End Sub

    Private Sub FillCombo(ByRef cbo As ComboBox, ByVal strQuery As String, ByVal strTablename As String, ByVal strdislaymember As String, ByVal strValuemember As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(mydataset, strTablename)
            cbo.DataSource = mydataset.Tables(strTablename).DefaultView
            cbo.DisplayMember = strdislaymember
            cbo.ValueMember = strValuemember
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub LoadInfo(ByVal strOrdinal_No As Long)
        Dim strQuery As String
        Dim value
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader

        strQuery = "SELECT Receipt_Date, Ordinal_No, Volume, Service_Code, Descriptions, List_Quantity, List_Detail, Invoice_Quantity, Charge_Cycle, Total_Money, Employ_Code, MaLoaiThu, Account_Code,  Pay_Date, Pay_No,Status,ID FROM Tbl_Receipts " & _
        " WHERE Ordinal_No = " & strOrdinal_No & " AND Receipt_Date = #" & DateTimePickerngay.Value.ToShortDateString & "# AND Maloaithu ='" & cmbHTthu.SelectedValue & "'"
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = strQuery
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then

                If Not oleread.IsDBNull(0) Then
                    DateTimePickerngay.Value = oleread.GetDateTime(0)
                    DateValue = oleread.GetDateTime(0)
                End If

                If Not oleread.IsDBNull(1) Then
                    txtso.Text = oleread.GetValue(1)
                    soPT = oleread.GetValue(1)
                End If

                If Not oleread.IsDBNull(2) Then
                    txtquyen.Text = oleread.GetValue(2)
                End If

                If Not oleread.IsDBNull(3) Then
                    Cbolydo.SelectedIndex = Cbolydo.FindString(oleread.GetString(3))
                End If

                If Not oleread.IsDBNull(4) Then
                    txtlydo.Text = oleread.GetString(4)
                End If


                If Not oleread.IsDBNull(5) Then
                    txtsoBK.Text = oleread.GetValue(5)
                End If

                If Not oleread.IsDBNull(6) Then
                    txtchitietbk.Text = oleread.GetString(6)
                End If

                If Not oleread.IsDBNull(7) Then
                    txtsohd.Text = oleread.GetValue(7)
                End If

                If Not oleread.IsDBNull(8) Then
                    DateTimePickerchuky.Value = CDate(oleread.GetDateTime(8))
                End If

                If Not oleread.IsDBNull(9) Then
                    txtsotien.Text = oleread.GetValue(9)
                End If


                If Not oleread.IsDBNull(10) Then
                    CboEmploy_code.SelectedIndex = CboEmploy_code.FindString(oleread.GetString(10))
                End If

                'Ma Loai thu
                If Not oleread.IsDBNull(11) Then
                    strMaloaiThu = oleread.GetString(11)
                    cmbHTthu.SelectedIndex = cmbHTthu.FindString(GetStringName("Tbl_LoaiThu", oleread.GetString(11)))
                End If

                If (cmbHTthu.SelectedValue = "GNT") Then
                    GroupBox3.Enabled = True
                    txtslunc.Visible = True
                    Label19.Visible = True
                Else
                    GroupBox3.Enabled = False
                    txtslunc.Visible = False
                    Label19.Visible = False
                End If

                If Not oleread.IsDBNull(12) Then
                    cboAccounts.SelectedIndex = cboAccounts.FindString(oleread.GetString(12))
                End If

                If Not oleread.IsDBNull(13) Then
                    DateTimePickerngaynop.Value = oleread.GetDateTime(13)
                End If

                If Not oleread.IsDBNull(14) Then
                    txtsoGNT.Text = oleread.GetValue(14)
                End If

                If Not oleread.IsDBNull(15) Then
                    SaveFlag = oleread.GetBoolean(15)
                End If

                If Not oleread.IsDBNull(16) Then
                    ValueID = oleread.GetValue(16)
                End If

            Else
                MsgBox("Không tìm thấy số phiếu thu này")
            End If
            olecommand.Dispose()
            oledbcon.Close()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :" & ex.ToString)
        End Try

    End Sub

    Private Sub UpdateReciept()

        strSQL = " UPDATE Tbl_Receipts SET Ordinal_No = " & CLng(Trim$(txtso.Text)) & _
         ",Volume =" & CInt(Trim$(txtquyen.Text)) & _
         ",Receipt_Date = '" & DateTimePickerngay.Value.ToShortDateString & _
         "',Service_Code ='" & Cbolydo.Text & _
         "',Descriptions ='" & txtlydo.Text & _
         "',List_Quantity =" & CInt(txtsoBK.Text) & _
         ",List_Detail ='" & txtchitietbk.Text & _
         "',Invoice_Quantity =" & CInt(txtsohd.Text) & _
         ",Charge_Cycle = '" & DateTimePickerchuky.Text & _
         "',Total_Money = " & CLng(txtsotien.Text) & _
         ",Employ_Code = '" & CboEmploy_code.Text & _
         "',MaLoaiThu = '" & cmbHTthu.SelectedValue & "'"

        Dim strMaLoai As String
        strMaLoai = cmbHTthu.SelectedValue
        Select Case strMaLoai
            Case "GNT"
                strSQL += " ,Pay_No = " & CInt(txtsoGNT.Text)
                strSQL += " ,Pay_Date = '" & DateTimePickerngaynop.Value.ToShortDateString & "'"
                strSQL += ", Account_Code = '" & cboAccounts.Text & "'"
                strSQL += ", NguoiNop = '" & CboEmploy_code.Text & "'"
            Case "UNC"
                strSQL += ", SLUNC = " & CInt(txtslunc.Text)
        End Select

        strSQL += " WHERE ID = " & ValueID
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strSQL
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
            'MsgBox("Đã cập nhật vào hệ thống!!")
        Catch ex As Exception
            MsgBox(" Lổi rồi người ơi!!!" & ex.ToString)
        End Try
        oledbcon.Close()

    End Sub

    Private Sub UpdateSoQuy()

        If (cmbHTthu.SelectedValue = "TM") Then

            strSQL = " UPDATE Tbl_Receipts_Expenses SET Recei_No = " & CLng(Trim$(txtso.Text)) & _
            ",Recei_Vol =" & CInt(Trim$(txtquyen.Text)) & _
            ",Recei_Expen_Date = '" & DateTimePickerngay.Value.ToShortDateString & _
            "',Descriptions ='" & txtlydo.Text & " Kỳ cước " & DateTimePickerchuky.Text & _
            "',Recei_Money = " & CLng(txtsotien.Text) & _
            " WHERE Recei_Expen_Date = #" & DateValue.ToShortDateString & "# AND Recei_No = " & soPT
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
                MsgBox(ex.ToString)
            End Try
            oledbcon.Close()
        End If
    End Sub

    Private Sub DeleteReciept()

        strSQL = " DELETE FROM Tbl_Receipts WHERE ID = " & ValueID
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strSQL
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
            MsgBox("Đã xóa phiếu thu :" & soPT & "  hệ thống!!")
        Catch ex As Exception
            MsgBox(" Lổi rồi người ơi!!!" & ex.ToString)
        End Try
        oledbcon.Close()

    End Sub

    Private Sub DeleteSoQuy()

        If (strMaloaiThu = "TM") Then

            strSQL = " DELETE FROM Tbl_Receipts_Expenses   WHERE Recei_Expen_Date = #" & DateValue.ToShortDateString & "# AND Recei_No = " & soPT
            Try
                oledbcon.Open()
                Dim olecommand As New OleDbCommand
                olecommand.CommandText = strSQL
                olecommand.CommandType = CommandType.Text
                olecommand.Connection = oledbcon
                olecommand.ExecuteNonQuery()
                olecommand.Dispose()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            oledbcon.Close()
        End If
    End Sub

    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click
        LoadInfo(CLng(txtso.Text))
    End Sub

    Private Sub cmddong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmddong.Click
        Me.Close()
    End Sub

    Private Sub frmEditReceipt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePickerngay.Value = Now
        DateTimePickerchuky.Value = Now
        Try
            mydataset = New DataSet
            FillDataSet()
        Catch eLoad As System.Exception
            System.Windows.Forms.MessageBox.Show(eLoad.Message)
        End Try

        start = True

        If (cbostations.Items.Count > 0) Then
            cbostations.SelectedIndex = 0
            Try
                strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
                Try
                    CboEmploy_code.DataSource = Nothing
                    Try
                        mydataset.Tables("Tbl_Employee").Clear()
                    Catch ex As Exception
                    End Try
                    CboEmploy_code.Items.Clear()
                    FillCombo(CboEmploy_code, strSQL, "Tbl_Employee", "Employ_Code", "Employ_Name")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
            End Try
        End If

        If (CboEmploy_code.Items.Count > 0) Then
            CboEmploy_code.SelectedIndex = 0
            txtEmployeeName.Text = CboEmploy_code.SelectedValue
        End If

        If (Cbolydo.Items.Count > 0) Then
            Cbolydo.SelectedIndex = 0
            txtlydo.Text = "Thu " & Cbolydo.SelectedValue
        End If

        If (cmbHTthu.Items.Count > 0) Then
            If (cmbHTthu.FindString("TIỀN MẶT") > 0) Then
                cmbHTthu.SelectedIndex = cmbHTthu.FindString("TIỀN MẶT")
            End If
        End If
        txtEmployeeName.ReadOnly = True

    End Sub

    Private Sub cbostations_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbostations.SelectedIndexChanged
        Try
            strSQL = "SELECT Employ_Code,Employ_Name FROM Tbl_Employee WHERE StationID = '" & cbostations.SelectedValue & "'"
            Try
                CboEmploy_code.DataSource = Nothing
                Try
                    mydataset.Tables("Tbl_Employee").Clear()
                Catch ex As Exception
                End Try
                CboEmploy_code.Items.Clear()
                FillCombo(CboEmploy_code, strSQL, "Tbl_Employee", "Employ_Code", "Employ_Name")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Cbolydo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbolydo.SelectedIndexChanged
        If (start) Then
            Try
                Dim str As String
                str = Cbolydo.Text
                txtlydo.Text = "Thu " & Cbolydo.SelectedValue
                If (str = "KHAC") Then
                    txtlydo.ReadOnly = False
                Else
                    txtlydo.ReadOnly = True
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub CboEmploy_code_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboEmploy_code.SelectedIndexChanged
        If (start) Then
            Try
                txtEmployeeName.Text = CboEmploy_code.SelectedValue
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub CboEmploy_code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CboEmploy_code.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (CboEmploy_code.FindString(CboEmploy_code.Text) > -1) Then
                CboEmploy_code.SelectedIndex = CboEmploy_code.FindString(CboEmploy_code.Text)
                txtEmployeeName.Text = CboEmploy_code.SelectedValue
            Else
                MsgBox("Không tìm thấy Mã CTV tương ứng", MsgBoxStyle.Critical, "Nhập sai")
                txtEmployeeName.Text = ""
                CboEmploy_code.Focus()
                Exit Sub
            End If
            Cbolydo.Focus()
        End If
    End Sub

    Private Sub Cbolydo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Cbolydo.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            If (Cbolydo.FindString(Cbolydo.Text) > -1) Then
                Cbolydo.SelectedIndex = Cbolydo.FindString(Cbolydo.Text)
            Else
                MsgBox("Không tìm thấy dịch vụ tương ứng", MsgBoxStyle.Critical, "Nhập sai")
                txtlydo.Text = ""
                Cbolydo.Focus()
                Exit Sub
            End If
            DateTimePickerchuky.Focus()
        End If
    End Sub

    Private Sub cmdLuu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLuu.Click
        Dim value
        value = MsgBox("Bạn có thật sự muốn cập nhật Phiếu thu này không!", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Xác nhận")
        If (value = vbYes) Then

            Dim strloaithu As String
            strloaithu = cmbHTthu.SelectedValue
            If (strMaloaiThu <> strloaithu) Then
                MsgBox("Không thể thay đổi qua lại giữa các hình thức thu." & vbCrLf & "Trong trường hợp nhập sai hình thức thu, bạn nên xóa phiếu thu này và nhập lại đúng loại.", MsgBoxStyle.Critical, "Lỗi cập nhật")
                Exit Sub
            End If

            If (Trim$(txtso.Text) = "" OrElse Not IsNumeric(Trim$(txtso.Text))) Then
                MsgBox("Chưa nhập phiếu thu ")
                txtso.Focus()
                Exit Sub
            End If

            If (Trim$(txtquyen.Text) = "" OrElse Not IsNumeric(Trim$(txtquyen.Text))) Then
                MsgBox("Chưa nhập số quyển ")
                txtquyen.Focus()
                Exit Sub
            End If

            If (CboEmploy_code.Text = "") Then
                MsgBox("Chưa chọn mã CTV ")
                CboEmploy_code.Focus()
                Exit Sub
            End If

            Select Case strloaithu
                Case "TM"
                    If (SaveFlag) Then
                        MsgBox("Phiếu thu này đã đuợc chi không thể cập nhật được!", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    UpdateReciept()
                    UpdateSoQuy()
                Case "GNT"
                    If (Trim$(txtsoGNT.Text) = "") Then
                        MsgBox("Số giấy nộp tiền chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                        txtsoGNT.Focus()
                        Exit Sub
                    End If

                    If (Not IsNumeric(Trim$(txtsoGNT.Text))) Then
                        MsgBox("Số giấy nộp tiền phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                        txtsoGNT.Focus()
                        txtsoGNT.SelectAll()
                        Exit Sub
                    End If

                    If (CLng(Trim$(txtsoGNT.Text)) < 1) Then
                        MsgBox("Số giấy nộp tiền phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                        txtsoGNT.Focus()
                        txtsoGNT.SelectAll()
                        Exit Sub
                    End If

                    If (Trim$(cboAccounts.Text) = "") Then
                        MsgBox("Số tài khoản chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                        cboAccounts.Focus()
                        Exit Sub
                    End If
                    UpdateReciept()
                Case "UNC"

                    If (Trim$(txtslunc.Text) = "") Then
                        MsgBox("Số lượng UNC chưa được nhập vào!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                        txtslunc.Focus()
                        Exit Sub
                    End If

                    If (Not IsNumeric(Trim$(txtslunc.Text))) Then
                        MsgBox("Số lượng UNC phải là số nguyên!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                        txtslunc.Focus()
                        txtslunc.SelectAll()
                        Exit Sub
                    End If

                    If (CLng(Trim$(txtslunc.Text)) < 1) Then
                        MsgBox("Số lượng UNC phải lớn hơn 0!", MsgBoxStyle.Critical, "Lổi nhập liệu.")
                        txtslunc.Focus()
                        txtslunc.SelectAll()
                        Exit Sub
                    End If
                    UpdateReciept()
            End Select

        End If

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim value
        value = MsgBox("Bạn có thật sự muốn xoá Phiếu thu này không!", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Xác nhận")
        If (value = vbYes) Then
            If (SaveFlag) Then
                MsgBox("Phiếu thu này đã đuợc chi không thể xóa được!", MsgBoxStyle.Critical)
                Exit Sub
            End If
            DeleteReciept()
            DeleteSoQuy()
        End If
    End Sub

    Private Function GetStringName(ByVal strTablename As String, ByVal strCode As String) As String
        Dim i As Integer
        Dim strresult As String
        For i = 0 To mydataset.Tables(strTablename).Rows.Count - 1
            strresult = mydataset.Tables(strTablename).Rows(i).Item(0)
            If (strresult.Equals(strCode)) Then
                strresult = mydataset.Tables(strTablename).Rows(i).Item(1)
                Exit For
            End If
        Next
        Return strresult
    End Function

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click

    End Sub
    Private Sub cmbHTthu_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHTthu.SelectedIndexChanged

    End Sub
End Class
