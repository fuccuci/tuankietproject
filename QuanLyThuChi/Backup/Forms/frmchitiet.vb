Imports System.data.OleDb
Public Class frmchitiet
    Inherits System.Windows.Forms.Form
    Private table As New DataTable
    Private strISDN As String
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal _strISDN As String)
        MyBase.New()
        InitializeComponent()
        FormatGridDSKH()
        strISDN = _strISDN
        LoadInfo()
        LoadDetailISDN()
    End Sub

    Public Sub New()
        MyBase.New()
        InitializeComponent()
        FormatGridDSKH()
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
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents DataGridDskh As System.Windows.Forms.DataGrid
    Friend WithEvents txtDiachi As System.Windows.Forms.TextBox
    Friend WithEvents txtHotenkh As System.Windows.Forms.TextBox
    Friend WithEvents txtSomay As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdclose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtHotenkh = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDiachi = New System.Windows.Forms.TextBox
        Me.txtSomay = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.DataGridDskh = New System.Windows.Forms.DataGrid
        Me.cmdclose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridDskh, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdclose)
        Me.GroupBox1.Controls.Add(Me.txtHotenkh)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtDiachi)
        Me.GroupBox1.Controls.Add(Me.txtSomay)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.DataGridDskh)
        Me.GroupBox1.Location = New System.Drawing.Point(3, -2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(621, 378)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txtHotenkh
        '
        Me.txtHotenkh.AutoSize = False
        Me.txtHotenkh.BackColor = System.Drawing.Color.White
        Me.txtHotenkh.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHotenkh.Location = New System.Drawing.Point(132, 121)
        Me.txtHotenkh.Name = "txtHotenkh"
        Me.txtHotenkh.Size = New System.Drawing.Size(360, 24)
        Me.txtHotenkh.TabIndex = 15
        Me.txtHotenkh.Text = ""
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(15, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 24)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Địa chỉ:"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(15, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(125, 24)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Tên Khách hàng:"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(15, 94)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 24)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Số máy:"
        '
        'txtDiachi
        '
        Me.txtDiachi.AutoSize = False
        Me.txtDiachi.BackColor = System.Drawing.Color.White
        Me.txtDiachi.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDiachi.Location = New System.Drawing.Point(132, 150)
        Me.txtDiachi.Name = "txtDiachi"
        Me.txtDiachi.Size = New System.Drawing.Size(360, 24)
        Me.txtDiachi.TabIndex = 16
        Me.txtDiachi.Text = ""
        '
        'txtSomay
        '
        Me.txtSomay.AutoSize = False
        Me.txtSomay.BackColor = System.Drawing.Color.White
        Me.txtSomay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSomay.ForeColor = System.Drawing.Color.IndianRed
        Me.txtSomay.Location = New System.Drawing.Point(132, 92)
        Me.txtSomay.Name = "txtSomay"
        Me.txtSomay.Size = New System.Drawing.Size(120, 24)
        Me.txtSomay.TabIndex = 14
        Me.txtSomay.Text = ""
        '
        'Label10
        '
        Me.Label10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label10.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label10.Font = New System.Drawing.Font("VNI-Allegie", 30.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.Label10.Location = New System.Drawing.Point(4, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(612, 73)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "Chi Tieát Khieáu Naïi Khaùch Haøng"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DataGridDskh
        '
        Me.DataGridDskh.DataMember = ""
        Me.DataGridDskh.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridDskh.Location = New System.Drawing.Point(8, 181)
        Me.DataGridDskh.Name = "DataGridDskh"
        Me.DataGridDskh.Size = New System.Drawing.Size(608, 192)
        Me.DataGridDskh.TabIndex = 0
        '
        'cmdclose
        '
        Me.cmdclose.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdclose.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdclose.Location = New System.Drawing.Point(512, 152)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(88, 24)
        Me.cmdclose.TabIndex = 21
        Me.cmdclose.Text = "&Đóng"
        '
        'frmchitiet
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(624, 382)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmchitiet"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CHI TIẾT KHIẾU NẠI THEO SỐ MÁY"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridDskh, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub FormatGridDSKH()

        With DataGridDskh
            .AllowNavigation = False
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Chi Tiết Khiếu Nại Khách Hàng...."
            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "Chitiet"
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

                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "SoCV"
                    .HeaderText = "Số CV"
                    .Width = 60
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "ISDN"
                    .HeaderText = "Số Điện Thoại"
                    .Width = 130
                    .Alignment = HorizontalAlignment.Center
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "Reci_Date1"
                    .HeaderText = "Ngày nhận"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "Reci_Date2"
                    .HeaderText = "Ngày báo nhận"
                    .Width = 120
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "Content_ComPlain"
                    .HeaderText = "Nội dung khiếu nại"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "Content_Explain"
                    .HeaderText = "Nội dung trả lời"
                    .Width = 200
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(6)
                    .MappingName = "Explain_Date"
                    .HeaderText = "Ngày trả lời"
                    .Width = 120
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(7)
                    .MappingName = "ComName"
                    .HeaderText = "Loại khiếu nại"
                    .Width = 120
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(8)
                    .MappingName = "StaticName"
                    .HeaderText = "Tình trạng"
                    .Width = 120
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                End With

            End With
        End With
        DataGridDskh.TableStyles.Add(TblStyle)
    End Sub
    Private Sub LoadDetailISDN()
        Try
            DataGridDskh.DataSource = Nothing
            Dim olecommand As New OleDbCommand
            Dim oleAdtapter As New OleDbDataAdapter
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = "SELECT SoCV,ISDN,Reci_Date1,Reci_Date2,Content_ComPlain,Content_Explain,Explain_Date ,Tbl_Type_Complain.TypeName AS ComName , Tbl_TypeStatics.TypeName AS StaticName  FROM tbl_Solve,Tbl_TypeStatics,Tbl_Type_Complain WHERE tbl_Solve.Typecom = Tbl_Type_Complain.TypeID AND tbl_Solve.States = Tbl_TypeStatics.TypeID AND  ISDN = '" & strISDN & "'"
            olecommand.Connection = oledbcon
            oleAdtapter.SelectCommand = olecommand
            oleAdtapter.Fill(table)
            table.TableName = "Chitiet"
            oleAdtapter.Dispose()
            olecommand.Dispose()
            DataGridDskh.DataSource = table
        Catch ex As Exception
            MsgBox("Lổi rồi người ơi!!" & ex.ToString)
        End Try
    End Sub
    Private Sub txtSomay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSomay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            KeyAscii = 0
            If (Trim$(txtSomay.Text) <> "") Then
                strISDN = Trim$(txtSomay.Text)
                LoadDetailISDN()
            End If
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub LoadInfo()
        txtSomay.Text = strISDN
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = " SELECT DISTINCT Customer_Name,Customer_Address FROM tbl_Solve WHERE ISDN = '" & strISDN & "'"
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then
                If Not oleread.IsDBNull(0) Then
                    txtHotenkh.Text = oleread.GetString(0)
                End If

                If Not oleread.IsDBNull(1) Then
                    txtDiachi.Text = oleread.GetString(1)
                End If
            End If
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox("Lỗi rồi người ơi :")
        End Try
        oledbcon.Close()

    End Sub

    Private Sub cmdclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdclose.Click
        Me.Close()
    End Sub
End Class
