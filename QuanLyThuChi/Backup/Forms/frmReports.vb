Imports System.Data.OleDb
Public Class frmReports
    Inherits System.Windows.Forms.Form
    Private mydataset As DataSet
    Private table As New DataTable
    Dim start As Boolean = False
    Dim Dsrpt As New DsSoQuy
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'FillDataSet()
        start = True
        DateTimePickerdenngay.Value = Now
        DateTimePickertungay.Value = Now
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
    Friend WithEvents DateTimePickerdenngay As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePickertungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdxem As System.Windows.Forms.Button
    Friend WithEvents cmdclose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmReports))
        Me.DateTimePickerdenngay = New System.Windows.Forms.DateTimePicker
        Me.DateTimePickertungay = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdxem = New System.Windows.Forms.Button
        Me.cmdclose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DateTimePickerdenngay
        '
        Me.DateTimePickerdenngay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickerdenngay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerdenngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerdenngay.Location = New System.Drawing.Point(245, 18)
        Me.DateTimePickerdenngay.Name = "DateTimePickerdenngay"
        Me.DateTimePickerdenngay.Size = New System.Drawing.Size(98, 26)
        Me.DateTimePickerdenngay.TabIndex = 75
        Me.DateTimePickerdenngay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'DateTimePickertungay
        '
        Me.DateTimePickertungay.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePickertungay.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickertungay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickertungay.Location = New System.Drawing.Point(67, 18)
        Me.DateTimePickertungay.Name = "DateTimePickertungay"
        Me.DateTimePickertungay.Size = New System.Drawing.Size(98, 26)
        Me.DateTimePickertungay.TabIndex = 66
        Me.DateTimePickertungay.Value = New Date(2005, 6, 23, 0, 0, 0, 0)
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label8.Location = New System.Drawing.Point(173, 21)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(69, 22)
        Me.Label8.TabIndex = 76
        Me.Label8.Text = "Đến ngày"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(5, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 22)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Từ ngày"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DateTimePickerdenngay)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.DateTimePickertungay)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(348, 56)
        Me.GroupBox1.TabIndex = 77
        Me.GroupBox1.TabStop = False
        '
        'cmdxem
        '
        Me.cmdxem.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdxem.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdxem.Location = New System.Drawing.Point(77, 64)
        Me.cmdxem.Name = "cmdxem"
        Me.cmdxem.Size = New System.Drawing.Size(80, 27)
        Me.cmdxem.TabIndex = 78
        Me.cmdxem.Text = "Xem BC"
        '
        'cmdclose
        '
        Me.cmdclose.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdclose.ForeColor = System.Drawing.SystemColors.Desktop
        Me.cmdclose.Location = New System.Drawing.Point(205, 64)
        Me.cmdclose.Name = "cmdclose"
        Me.cmdclose.Size = New System.Drawing.Size(80, 27)
        Me.cmdclose.TabIndex = 79
        Me.cmdclose.Text = "Đóng"
        '
        'frmReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(362, 104)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdxem)
        Me.Controls.Add(Me.cmdclose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Báo cáo quỹ"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Function GetTable() As DataTable
        Return table
    End Function


    Private Sub cmdclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdclose.Click
        Me.Close()
    End Sub

    Private Sub cmdxem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdxem.Click
        Dim rpt As CrystalReportBaoCaoQuy
        rpt = New CrystalReportBaoCaoQuy
        Dsrpt.Clear()

        'Fill Tu ngay den ngay
        Dim Newrow As DataRow
        Newrow = Dsrpt.Tables("NgayBaoCao").NewRow

        Dsrpt.Tables("NgayBaoCao").Rows.Add(Newrow)
        Dsrpt.Tables("NgayBaoCao").Rows(0).Item("TuNgay") = DateTimePickertungay.Text
        Dsrpt.Tables("NgayBaoCao").Rows(0).Item("Denngay") = DateTimePickerdenngay.Text

        Dim strQuery As String
        'Fill Noi dung so quy
        strQuery = "SELECT Tbl_Receipts_Expenses.ID, Tbl_Receipts_Expenses.Recei_Expen_Date, Tbl_Receipts_Expenses.Recei_No, Tbl_Receipts_Expenses.Expen_No, Tbl_Receipts_Expenses.Descriptions, Tbl_Receipts_Expenses.Recei_Money, Tbl_Receipts_Expenses.Expen_Money, [Recei_Money]-[Expen_Money] AS TonQuy FROM Tbl_Receipts_Expenses WHERE Recei_Expen_Date BETWEEN #" & DateTimePickertungay.Value.ToShortDateString & "# AND #" & DateTimePickerdenngay.Value.ToShortDateString & "# Order By Recei_Expen_Date,Tbl_Receipts_Expenses.Recei_No "
        FillReports(strQuery, "QueryReport_SoQuy")

        'Fill TondauKy
        strQuery = "SELECT SUM([Recei_Money]-[Expen_Money]) AS DauKy FROM Tbl_Receipts_Expenses WHERE Recei_Expen_Date < #" & DateTimePickertungay.Value.ToShortDateString & "# "
        FillReports(strQuery, "QueryTonDauKy")
        'Fill Ton cuoi ky
        strQuery = "SELECT SUM([Recei_Money]-[Expen_Money]) AS CuoiKy FROM Tbl_Receipts_Expenses  "
        FillReports(strQuery, "QueryTonCuoiKy")

        rpt.SetDataSource(Dsrpt)
        Dim frm As New frmPreview
        frm.CrystalReportViewerReceipts.ReportSource = rpt
        frm.ShowDialog()
    End Sub

    Private Sub FillReports(ByVal strQuery As String, ByVal strTableName As String)
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(Dsrpt, strTableName)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    Private Sub DateTimePickertungay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickertungay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            DateTimePickerdenngay.Focus()
        End If
    End Sub

    Private Sub DateTimePickerdenngay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePickerdenngay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii = 13) Then
            cmdxem.Focus()
        End If
    End Sub
End Class
