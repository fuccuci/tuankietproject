Imports System.Data.OleDb
Public Class frmBaoCaoThuCuoc
    Inherits System.Windows.Forms.Form
    Private dsDauKy As DataSet
    'Dim dsDaThu As dataset
    Private SumTien As Double
    Dim splitn As New SplitNumbers

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
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabSoLieu As System.Windows.Forms.TabControl
    Friend WithEvents dgridDauKy As System.Windows.Forms.DataGrid
    Friend WithEvents dgridKHDaThu As System.Windows.Forms.DataGrid
    Friend WithEvents dgridKHChuaThu As System.Windows.Forms.DataGrid
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents dgridTyLeThu As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TabSoLieu = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgridDauKy = New System.Windows.Forms.DataGrid
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgridKHDaThu = New System.Windows.Forms.DataGrid
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgridKHChuaThu = New System.Windows.Forms.DataGrid
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.dgridTyLeThu = New System.Windows.Forms.DataGrid
        Me.TabSoLieu.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgridDauKy, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgridKHDaThu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgridKHChuaThu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.dgridTyLeThu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabSoLieu
        '
        Me.TabSoLieu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabSoLieu.Controls.Add(Me.TabPage1)
        Me.TabSoLieu.Controls.Add(Me.TabPage2)
        Me.TabSoLieu.Controls.Add(Me.TabPage3)
        Me.TabSoLieu.Controls.Add(Me.TabPage4)
        Me.TabSoLieu.Location = New System.Drawing.Point(4, 4)
        Me.TabSoLieu.Name = "TabSoLieu"
        Me.TabSoLieu.SelectedIndex = 0
        Me.TabSoLieu.Size = New System.Drawing.Size(775, 388)
        Me.TabSoLieu.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgridDauKy)
        Me.TabPage1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(767, 359)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Khách hàng đầu kỳ"
        '
        'dgridDauKy
        '
        Me.dgridDauKy.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgridDauKy.DataMember = ""
        Me.dgridDauKy.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgridDauKy.Location = New System.Drawing.Point(0, 4)
        Me.dgridDauKy.Name = "dgridDauKy"
        Me.dgridDauKy.Size = New System.Drawing.Size(767, 356)
        Me.dgridDauKy.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgridKHDaThu)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(767, 362)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Khách hàng đã thu"
        '
        'dgridKHDaThu
        '
        Me.dgridKHDaThu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgridKHDaThu.DataMember = ""
        Me.dgridKHDaThu.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgridKHDaThu.Location = New System.Drawing.Point(0, 1)
        Me.dgridKHDaThu.Name = "dgridKHDaThu"
        Me.dgridKHDaThu.Size = New System.Drawing.Size(767, 359)
        Me.dgridKHDaThu.TabIndex = 1
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgridKHChuaThu)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(767, 362)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Khách hàng chưa thu"
        '
        'dgridKHChuaThu
        '
        Me.dgridKHChuaThu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgridKHChuaThu.DataMember = ""
        Me.dgridKHChuaThu.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgridKHChuaThu.Location = New System.Drawing.Point(0, 1)
        Me.dgridKHChuaThu.Name = "dgridKHChuaThu"
        Me.dgridKHChuaThu.Size = New System.Drawing.Size(767, 359)
        Me.dgridKHChuaThu.TabIndex = 2
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.dgridTyLeThu)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(767, 362)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Tỷ lệ thu cước"
        '
        'dgridTyLeThu
        '
        Me.dgridTyLeThu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgridTyLeThu.DataMember = ""
        Me.dgridTyLeThu.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgridTyLeThu.Location = New System.Drawing.Point(0, 1)
        Me.dgridTyLeThu.Name = "dgridTyLeThu"
        Me.dgridTyLeThu.Size = New System.Drawing.Size(767, 359)
        Me.dgridTyLeThu.TabIndex = 3
        '
        'frmBaoCaoThuCuoc
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(784, 397)
        Me.Controls.Add(Me.TabSoLieu)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmBaoCaoThuCuoc"
        Me.Text = "Thống Kê Tình Hình Thu Cước Dịch Vụ 178"
        Me.TabSoLieu.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgridDauKy, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgridKHDaThu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgridKHChuaThu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.dgridTyLeThu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FormatDataGridDauKy()

        With dgridDauKy
            .AllowNavigation = False
            '.DataMember = "Tbl_SoGiaoDauKy"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách khách hàng đầu kỳ"
            .Font = New System.Drawing.Font("Times New Roman", 11.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "tblChildDauKy"
            .BackColor = System.Drawing.Color.MintCream
            .ForeColor = System.Drawing.Color.Navy
            .GridLineColor = System.Drawing.Color.MediumBlue
            .HeaderBackColor = System.Drawing.Color.Lavender
            .HeaderForeColor = System.Drawing.Color.Navy
            .AlternatingBackColor = Color.LightGray
            .HeaderFont = New System.Drawing.Font("Times New Roman", 11.0!, FontStyle.Bold)
            .RowHeaderWidth = 10
            .ReadOnly = False

            With .GridColumnStyles
                ' Set datagrid ColumnStyle for ID field

                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "MaNV"
                    .HeaderText = "Mã CTV   "
                    .Width = 150
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "SoTB"
                    .HeaderText = "Số thuê bao"
                    .Width = 120
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With
             
                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "NoTruoc"
                    .HeaderText = "Nợ trước"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "DieuChinh"
                    .HeaderText = "Ðiều chỉnh"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "PhatSinh"
                    .HeaderText = "Phát sinh"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "Thue"
                    .HeaderText = "Thuế"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(6)
                    .MappingName = "TongCuoc"
                    .HeaderText = "Tổng cước"
                    .Width = 150
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With
            End With
        End With
        dgridDauKy.TableStyles.Add(TblStyle)
        'dgridDauKy.TableStyles.Add(TblStyle1)
    End Sub

    'Private Sub FormatDataGridKHDaThu()

    '    With dgridDauKy
    '        .AllowNavigation = False
    '        .DataMember = "Tbl_SoLieuBaoCao"
    '        .BackgroundColor = System.Drawing.Color.LightSteelBlue
    '        .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
    '        .CaptionForeColor = System.Drawing.Color.MediumBlue
    '        .ParentRowsBackColor = System.Drawing.Color.Lavender
    '        .ParentRowsForeColor = System.Drawing.Color.SlateBlue
    '        .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
    '        .CaptionText = "Danh sách khách hàng dã thu"
    '        .Font = New System.Drawing.Font("Times New Roman", 11.0!)
    '    End With

    '    Dim TblStyle As New DataGridTableStyle
    '    With TblStyle
    '        .MappingName = "tblChildDaThu"
    '        .BackColor = System.Drawing.Color.MintCream
    '        .ForeColor = System.Drawing.Color.Navy
    '        .GridLineColor = System.Drawing.Color.MediumBlue
    '        .HeaderBackColor = System.Drawing.Color.Lavender
    '        .HeaderForeColor = System.Drawing.Color.Navy
    '        .AlternatingBackColor = Color.LightGray
    '        .HeaderFont = New System.Drawing.Font("Times New Roman", 11.0!, FontStyle.Bold)
    '        .RowHeaderWidth = 10
    '        .ReadOnly = False

    '        With .GridColumnStyles
    '             Set datagrid ColumnStyle for ID field

    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(0)
    '                .MappingName = "MaNV"
    '                .HeaderText = "Mã CTV   "
    '                .Width = 150
    '                .Alignment = HorizontalAlignment.Left
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With

    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(1)
    '                .MappingName = "SoTB"
    '                .HeaderText = "Số thuê bao"
    '                .Width = 120
    '                .Alignment = HorizontalAlignment.Left
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With
    '            .Add(New DataGridDateTimePicker)
    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(2)
    '                .MappingName = "TenKH"
    '                .HeaderText = "Tên khách hàng"
    '                .Width = 200
    '                .Alignment = HorizontalAlignment.Left
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With

    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(3)
    '                .MappingName = "SoHD"
    '                .HeaderText = "Số hóa đơn"
    '                .Width = 150
    '                .Alignment = HorizontalAlignment.Left
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With

    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(4)
    '                .MappingName = "TuNgay"
    '                .HeaderText = "Từ ngày"
    '                .Width = 100
    '                .Alignment = HorizontalAlignment.Right
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With

    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(5)
    '                .MappingName = "DenNgay"
    '                .HeaderText = "Ðến ngày"
    '                .Width = 100
    '                .Alignment = HorizontalAlignment.Right
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With

    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(6)
    '                .MappingName = "SoTienTra"
    '                .HeaderText = "Số tiền trả"
    '                .Width = 150
    '                .Alignment = HorizontalAlignment.Right
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With

    '            .Add(New DataGridTextBoxColumn)
    '            With .Item(7)
    '                .MappingName = "NgayTra"
    '                .HeaderText = "Ngày trả"
    '                .Width = 100
    '                .Alignment = HorizontalAlignment.Right
    '                .NullText = String.Empty
    '                .ReadOnly = True
    '            End With

    '        End With
    '    End With
    '    dgridKHDaThu.TableStyles.Add(TblStyle)
    'End Sub

    Private Sub FormatDataGridChuaThu()

        With dgridKHChuaThu
            .AllowNavigation = False
            '.DataMember = "Tbl_SoGiaoDauKy"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Danh sách khách hàng chưa thu"
            .Font = New System.Drawing.Font("Times New Roman", 11.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "tblChildChuaThu"
            .BackColor = System.Drawing.Color.MintCream
            .ForeColor = System.Drawing.Color.Navy
            .GridLineColor = System.Drawing.Color.MediumBlue
            .HeaderBackColor = System.Drawing.Color.Lavender
            .HeaderForeColor = System.Drawing.Color.Navy
            .AlternatingBackColor = Color.LightGray
            .HeaderFont = New System.Drawing.Font("Times New Roman", 11.0!, FontStyle.Bold)
            .RowHeaderWidth = 10
            .ReadOnly = False

            With .GridColumnStyles
                ' Set datagrid ColumnStyle for ID field

                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "MaNV"
                    .HeaderText = "Mã CTV   "
                    .Width = 150
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "SoTB"
                    .HeaderText = "Số thuê bao"
                    .Width = 120
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "NoTruoc"
                    .HeaderText = "Nợ trước"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "DieuChinh"
                    .HeaderText = "Ðiều chỉnh"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "PhatSinh"
                    .HeaderText = "Phát sinh"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "Thue"
                    .HeaderText = "Thuế"
                    .Width = 100
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(6)
                    .MappingName = "TongCuoc"
                    .HeaderText = "Tổng cước"
                    .Width = 150
                    .Alignment = HorizontalAlignment.Right
                    .NullText = Long.MinValue
                    .ReadOnly = True
                End With

            End With
        End With
        dgridKHChuaThu.TableStyles.Add(TblStyle)
    End Sub



    Private Sub FormatDataGridTyLe()

        With dgridTyLeThu
            .AllowNavigation = True
            '.DataMember = "Tbl_SoGiaoDauKy"
            .BackgroundColor = System.Drawing.Color.LightSteelBlue
            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
            .CaptionForeColor = System.Drawing.Color.MediumBlue
            .ParentRowsBackColor = System.Drawing.Color.Lavender
            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
            .CaptionText = "Tỷ lệ thu cước của từng CTV"
            .Font = New System.Drawing.Font("Times New Roman", 11.0!)
        End With

        Dim TblStyle As New DataGridTableStyle
        With TblStyle
            .MappingName = "Tbl_TyLeThu"
            .BackColor = System.Drawing.Color.MintCream
            .ForeColor = System.Drawing.Color.Navy
            .GridLineColor = System.Drawing.Color.MediumBlue
            .HeaderBackColor = System.Drawing.Color.Lavender
            .HeaderForeColor = System.Drawing.Color.Navy
            .AlternatingBackColor = Color.LightGray
            .HeaderFont = New System.Drawing.Font("Times New Roman", 11.0!, FontStyle.Bold)
            .RowHeaderWidth = 10
            .ReadOnly = False

            With .GridColumnStyles
                ' Set datagrid ColumnStyle for ID field

                .Add(New DataGridTextBoxColumn)
                With .Item(0)
                    .MappingName = "MaNV"
                    .HeaderText = "Mã CTV   "
                    .Width = 150
                    .Alignment = HorizontalAlignment.Left
                    .NullText = String.Empty
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(1)
                    .MappingName = "TongTB"
                    .HeaderText = "Tồng Số KH"
                    .Width = 60
                    .Alignment = HorizontalAlignment.Right
                    .NullText = "0"
                    .ReadOnly = True
                End With
                '.Add(New DataGridDateTimePicker)
                .Add(New DataGridTextBoxColumn)
                With .Item(2)
                    .MappingName = "TongCuoc"
                    .HeaderText = "Tổng tiến giao"
                    .Width = 60
                    .Alignment = HorizontalAlignment.Right
                    .NullText = "0"
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(3)
                    .MappingName = "TongTBDT"
                    .HeaderText = "Tổng KHDT"
                    .Width = 60
                    .Alignment = HorizontalAlignment.Right
                    .NullText = "0"
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(4)
                    .MappingName = "TongTienDT"
                    .HeaderText = "Tổng tiền thu"
                    .Width = 60
                    .Alignment = HorizontalAlignment.Right
                    .NullText = "0"
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(5)
                    .MappingName = "TyLeTB"
                    .HeaderText = "Tỷ lệ TB"
                    .Width = 60
                    .Alignment = HorizontalAlignment.Right
                    .NullText = "0%"
                    .ReadOnly = True
                End With

                .Add(New DataGridTextBoxColumn)
                With .Item(6)
                    .MappingName = "TyleTien"
                    .HeaderText = "Tỷ lệ tiền"
                    .Width = 60
                    .Alignment = HorizontalAlignment.Right
                    .NullText = "0%"
                    .ReadOnly = True
                End With
            End With
        End With
        dgridTyLeThu.TableStyles.Add(TblStyle)
    End Sub

    Private Function SumMoney(ByVal ds As DataSet, ByVal tablename As String, ByVal columnname As String) As Double
        Dim table As New DataTable
        'Dim ds As New DataSet
        Dim result As Double = 0
        Try
            'ds = dgrid.DataSource
            table = ds.Tables(tablename)
            Dim i As Integer
            For i = 0 To table.Rows.Count - 1
                result += table.Rows(i).Item(columnname)
            Next
        Catch ex As Exception

        End Try

        Return result
    End Function
    Private Sub FillDataset(ByVal dgrid As DataGrid, ByVal strQuery As String, ByVal tblName As String) 'ByVal 'dgrid As DataGrid,
        Try

            'dsDauKy = New dataset
            dgrid.DataSource = Nothing
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            da = New OleDbDataAdapter(cmd)
            da.Fill(dsDauKy, tblName)
            dgrid.DataSource = dsDauKy.Tables(tblName)
        Catch ex As Exception
            MsgBox("Error Fill DataSet :" & ex.ToString)
        End Try
    End Sub

    Private Sub FillDataCommand(ByVal strQuery As String, ByVal tblName As String)
        Try

            Dim cmd As New OleDbCommand
            cmd.Connection = oledbcon
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = strQuery
            da = New OleDbDataAdapter(cmd)
            da.Fill(dsDauKy, tblName)
        Catch ex As Exception
            MsgBox("Error Fill DataSet :" & ex.ToString)
        End Try
    End Sub

    Private Sub frmBaoCaoThuCuoc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dsDauKy = New DataSet
        'FormatDataGridDauKy()
        'FormatDataGridKHDaThu()
        'FormatDataGridChuaThu()
        FormatDataGridTyLe()
        strSQL = "SELECT Tbl_SoGiaoDauKy.MaNV, Count(Tbl_SoGiaoDauKy.SoTB) AS SoTB, Sum(Tbl_SoGiaoDauKy.NoTruoc) AS NoTruoc, Sum(Tbl_SoGiaoDauKy.DieuChinh) AS DieuChinh, Sum(Tbl_SoGiaoDauKy.PhatSinh) AS PhatSinh, Sum(Tbl_SoGiaoDauKy.Thue) AS Thue, Sum(Tbl_SoGiaoDauKy.TongCuoc) AS TongCuoc"
        strSQL = strSQL & " FROM Tbl_SoGiaoDauKy"
        strSQL = strSQL & " GROUP BY Tbl_SoGiaoDauKy.MaNV"
        Dim strSQL1 = "SELECT MaNV, SoTB, TenKH, DiaChi, NoTruoc, DieuChinh, PhatSinh, Thue, TongCuoc FROM Tbl_SoGiaoDauKy ORDER BY MaNV"
        Dim m_da_tblName1 As New OleDbDataAdapter(strSQL, oledbcon)
        Dim m_da_tblName2 As New OleDbDataAdapter(strSQL1, oledbcon)
        Try
            m_da_tblName1.Fill(dsDauKy, "Tbl_SoGiaoDauKy")
            m_da_tblName2.Fill(dsDauKy, "Tbl_SoGiaoDauKy1")

            Dim data_relation As New DataRelation("Xem chi tiết", _
                dsDauKy.Tables("Tbl_SoGiaoDauKy").Columns("MANV"), _
                dsDauKy.Tables("Tbl_SoGiaoDauKy1").Columns("MANV"))
            dsDauKy.Relations.Add(data_relation)
            dgridDauKy.SetDataBinding(dsDauKy, "Tbl_SoGiaoDauKy")
        Catch ex As Exception
        End Try
        splitn.strnumbers = CStr(SumMoney(dsDauKy, "Tbl_SoGiaoDauKy", "TongCuoc"))
        dgridDauKy.CaptionText = "Danh sách khách hàng đầu kỳ:" & CStr(SumMoney(dsDauKy, "Tbl_SoGiaoDauKy", "SoTB")) + " KH" + ", " + "Tổng tiền: " & splitn.Splitnumer(",")
        strSQL = ""
        strSQL1 = ""
        'dsDauKy.Tables("Tbl_SogiaoDauKy").Clear()
    End Sub

    Private Sub TabSoLieu_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabSoLieu.SelectedIndexChanged
        Select Case TabSoLieu.SelectedIndex
            Case 0
                dsDauKy = New DataSet
                dsDauKy.Clear()
                strSQL = "SELECT Tbl_SoGiaoDauKy.MaNV, Count(Tbl_SoGiaoDauKy.SoTB) AS SoTB, Sum(Tbl_SoGiaoDauKy.NoTruoc) AS NoTruoc, Sum(Tbl_SoGiaoDauKy.DieuChinh) AS DieuChinh, Sum(Tbl_SoGiaoDauKy.PhatSinh) AS PhatSinh, Sum(Tbl_SoGiaoDauKy.Thue) AS Thue, Sum(Tbl_SoGiaoDauKy.TongCuoc) AS TongCuoc"
                strSQL = strSQL & " FROM Tbl_SoGiaoDauKy"
                strSQL = strSQL & " GROUP BY Tbl_SoGiaoDauKy.MaNV"
                Dim strSQL1 = "SELECT MaNV, SoTB, TenKH, DiaChi, NoTruoc, DieuChinh, PhatSinh, Thue, TongCuoc FROM Tbl_SoGiaoDauKy ORDER BY MaNV"
                Dim m_da_tblName1 As New OleDbDataAdapter(strSQL, oledbcon)
                Dim m_da_tblName2 As New OleDbDataAdapter(strSQL1, oledbcon)
                Try
                    m_da_tblName1.Fill(dsDauKy, "tblChildDauKy")
                    m_da_tblName2.Fill(dsDauKy, "tblChildDauKy1")

                    Dim data_relation As New DataRelation("Xem chi tiết", _
                        dsDauKy.Tables("tblChildDauKy").Columns("MANV"), _
                        dsDauKy.Tables("tblChildDauKy1").Columns("MANV"))
                    dsDauKy.Relations.Add(data_relation)
                    dgridDauKy.SetDataBinding(dsDauKy, "tblChildDauKy")
                Catch ex As Exception
                End Try
                splitn.strnumbers = CStr(SumMoney(dsDauKy, "tblChildDauKy", "TongCuoc"))
                dgridDauKy.CaptionText = "Danh sách khách hàng đầu kỳ:" & CStr(SumMoney(dsDauKy, "tblChildDauKy", "SoTB")) + " KH" + ", " + "Tổng tiền: " & splitn.Splitnumer(",")
                strSQL = ""
                strSQL1 = ""

            Case 1
                dsDauKy = New DataSet
                dsDauKy.Clear()

                strSQL = "SELECT MaNV, Count(Tbl_SoLieuBaoCao.SoTB) AS SoTB, Sum(Tbl_SoLieuBaoCao.SoTienTra) AS SoTien"
                strSQL = strSQL & " FROM Tbl_SoLieuBaoCao"
                strSQL = strSQL & " GROUP BY Tbl_SoLieuBaoCao.MaNV"
                Dim strSQL1 = "SELECT MaNV,SoTB ,TenKH ,SoHD ,TuNgay , DenNgay ,SoTienTra , NgayTra FROM Tbl_SoLieuBaoCao ORDER BY MaNV "
                Dim m_da_tblName1 As New OleDbDataAdapter(strSQL, oledbcon)
                Dim m_da_tblName2 As New OleDbDataAdapter(strSQL1, oledbcon)
                Try
                    m_da_tblName1.Fill(dsDauKy, "tblChildDaThu")
                    m_da_tblName2.Fill(dsDauKy, "tblChildDaThu1")

                    Dim data_relation As New DataRelation("Xem chi tiết", _
                        dsDauKy.Tables("tblChildDaThu").Columns("MANV"), _
                        dsDauKy.Tables("tblChildDaThu1").Columns("MANV"))
                    dsDauKy.Relations.Add(data_relation)
                    dgridKHDaThu.SetDataBinding(dsDauKy, "tblChildDaThu")
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                splitn.strnumbers = CStr(SumMoney(dsDauKy, "tblChildDaThu", "SoTien"))
                dgridKHDaThu.CaptionText = "Danh sách khách hàng đã thu:" & CStr(SumMoney(dsDauKy, "tblChildDaThu", "SoTB")) + " KH" + ", " + "Tổng tiền: " & splitn.Splitnumer(",")
                strSQL = ""
                strSQL1 = ""
            Case 2
                dsDauKy = New DataSet
                dsDauKy.Clear()
                strSQL = " SELECT Tbl_SoGiaoDauKy.MaNV, Count(Tbl_SoGiaoDauKy.SoTB) AS SoTB, Sum(Tbl_SoGiaoDauKy.NoTruoc) AS NoTruoc, Sum(Tbl_SoGiaoDauKy.DieuChinh) AS DieuChinh, Sum(Tbl_SoGiaoDauKy.PhatSinh) AS PhatSinh, Sum(Tbl_SoGiaoDauKy.Thue) AS Thue, Sum(Tbl_SoGiaoDauKy.TongCuoc) AS TongCuoc"
                strSQL = strSQL & " FROM Tbl_SoGiaoDauKy LEFT JOIN Tbl_SoLieuBaoCao ON Tbl_SoGiaoDauKy.SoTB = Tbl_SoLieuBaoCao.SoTB"
                strSQL = strSQL & " GROUP BY Tbl_SoGiaoDauKy.MaNV, Tbl_SoLieuBaoCao.SoTB"
                strSQL = strSQL & " HAVING(((Tbl_SoLieuBaoCao.SoTB) Is Null))"

                Dim strSQL1 = "SELECT MaNV, SoTB, TenKH, DiaChi, NoTruoc, DieuChinh, PhatSinh, Thue, TongCuoc FROM Tbl_SoGiaoDauKy ORDER BY MaNV"
                Dim m_da_tblName1 As New OleDbDataAdapter(strSQL, oledbcon)
                Dim m_da_tblName2 As New OleDbDataAdapter(strSQL1, oledbcon)
                Try
                    m_da_tblName1.Fill(dsDauKy, "tblChildChuaThu")
                    m_da_tblName2.Fill(dsDauKy, "tblChildChuaThu1")

                    Dim data_relation As New DataRelation("Xem chi tiết", _
                        dsDauKy.Tables("tblChildChuaThu").Columns("MANV"), _
                        dsDauKy.Tables("tblChildChuaThu1").Columns("MANV"))
                    dsDauKy.Relations.Add(data_relation)
                    dgridKHChuaThu.SetDataBinding(dsDauKy, "tblChildChuaThu")
                Catch ex As Exception
                End Try
                splitn.strnumbers = CStr(SumMoney(dsDauKy, "tblChildChuaThu", "TongCuoc"))
                dgridKHChuaThu.CaptionText = "Danh sách khách hàng đầu kỳ:" & CStr(SumMoney(dsDauKy, "tblChildChuaThu", "SoTB")) + " KH" + ", " + "Tổng tiền: " & splitn.Splitnumer(",")
                strSQL = ""
                strSQL1 = ""

            Case 3
                dgridTyLeThu.DataSource = Nothing
                FillDataCommand("qryTyLe", "Tbl_TyLeThu")
                dgridTyLeThu.DataSource = dsDauKy.Tables("Tbl_TyLeThu")
                FillDataCommand("qryTyLeTram", "Tbl_TyLeThuTram")
                dgridTyLeThu.CaptionText = "Tỷ lệ thu cùa trạm:" + "KH: " & CStr(dsDauKy.Tables("Tbl_TyLeThuTram").Rows(0)("TyLeTB")) + "%" + ", " + "Tổng tiền: " & CStr(dsDauKy.Tables("Tbl_TyLeThuTram").Rows(0)("TyLeTien")) + "%"
        End Select

    End Sub

End Class
