'Imports System.Data.OleDb
'Imports System.ComponentModel
'Imports QuanLyCTV.DataGridTextBoxCombo
'Public Class frmStatistics
'    Inherits System.Windows.Forms.Form
'    Private mydataset As DataSet
'    Friend WithEvents OleDbSelectCust As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbInsertCust As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbUpdateCust As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbDeleteCust As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbSelectSolve As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbInsertSolve As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbUpdateSolve As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbDeleteSolve As System.Data.OleDb.OleDbCommand
'    Friend WithEvents OleDbDataAdapterCust As System.Data.OleDb.OleDbDataAdapter
'    Friend WithEvents OleDbDataAdapterSolve As System.Data.OleDb.OleDbDataAdapter
'    Friend WithEvents objdsCustomers As QuanLyCTV.dsCustomers
'#Region " Windows Form Designer generated code "

'    Public Sub New()
'        MyBase.New()

'        'This call is required by the Windows Form Designer.
'        InitializeComponent()

'        'Add any initialization after the InitializeComponent() call

'    End Sub

'    'Form overrides dispose to clean up the component list.
'    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
'        If disposing Then
'            If Not (components Is Nothing) Then
'                components.Dispose()
'            End If
'        End If
'        MyBase.Dispose(disposing)
'    End Sub

'    'Required by the Windows Form Designer
'    Private components As System.ComponentModel.IContainer

'    'NOTE: The following procedure is required by the Windows Form Designer
'    'It can be modified using the Windows Form Designer.  
'    'Do not modify it using the code editor.
'    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
'    Friend WithEvents DataGridDskh As System.Windows.Forms.DataGrid
'    Friend WithEvents Label4 As System.Windows.Forms.Label
'    Friend WithEvents ContextMenuGrid As System.Windows.Forms.ContextMenu
'    Friend WithEvents MenuItemNewComplain As System.Windows.Forms.MenuItem
'    Friend WithEvents Label1 As System.Windows.Forms.Label
'    Friend WithEvents cmdclose As System.Windows.Forms.Button
'    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
'    Friend WithEvents DataGridSovle As System.Windows.Forms.DataGrid
'    Friend WithEvents Label2 As System.Windows.Forms.Label
'    Friend WithEvents MenuItemUpdateComplain As System.Windows.Forms.MenuItem
'    Friend WithEvents MenuItemDelComplain As System.Windows.Forms.MenuItem
'    Friend WithEvents MenuItemTim As System.Windows.Forms.MenuItem
'    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
'        Me.GroupBox1 = New System.Windows.Forms.GroupBox
'        Me.DataGridDskh = New System.Windows.Forms.DataGrid
'        Me.ContextMenuGrid = New System.Windows.Forms.ContextMenu
'        Me.MenuItemNewComplain = New System.Windows.Forms.MenuItem
'        Me.MenuItemUpdateComplain = New System.Windows.Forms.MenuItem
'        Me.MenuItemDelComplain = New System.Windows.Forms.MenuItem
'        Me.MenuItemTim = New System.Windows.Forms.MenuItem
'        Me.Label4 = New System.Windows.Forms.Label
'        Me.Label1 = New System.Windows.Forms.Label
'        Me.cmdclose = New System.Windows.Forms.Button
'        Me.GroupBox2 = New System.Windows.Forms.GroupBox
'        Me.DataGridSovle = New System.Windows.Forms.DataGrid
'        Me.Label2 = New System.Windows.Forms.Label
'        Me.GroupBox1.SuspendLayout()
'        CType(Me.DataGridDskh, System.ComponentModel.ISupportInitialize).BeginInit()
'        Me.GroupBox2.SuspendLayout()
'        CType(Me.DataGridSovle, System.ComponentModel.ISupportInitialize).BeginInit()
'        Me.SuspendLayout()
'        '
'        'GroupBox1
'        '
'        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
'                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.GroupBox1.Controls.Add(Me.DataGridDskh)
'        Me.GroupBox1.Location = New System.Drawing.Point(4, 95)
'        Me.GroupBox1.Name = "GroupBox1"
'        Me.GroupBox1.Size = New System.Drawing.Size(754, 241)
'        Me.GroupBox1.TabIndex = 0
'        Me.GroupBox1.TabStop = False
'        '
'        'DataGridDskh
'        '
'        Me.DataGridDskh.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
'                    Or System.Windows.Forms.AnchorStyles.Left) _
'                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.DataGridDskh.CaptionText = "Danh sách khách hàng"
'        Me.DataGridDskh.ContextMenu = Me.ContextMenuGrid
'        Me.DataGridDskh.DataMember = ""
'        Me.DataGridDskh.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
'        Me.DataGridDskh.HeaderForeColor = System.Drawing.SystemColors.ControlText
'        Me.DataGridDskh.Location = New System.Drawing.Point(7, 11)
'        Me.DataGridDskh.Name = "DataGridDskh"
'        Me.DataGridDskh.Size = New System.Drawing.Size(742, 220)
'        Me.DataGridDskh.TabIndex = 21
'        '
'        'ContextMenuGrid
'        '
'        Me.ContextMenuGrid.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemNewComplain, Me.MenuItemUpdateComplain, Me.MenuItemDelComplain, Me.MenuItemTim})
'        '
'        'MenuItemNewComplain
'        '
'        Me.MenuItemNewComplain.Index = 0
'        Me.MenuItemNewComplain.Text = "&Thêm "
'        '
'        'MenuItemUpdateComplain
'        '
'        Me.MenuItemUpdateComplain.Index = 1
'        Me.MenuItemUpdateComplain.Text = "&Cập Nhật "
'        '
'        'MenuItemDelComplain
'        '
'        Me.MenuItemDelComplain.Index = 2
'        Me.MenuItemDelComplain.Text = "&Xóa"
'        '
'        'MenuItemTim
'        '
'        Me.MenuItemTim.Index = 3
'        Me.MenuItemTim.Text = "&Tìm"
'        '
'        'Label4
'        '
'        Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
'                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.Label4.BackColor = System.Drawing.SystemColors.ControlLightLight
'        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
'        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
'        Me.Label4.Font = New System.Drawing.Font("VNI-Allegie", 35.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
'        Me.Label4.ForeColor = System.Drawing.SystemColors.InactiveCaption
'        Me.Label4.Location = New System.Drawing.Point(6, 6)
'        Me.Label4.Name = "Label4"
'        Me.Label4.Size = New System.Drawing.Size(746, 73)
'        Me.Label4.TabIndex = 11
'        Me.Label4.Text = "Giaûi Quyeát Khieáu Naïi Khaùch Haøng"
'        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'        '
'        'Label1
'        '
'        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
'                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.Label1.BackColor = System.Drawing.Color.Navy
'        Me.Label1.Location = New System.Drawing.Point(6, 87)
'        Me.Label1.Name = "Label1"
'        Me.Label1.Size = New System.Drawing.Size(740, 3)
'        Me.Label1.TabIndex = 12
'        '
'        'cmdclose
'        '
'        Me.cmdclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.cmdclose.Location = New System.Drawing.Point(674, 535)
'        Me.cmdclose.Name = "cmdclose"
'        Me.cmdclose.Size = New System.Drawing.Size(80, 24)
'        Me.cmdclose.TabIndex = 13
'        Me.cmdclose.Text = "&Đóng"
'        '
'        'GroupBox2
'        '
'        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
'                    Or System.Windows.Forms.AnchorStyles.Left) _
'                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.GroupBox2.Controls.Add(Me.DataGridSovle)
'        Me.GroupBox2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
'        Me.GroupBox2.Location = New System.Drawing.Point(4, 355)
'        Me.GroupBox2.Name = "GroupBox2"
'        Me.GroupBox2.Size = New System.Drawing.Size(754, 176)
'        Me.GroupBox2.TabIndex = 14
'        Me.GroupBox2.TabStop = False
'        Me.GroupBox2.Text = "Chi Tiết Các Lần Khiếu Nại Trước .."
'        '
'        'DataGridSovle
'        '
'        Me.DataGridSovle.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
'                    Or System.Windows.Forms.AnchorStyles.Left) _
'                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.DataGridSovle.ContextMenu = Me.ContextMenuGrid
'        Me.DataGridSovle.DataMember = ""
'        Me.DataGridSovle.HeaderForeColor = System.Drawing.SystemColors.ControlText
'        Me.DataGridSovle.Location = New System.Drawing.Point(7, 23)
'        Me.DataGridSovle.Name = "DataGridSovle"
'        Me.DataGridSovle.Size = New System.Drawing.Size(742, 144)
'        Me.DataGridSovle.TabIndex = 24
'        '
'        'Label2
'        '
'        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
'                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'        Me.Label2.BackColor = System.Drawing.Color.Navy
'        Me.Label2.Location = New System.Drawing.Point(9, 343)
'        Me.Label2.Name = "Label2"
'        Me.Label2.Size = New System.Drawing.Size(740, 3)
'        Me.Label2.TabIndex = 15
'        Me.Label2.Text = "Giải Quyết Khiếu Nại"
'        '
'        'frmStatistics
'        '
'        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
'        Me.ClientSize = New System.Drawing.Size(762, 566)
'        Me.Controls.Add(Me.Label2)
'        Me.Controls.Add(Me.GroupBox2)
'        Me.Controls.Add(Me.cmdclose)
'        Me.Controls.Add(Me.Label1)
'        Me.Controls.Add(Me.Label4)
'        Me.Controls.Add(Me.GroupBox1)
'        Me.Name = "frmStatistics"
'        Me.Text = "Giải Quyết Khiếu Nại Khách Hàng ..."
'        Me.GroupBox1.ResumeLayout(False)
'        CType(Me.DataGridDskh, System.ComponentModel.ISupportInitialize).EndInit()
'        Me.GroupBox2.ResumeLayout(False)
'        CType(Me.DataGridSovle, System.ComponentModel.ISupportInitialize).EndInit()
'        Me.ResumeLayout(False)

'    End Sub

'#End Region
'    Private Sub FormatGridDSKH()

'        With DataGridDskh
'            .AllowNavigation = False
'            .DataMember = "Tbl_Customers_Complain"
'            .DataSource = objdsCustomers
'            .BackgroundColor = System.Drawing.Color.LightSteelBlue
'            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
'            .CaptionForeColor = System.Drawing.Color.MediumBlue
'            .ParentRowsBackColor = System.Drawing.Color.Lavender
'            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
'            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
'            .CaptionText = "Danh Sách Khách Hàng Khiếu Nại ...."
'            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
'        End With

'        Dim TblStyle As New DataGridTableStyle
'        With TblStyle
'            .MappingName = "Tbl_Customers_Complain"
'            .BackColor = System.Drawing.Color.MintCream
'            .ForeColor = System.Drawing.Color.Navy
'            .GridLineColor = System.Drawing.Color.MediumBlue
'            .HeaderBackColor = System.Drawing.Color.Lavender
'            .HeaderForeColor = System.Drawing.Color.Navy
'            .AlternatingBackColor = Color.LightGray
'            .HeaderFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
'            .RowHeaderWidth = 10
'            .ReadOnly = False

'            With .GridColumnStyles
'                .Add(New DataGridTextBoxColumn)
'                With .Item(0)
'                    .MappingName = "ISDN"
'                    .HeaderText = "Số Điện Thoại"
'                    .Width = 130
'                    .Alignment = HorizontalAlignment.Center
'                    .NullText = String.Empty
'                End With

'                .Add(New DataGridTextBoxColumn)
'                With .Item(1)
'                    .MappingName = "Customer_Name"
'                    .HeaderText = "Tên Khách Hàng/Tên Công Ty"
'                    .Width = 300
'                    .Alignment = HorizontalAlignment.Left
'                    .NullText = String.Empty
'                End With

'                .Add(New DataGridTextBoxColumn)
'                With .Item(2)
'                    .MappingName = "ADDRESS"
'                    .HeaderText = "Địa Chỉ Khách Hàng"
'                    .Width = 300
'                    .Alignment = HorizontalAlignment.Left
'                    .NullText = String.Empty
'                End With

'            End With
'        End With
'        DataGridDskh.TableStyles.Add(TblStyle)
'    End Sub

'    Private Sub FormatGridSolve()

'        With DataGridSovle
'            .AllowNavigation = False
'            .DataMember = "Tbl_Customers_Complain.Relation_Cust_Solve"
'            .DataSource = objdsCustomers
'            .BackgroundColor = System.Drawing.Color.LightSteelBlue
'            .CaptionBackColor = System.Drawing.Color.DarkSeaGreen
'            .CaptionForeColor = System.Drawing.Color.MediumBlue
'            .ParentRowsBackColor = System.Drawing.Color.Lavender
'            .ParentRowsForeColor = System.Drawing.Color.SlateBlue
'            .CaptionFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
'            .CaptionText = "Danh Sách Khách Hàng Khiếu Nại ...."
'            .Font = New System.Drawing.Font("Times New Roman", 12.0!)
'        End With

'        Dim TblStyle As New DataGridTableStyle
'        With TblStyle
'            .MappingName = "Tbl_Solve"
'            .BackColor = System.Drawing.Color.MintCream
'            .ForeColor = System.Drawing.Color.Navy
'            .GridLineColor = System.Drawing.Color.MediumBlue
'            .HeaderBackColor = System.Drawing.Color.Lavender
'            .HeaderForeColor = System.Drawing.Color.Navy
'            .AlternatingBackColor = Color.LightGray
'            .HeaderFont = New System.Drawing.Font("Times New Roman", 12.0!, FontStyle.Bold)
'            .RowHeaderWidth = 10
'            .ReadOnly = False
'            ' Set column styles
'            With .GridColumnStyles

'                .Add(New DataGridTextBoxColumn)
'                With .Item(0)
'                    .MappingName = "ISDN"
'                    .HeaderText = "Số ĐT"
'                    .Width = 80
'                    .Alignment = HorizontalAlignment.Center
'                    .NullText = String.Empty
'                End With

'                .Add(New DataGridTextBoxColumn)
'                With .Item(1)
'                    .MappingName = "Complain_Date"
'                    .HeaderText = "Ngày Báo"
'                    .Width = 100
'                    .Alignment = HorizontalAlignment.Left
'                    .NullText = String.Empty
'                End With

'                .Add(New DataGridTextBoxColumn)
'                With .Item(2)
'                    .MappingName = "Content_ComPlain"
'                    .HeaderText = "Nội Dung Khiếu Nại"
'                    .Width = 300
'                    .Alignment = HorizontalAlignment.Left
'                    .NullText = String.Empty
'                End With

'                .Add(New DataGridTextBoxColumn)
'                With .Item(3)
'                    .MappingName = "Explain_Date"
'                    .HeaderText = "Ngày Giải Quyết"
'                    .Width = 100
'                    .Alignment = HorizontalAlignment.Left
'                    .NullText = String.Empty
'                End With

'                .Add(New DataGridTextBoxColumn)
'                With .Item(4)
'                    .MappingName = "Content_Explain"
'                    .HeaderText = "Nội Dung Trả Lời"
'                    .Width = 300
'                    .Alignment = HorizontalAlignment.Left
'                    .NullText = String.Empty
'                End With

'            End With
'        End With
'        Dim ComboTextCol As New DataGridComboBoxColumn
'        ComboTextCol.MappingName = "TypeCom"
'        ComboTextCol.HeaderText = "Loại Khiếu Nại"
'        ComboTextCol.Width = 120
'        ComboTextCol.ColumnComboBox.DataSource = mydataset.Tables("Tbl_Type_Complain").DefaultView
'        ComboTextCol.ColumnComboBox.DisplayMember = "TypeName"
'        ComboTextCol.ColumnComboBox.ValueMember = "TypeID"
'        TblStyle.GridColumnStyles.Add(ComboTextCol)

'        Dim ComboTextCol1 As New DataGridComboBoxColumn
'        ComboTextCol1.MappingName = "States"
'        ComboTextCol1.HeaderText = "Tình Trạng"
'        ComboTextCol1.Width = 120
'        ComboTextCol1.ColumnComboBox.DataSource = mydataset.Tables("StaticsType").DefaultView
'        ComboTextCol1.ColumnComboBox.DisplayMember = "TypeName"
'        ComboTextCol1.ColumnComboBox.ValueMember = "TypeID"
'        TblStyle.PreferredRowHeight = ComboTextCol1.ColumnComboBox.Height + 10
'        TblStyle.GridColumnStyles.Add(ComboTextCol1)
'        DataGridSovle.TableStyles.Add(TblStyle)
'    End Sub

'    Private Sub frmStatistics_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        Try
'            mydataset = New DataSet
'            'FillDataset()
'            InitComponents()
'            LoadDataSet()
'            FormatGridDSKH()
'            FormatGridSolve()
'        Catch eLoad As System.Exception
'            System.Windows.Forms.MessageBox.Show(eLoad.Message)
'        End Try
'    End Sub

'    Private Sub cmdclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdclose.Click
'        Me.Close()
'    End Sub

'    Private Sub FillDataset()
'        strSQL = "SELECT * FROM Tbl_Customers_Complain "

'        Try
'            Dim cmd As New OleDbCommand(strSQL, oledbcon)
'            da = New OleDbDataAdapter(cmd)
'            da.Fill(mydataset, "Customers")
'        Catch ex As Exception
'            MsgBox(ex.ToString)
'        End Try

'        strSQL = "SELECT * FROM Tbl_Solve "
'        Try
'            Dim cmd As New OleDbCommand(strSQL, oledbcon)
'            da = New OleDbDataAdapter(cmd)
'            da.Fill(mydataset, "Solve")
'        Catch ex As Exception
'            MsgBox(ex.ToString)
'        End Try

'        strSQL = "SELECT TypeName,TypeID FROM Tbl_TypeStatics "
'        Try
'            Dim cmd As New OleDbCommand(strSQL, oledbcon)
'            da = New OleDbDataAdapter(cmd)
'            da.Fill(mydataset, "StaticsType")
'        Catch ex As Exception
'            MsgBox(ex.ToString)
'        End Try

'        'Dim myDataRelation As DataRelation
'        'myDataRelation = New DataRelation("Cust_Solve", mydataset.Tables("Customers").Columns("ISDN"), mydataset.Tables("Solve").Columns("ISDN"))
'        'mydataset.Relations.Add(myDataRelation)
'        'DataGridDskh.SetDataBinding(mydataset, "Customers")
'        'DataGridDskh.show
'        'DataGrid1.SetDataBinding(mydataset, "Customers.Cust_Solve")
'        'DataGridDskh.DataSource = dsoject.Tbl_Customers_Complain
'    End Sub

'    Public Sub cmdUpdate()
'        Try
'            Me.UpdateDataSet()
'        Catch eUpdate As System.Exception
'            System.Windows.Forms.MessageBox.Show(eUpdate.Message)
'        End Try
'    End Sub

'    Private Sub InitComponents()
'        Me.OleDbSelectCust = New System.Data.OleDb.OleDbCommand
'        Me.OleDbInsertCust = New System.Data.OleDb.OleDbCommand
'        Me.OleDbUpdateCust = New System.Data.OleDb.OleDbCommand
'        Me.OleDbDeleteCust = New System.Data.OleDb.OleDbCommand
'        Me.OleDbSelectSolve = New System.Data.OleDb.OleDbCommand
'        Me.OleDbInsertSolve = New System.Data.OleDb.OleDbCommand
'        Me.OleDbUpdateSolve = New System.Data.OleDb.OleDbCommand
'        Me.OleDbDeleteSolve = New System.Data.OleDb.OleDbCommand
'        Me.OleDbDataAdapterCust = New System.Data.OleDb.OleDbDataAdapter
'        Me.OleDbDataAdapterSolve = New System.Data.OleDb.OleDbDataAdapter
'        Me.objdsCustomers = New dsCustomers

'        'oledbselectCust
'        '
'        Me.OleDbSelectCust.CommandText = "SELECT ISDN, Address, Customer_Name, Address  FROM Tbl_Customers_Complain"
'        Me.OleDbSelectCust.Connection = oledbcon
'        '
'        'OleDbInsertCust
'        '
'        Me.OleDbInsertCust.CommandText = "INSERT INTO Tbl_Customers_Complain(ISDN, Customer_Name,Address ) VALUES (?, ?, ?)"
'        Me.OleDbInsertCust.Connection = oledbcon
'        Me.OleDbInsertCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, "ISDN"))
'        Me.OleDbInsertCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Customer_Name", System.Data.OleDb.OleDbType.VarWChar, 255, "Customer_Name"))
'        Me.OleDbInsertCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Address", System.Data.OleDb.OleDbType.VarWChar, 255, "Address"))
'        '
'        'OleDbUpdateCust
'        '
'        Me.OleDbUpdateCust.CommandText = "UPDATE Tbl_Customers_Complain SET ISDN = ?, Customer_Name = ?,Address  = ? WHERE " & _
'        "(ISDN = ?) AND (Address = ? OR ? IS NULL AND Address IS NULL) AND (Customer_Name" & _
'        " = ? OR ? IS NULL AND Customer_Name IS NULL)"
'        Me.OleDbUpdateCust.Connection = oledbcon
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, "ISDN"))
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Customer_Name", System.Data.OleDb.OleDbType.VarWChar, 255, "Customer_Name"))
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Address", System.Data.OleDb.OleDbType.VarWChar, 255, "Address"))
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ISDN", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Address", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Address", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Address1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Address", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Customer_Name", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Customer_Name", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbUpdateCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Customer_Name1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Customer_Name", System.Data.DataRowVersion.Original, Nothing))
'        '
'        'OleDbDeleteCust
'        '
'        Me.OleDbDeleteCust.CommandText = "DELETE FROM Tbl_Customers_Complain WHERE (ISDN = ?) AND (Address = ? OR ? IS NULL" & _
'        " AND Address IS NULL) AND (Customer_Name = ? OR ? IS NULL AND Customer_Name IS N" & _
'        "ULL)"
'        Me.OleDbDeleteCust.Connection = oledbcon
'        Me.OleDbDeleteCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ISDN", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbDeleteCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Address", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Address", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbDeleteCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Address1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Address", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbDeleteCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Customer_Name", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Customer_Name", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbDeleteCust.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Customer_Name1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Customer_Name", System.Data.DataRowVersion.Original, Nothing))
'        '
'        'OleDbSelectSolve
'        '
'        Me.OleDbSelectSolve.CommandText = "SELECT ID, ISDN, Complain_Date, Content_ComPlain, Explain_Date,Content_Explain,TypeCom,  States FROM Tbl_Solve"
'        Me.OleDbSelectSolve.Connection = oledbcon
'        '
'        'OleDbInsertSolve
'        '
'        Me.OleDbInsertSolve.CommandText = "INSERT INTO Tbl_Solve(ISDN,Complain_Date, Content_ComPlain,Explain_Date,  Content_Explain, TypeCom, States) VALUES (?, ?, ?, ?, ?, ?,?)"
'        Me.OleDbInsertSolve.Connection = oledbcon
'        Me.OleDbInsertSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, "ISDN"))
'        Me.OleDbInsertSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Complain_Date", System.Data.OleDb.OleDbType.VarWChar, 10, "Complain_Date"))
'        Me.OleDbInsertSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Content_ComPlain", System.Data.OleDb.OleDbType.VarWChar, 255, "Content_ComPlain"))
'        Me.OleDbInsertSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Explain_Date", System.Data.OleDb.OleDbType.VarWChar, 10, "Explain_Date"))
'        Me.OleDbInsertSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Content_Explain", System.Data.OleDb.OleDbType.VarWChar, 255, "Content_Explain"))
'        Me.OleDbInsertSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCom", System.Data.OleDb.OleDbType.VarWChar, 50, "TypeCom"))
'        Me.OleDbInsertSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("States", System.Data.OleDb.OleDbType.VarWChar, 50, "States"))
'        '
'        'OleDbUpdateSolve
'        '
'        Me.OleDbUpdateSolve.CommandText = "UPDATE Tbl_Solve SET  ISDN = ?, Complain_Date = ?, Content_ComPlain = ?, Explain_Date = ?, Content_Explain = ?" & _
'        " , TypeCom =? , States = ? WHERE (ID = ?) AND (ISDN = ?)"
'        Me.OleDbUpdateSolve.Connection = oledbcon
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, "ISDN"))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Complain_Date", System.Data.OleDb.OleDbType.VarWChar, 10, "Complain_Date"))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Content_ComPlain", System.Data.OleDb.OleDbType.VarWChar, 255, "Content_ComPlain"))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Explain_Date", System.Data.OleDb.OleDbType.VarWChar, 10, "Explain_Date"))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Content_Explain", System.Data.OleDb.OleDbType.VarWChar, 255, "Content_Explain"))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCom", System.Data.OleDb.OleDbType.VarWChar, 50, "TypeCom"))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("States", System.Data.OleDb.OleDbType.VarWChar, 50, "States"))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ID", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbUpdateSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ISDN", System.Data.DataRowVersion.Original, Nothing))

'        'OleDbDeleteSolve
'        '
'        Me.OleDbDeleteSolve.CommandText = "DELETE FROM Tbl_Solve WHERE (ID = ?) AND (ISDN = ?  )"
'        Me.OleDbDeleteSolve.Connection = oledbcon
'        Me.OleDbDeleteSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ID", System.Data.DataRowVersion.Original, Nothing))
'        Me.OleDbDeleteSolve.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ISDN", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ISDN", System.Data.DataRowVersion.Original, Nothing))

'        '
'        'OleDbDataAdapter1
'        '
'        Me.OleDbDataAdapterCust.DeleteCommand = Me.OleDbDeleteCust
'        Me.OleDbDataAdapterCust.InsertCommand = Me.OleDbInsertCust
'        Me.OleDbDataAdapterCust.SelectCommand = Me.OleDbSelectCust
'        Me.OleDbDataAdapterCust.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Tbl_Customers_Complain", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ISDN", "ISDN"), New System.Data.Common.DataColumnMapping("Customer_Name", "Customer_Name"), New System.Data.Common.DataColumnMapping("Address", "Address")})})
'        Me.OleDbDataAdapterCust.UpdateCommand = Me.OleDbUpdateCust
'        '
'        'OleDbDataAdapter2
'        '
'        Me.OleDbDataAdapterSolve.DeleteCommand = Me.OleDbDeleteSolve
'        Me.OleDbDataAdapterSolve.InsertCommand = Me.OleDbInsertSolve
'        Me.OleDbDataAdapterSolve.SelectCommand = Me.OleDbSelectSolve
'        Me.OleDbDataAdapterSolve.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Tbl_Solve", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ID", "ID"), New System.Data.Common.DataColumnMapping("ISDN", "ISDN"), New System.Data.Common.DataColumnMapping("Complain_Date", "Complain_Date"), New System.Data.Common.DataColumnMapping("Content_ComPlain", "Content_ComPlain"), New System.Data.Common.DataColumnMapping("Explain_Date", "Explain_Date"), New System.Data.Common.DataColumnMapping("Content_Explain", "Content_Explain"), New System.Data.Common.DataColumnMapping("TypeCom", "TypeCom"), New System.Data.Common.DataColumnMapping("States", "States")})})
'        Me.OleDbDataAdapterSolve.UpdateCommand = Me.OleDbUpdateSolve

'        'objdsCustomers
'        Me.objdsCustomers.DataSetName = "dsCustomers"
'        Me.objdsCustomers.Locale = New System.Globalization.CultureInfo("en-US")

'    End Sub
'    'Public Sub LoadDataSet()

'    '    strSQL = "SELECT TypeName,TypeID FROM Tbl_Type_Complain "
'    '    Try
'    '        Dim cmd As New OleDbCommand(strSQL, oledbcon)
'    '        da = New OleDbDataAdapter(cmd)
'    '        da.Fill(mydataset, "Tbl_Type_Complain")
'    '    Catch ex As Exception
'    '        MsgBox(ex.ToString)
'    '    End Try

'    '    strSQL = "SELECT TypeName,TypeID FROM Tbl_TypeStatics "
'    '    Try
'    '        Dim cmd As New OleDbCommand(strSQL, oledbcon)
'    '        da = New OleDbDataAdapter(cmd)
'    '        da.Fill(mydataset, "StaticsType")
'    '    Catch ex As Exception
'    '        MsgBox(ex.ToString)
'    '    End Try

'    '    Dim objDataSetTemp As New dsCustomers
'    '    Try
'    '        Me.FillDataset(objDataSetTemp)
'    '    Catch eFillDataSet As System.Exception
'    '        Throw eFillDataSet
'    '    End Try
'    '    Try
'    '        DataGridDskh.DataSource = Nothing
'    '        DataGridSovle.DataSource = Nothing
'    '        objdsCustomers.Clear()
'    '        objdsCustomers.Merge(objDataSetTemp)
'    '        DataGridDskh.SetDataBinding(objdsCustomers, "Tbl_Customers_Complain")
'    '        DataGridSovle.SetDataBinding(objdsCustomers, "Tbl_Customers_Complain.Relation_Cust_Solve")
'    '    Catch eLoadMerge As System.Exception
'    '        Throw eLoadMerge
'    '    End Try
'    'End Sub

'    'Public Sub FillDataSet(ByVal dS As dsCustomers)
'    '    dS.EnforceConstraints = False
'    '    Try
'    '        oledbcon.Open()
'    '        Me.OleDbDataAdapterCust.Fill(dS)
'    '        Me.OleDbDataAdapterSolve.Fill(dS)
'    '    Catch fillException As System.Exception
'    '        Throw fillException
'    '    Finally
'    '        dS.EnforceConstraints = True
'    '        oledbcon.Close()
'    '    End Try
'    'End Sub

'    'Public Sub UpdateDataSet()
'    '    Dim objDataSetChanges As dsCustomers = New dsCustomers
'    '    Me.BindingContext(objdsCustomers, "Tbl_Customers_Complain").EndCurrentEdit()
'    '    Me.BindingContext(objdsCustomers, "Tbl_Solve").EndCurrentEdit()
'    '    objDataSetChanges = CType(objdsCustomers.GetChanges, dsCustomers)
'    '    If (Not (objDataSetChanges) Is Nothing) Then
'    '        Try
'    '            Me.UpdateDataSource(objDataSetChanges)
'    '            Me.objdsCustomers.Merge(objDataSetChanges)
'    '            Me.objdsCustomers.AcceptChanges()
'    '        Catch eUpdate As System.Exception
'    '            Throw eUpdate
'    '        End Try
'    '    End If

'    'End Sub
'    'Public Sub UpdateDataSource(ByVal ChangedRows As dsCustomers)
'    '    Try
'    '        If (Not (ChangedRows) Is Nothing) Then
'    '            oledbcon.Open()
'    '            Me.OleDbDataAdapterCust.Update(ChangedRows)
'    '            Me.OleDbDataAdapterSolve.Update(ChangedRows)
'    '        End If
'    '    Catch updateException As System.Exception
'    '        Throw updateException
'    '    Finally
'    '        oledbcon.Close()
'    '    End Try
'    'End Sub

'    'Private Sub MenuItemTim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemTim.Click
'    '    Dim ISDN As String
'    '    Dim i As Integer
'    '    ISDN = InputBox("Nhập Vào Số Cần Tìm ...", "Tìm ...")
'    '    For i = 0 To DataGridDskh.VisibleRowCount - 2
'    '        With DataGridDskh
'    '            If (Trim$(.Item(i, 0)) = ISDN) Then
'    '                .Select(i)
'    '                Exit Sub
'    '            End If
'    '        End With
'    '    Next
'    'End Sub
'End Class
