Imports System.Data.OleDb
Public Class frmImport
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFolder As System.Windows.Forms.TextBox
    Friend WithEvents btnChoose As System.Windows.Forms.Button
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents optSoDauKy As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmImport))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.optSoDauKy = New System.Windows.Forms.RadioButton
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnImport = New System.Windows.Forms.Button
        Me.btnChoose = New System.Windows.Forms.Button
        Me.txtFolder = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.btnImport)
        Me.GroupBox1.Controls.Add(Me.btnChoose)
        Me.GroupBox1.Controls.Add(Me.txtFolder)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(356, 96)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton1)
        Me.GroupBox2.Controls.Add(Me.optSoDauKy)
        Me.GroupBox2.Location = New System.Drawing.Point(64, 32)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(256, 32)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        '
        'RadioButton1
        '
        Me.RadioButton1.Location = New System.Drawing.Point(136, 12)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(112, 16)
        Me.RadioButton1.TabIndex = 7
        Me.RadioButton1.Text = "Số Liệu Báo Cáo"
        '
        'optSoDauKy
        '
        Me.optSoDauKy.Location = New System.Drawing.Point(8, 12)
        Me.optSoDauKy.Name = "optSoDauKy"
        Me.optSoDauKy.Size = New System.Drawing.Size(104, 16)
        Me.optSoDauKy.TabIndex = 6
        Me.optSoDauKy.Text = "Số Giao Đầu Kỳ"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(200, 66)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 24)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(88, 67)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(88, 24)
        Me.btnImport.TabIndex = 3
        Me.btnImport.Text = "Import"
        '
        'btnChoose
        '
        Me.btnChoose.Location = New System.Drawing.Point(320, 10)
        Me.btnChoose.Name = "btnChoose"
        Me.btnChoose.Size = New System.Drawing.Size(32, 19)
        Me.btnChoose.TabIndex = 2
        Me.btnChoose.Text = "..."
        '
        'txtFolder
        '
        Me.txtFolder.Location = New System.Drawing.Point(64, 10)
        Me.txtFolder.Name = "txtFolder"
        Me.txtFolder.Size = New System.Drawing.Size(256, 20)
        Me.txtFolder.TabIndex = 1
        Me.txtFolder.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "File Folder:"
        '
        'frmImport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(365, 99)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmImport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import WorkSheet Into Database"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim ds As New dsTableImport
    Private Sub UpdateDB(ByVal ds As DataSet, ByVal tblName As String)
        Try
            oledbcon.Open()
            Dim olecmd As OleDbCommand
            Dim i As Integer
            Dim strSQL As String
            For i = 0 To ds.Tables(0).Rows.Count - 1
                strSQL = " INSERT INTO " & tblName & "(STT,TenKh,MaKH,SoTB,Item_No,DiaChi,NoTruoc,DieuChinh,PhatSinh,Thue,TongCuoc,MaTram,MaNV) " & _
                " VALUES('" & ds.Tables(0).Rows(i)("STT") & _
                "','" & ds.Tables(0).Rows(i)("TenKH") & _
                "','" & ds.Tables(0).Rows(i)("MaKH") & _
                "','" & ds.Tables(0).Rows(i)("SoTB") & _
                "','" & ds.Tables(0).Rows(i)("Item_No") & _
                "','" & ds.Tables(0).Rows(i)("DiaChi") & _
                "'," & CDbl(ds.Tables(0).Rows(i)("NoTruoc")) & _
                "," & CDbl(ds.Tables(0).Rows(i)("DieuChinh")) & _
                "," & CDbl(ds.Tables(0).Rows(i)("PhatSinh")) & _
                "," & CDbl(ds.Tables(0).Rows(i)("Thue")) & _
                "," & CDbl(ds.Tables(0).Rows(i)("TongCuoc")) & _
                ",'" & ds.Tables(0).Rows(i)("MaTram") & _
                "','" & ds.Tables(0).Rows(i)("MaNV") & "')"
                olecmd = New OleDbCommand(strSQL, oledbcon)
                olecmd.ExecuteNonQuery()
            Next
            
            MsgBox("Đã import xong!!!")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        oledbcon.Close()
    End Sub
    Private Sub UpdateSoLieuCacTram(ByVal ds As DataSet, ByVal tblName As String)
        Try
            oledbcon.Open()
            Dim olecmd As OleDbCommand
            Dim i As Integer
            Dim strSQL As String
            For i = 0 To ds.Tables("SoLieuBaoCao").Rows.Count - 1
                strSQL = " INSERT INTO " & tblName & "(STT,SoHD,MaKH,SoTB,TenKH,TuNgay,DenNgay,SoTienTra,NgayTra,MaTram,MaNV) " & _
                " VALUES('" & ds.Tables("SoLieuBaoCao").Rows(i)("STT") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("SoHD") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("MaKH") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("SoTB") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("TenKH") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("TuNgay") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("DenNgay") & _
                "'," & ds.Tables("SoLieuBaoCao").Rows(i)("SoTienTra") & _
                ",'" & ds.Tables("SoLieuBaoCao").Rows(i)("NgayTra") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("MaTram") & _
                "','" & ds.Tables("SoLieuBaoCao").Rows(i)("MaNV") & "')"
                olecmd = New OleDbCommand(strSQL, oledbcon)
                olecmd.ExecuteNonQuery()
            Next

            MsgBox("Đã import xong!!!")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        oledbcon.Close()
    End Sub
    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        'OpenConnect()
        ImportExcel(txtFolder.Text)
    End Sub

    Private Sub btnChoose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChoose.Click
        Dim varfolderBrowserDialog As New OpenFileDialog
        Dim result As DialogResult = varfolderBrowserDialog.ShowDialog
        If result = DialogResult.OK Then
            txtFolder.Text = varfolderBrowserDialog.FileName
        End If
    End Sub

    Private Sub ImportExcel(ByVal strPath As String)
        Dim excelApp As New Excel.Application
        Dim excelBook As Excel.Workbook = excelApp.Workbooks.Open(strPath)
        Dim excelWorkSheet As Excel.Worksheet = CType(excelBook.Worksheets("Sheet1"), Excel.Worksheet)
        Dim myRow As DataRow
        Dim irows, count As Integer
        Dim olecmd As OleDbCommand

        Dim sqlInsert As String
        Dim varA, varB, varC, varD, varE, varF, varG, varH, varI, varJ, varK, varL, varMaTram, varMaNV As String

        'OpenConnect()

        Try
            If optSoDauKy.Checked = True Then
                With excelWorkSheet
                    count = 0
                    For irows = 8 To .Rows.Count
                        varA = .Range("A" & irows.ToString).Value
                        If varA = Nothing Then
                            Exit For
                        End If
                        If Mid(varA, 1, 9) = "   Tổ thu" Then
                            varMaTram = Mid(.Range("A" & irows.ToString).Value, 34, (Len(.Range("A" & irows.ToString).Value) - 34) + 1)
                        End If
                        If Mid(varA, 1, 15) = "      Nhân viên" Then
                            varMaNV = Trim(Mid(.Range("A" & irows.ToString).Value, 17, 17))
                        End If
                        'Mid(varA, 1, 9) <> "   Tổ thu" And Mid(varA, 1, 15) <> "      Nhân viên" And Mid(varA, 1, 16) <> "         HT giao" 
                        If IsNumeric(varA) Then
                            varB = .Range("B" & irows.ToString).Value
                            varC = .Range("C" & irows.ToString).Value
                            varD = .Range("D" & irows.ToString).Value
                            'varE = .Range("E" & irows.ToString).Value
                            varF = .Range("F" & irows.ToString).Value
                            varG = .Range("G" & irows.ToString).Value
                            varH = .Range("H" & irows.ToString).Value
                            varI = .Range("I" & irows.ToString).Value
                            varJ = .Range("J" & irows.ToString).Value
                            varK = .Range("K" & irows.ToString).Value
                            varL = .Range("L" & irows.ToString).Value

                            myRow = ds.Tables("MyTable").NewRow
                            ds.Tables("MyTable").Rows.Add(myRow)
                            ds.Tables("MyTable").Rows(count).Item("STT") = varA
                            ds.Tables("MyTable").Rows(count).Item("TenKH") = varB
                            ds.Tables("MyTable").Rows(count).Item("MaKH") = varC
                            ds.Tables("MyTable").Rows(count).Item("SoTB") = varD
                            ds.Tables("MyTable").Rows(count).Item("Item_No") = varF
                            ds.Tables("MyTable").Rows(count).Item("DiaChi") = varG
                            ds.Tables("MyTable").Rows(count).Item("NoTruoc") = CDbl(varH)
                            ds.Tables("MyTable").Rows(count).Item("DieuChinh") = CDbl(varI)
                            ds.Tables("MyTable").Rows(count).Item("PhatSinh") = CDbl(varJ)
                            ds.Tables("MyTable").Rows(count).Item("Thue") = CDbl(varK)
                            ds.Tables("MyTable").Rows(count).Item("TongCuoc") = CDbl(varL)
                            ds.Tables("MyTable").Rows(count).Item("MaTram") = varMaTram
                            ds.Tables("MyTable").Rows(count).Item("MaNV") = varMaNV
                            ''''''''''''''''
                            count = count + 1
                        End If
                    Next
                End With
                UpdateDB(ds, "Tbl_SoGiaoDauKy")
            Else
                With excelWorkSheet
                    count = 0
                    For irows = 8 To .Rows.Count
                        varA = .Range("A" & irows.ToString).Value
                        If varA = Nothing Then
                            Exit For
                        End If
                        If Mid(varA, 1, 9) = "   Tổ thu" Then
                            varMaTram = Mid(.Range("A" & irows.ToString).Value, 34, (Len(.Range("A" & irows.ToString).Value) - 33) + 1)
                        End If
                        If Mid(varA, 1, 16) = "    CTV thu cước" Then
                            varMaNV = Mid(.Range("A" & irows.ToString).Value, 18, 17)
                            irows = irows + 3
                        End If
                        'Mid(varA, 1, 9) <> "   Tổ thu" And Mid(varA, 1, 16) <> "    CTV thu cước"
                        If IsNumeric(varA) Then
                            varB = .Range("B" & irows.ToString).Value
                            varC = .Range("C" & irows.ToString).Value 'Ma KH
                            varD = .Range("D" & irows.ToString).Value 'So TB
                            varE = .Range("E" & irows.ToString).Value 'TenKh
                            varF = .Range("F" & irows.ToString).Value 'Tu ngay
                            varG = .Range("G" & irows.ToString).Value 'Den Ngay
                            varH = .Range("H" & irows.ToString).Value 'So Tien tra
                            varI = .Range("I" & irows.ToString).Value 'Ngay Tra

                            myRow = ds.Tables("SoLieuBaoCao").NewRow
                            ds.Tables("SoLieuBaoCao").Rows.Add(myRow)
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("STT") = varA
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("SoHD") = varB
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("MaKH") = varC
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("SoTB") = varD
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("TenKH") = varE
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("TuNgay") = varF
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("DenNgay") = varG
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("SoTienTra") = CDbl(varH)
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("NgayTra") = varI
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("MaTram") = varMaTram
                            ds.Tables("SoLieuBaoCao").Rows(count).Item("MaNV") = Trim(varMaNV)
                            ''''''''''''''''
                            count = count + 1
                        End If
                    Next
                End With
                UpdateSoLieuCacTram(ds, "Tbl_SoLieuBaoCao")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        excelBook.Close()
        excelApp.Quit()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmImport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        optSoDauKy.Checked = True
    End Sub
End Class
