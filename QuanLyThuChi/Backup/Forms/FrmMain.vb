Imports System.IO
Public Class FrmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        strPrinterName = GetSetting("QuanLyCTV", "QuanLyCTV", "PrinterName")
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
    Friend WithEvents FolderBrowser As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents StatusBarMain As System.Windows.Forms.StatusBar
    Friend WithEvents ToolBarMain As System.Windows.Forms.ToolBar
    Friend WithEvents IconList As System.Windows.Forms.ImageList
    Friend WithEvents ToolBarcmdExit As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarcmdFind As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarcmdcacl As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarreport As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarcmdhelp As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarcmdWindowsVersion As System.Windows.Forms.ToolBarButton
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents TimerMain As System.Windows.Forms.Timer
    Friend WithEvents MainMN As System.Windows.Forms.MainMenu
    Friend WithEvents ToolBarcmdNewEmployee As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtoncmdPhieuthu As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarCmdPhieuChi As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarcmdListReceipts As System.Windows.Forms.ToolBarButton
    Friend WithEvents ContextMenuReports As System.Windows.Forms.ContextMenu
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents ToolBarButtoncmddaily As System.Windows.Forms.ToolBarButton
    Friend WithEvents MnuBaocaoQuy As System.Windows.Forms.MenuItem
    Friend WithEvents MnuBaocaoThuChi As System.Windows.Forms.MenuItem
    Friend WithEvents MnuBaoCaoChiNop As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTonhHopGNT As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemBaoCaoNgay As System.Windows.Forms.MenuItem
    Friend WithEvents MnuBaocaoTongHopThu As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarCmdLogin As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarCmdLogOut As System.Windows.Forms.ToolBarButton
    Friend WithEvents StatusBarPanel4 As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmMain))
        Me.FolderBrowser = New System.Windows.Forms.FolderBrowserDialog
        Me.StatusBarMain = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.ToolBarMain = New System.Windows.Forms.ToolBar
        Me.ToolBarcmdExit = New System.Windows.Forms.ToolBarButton
        Me.ToolBarCmdLogin = New System.Windows.Forms.ToolBarButton
        Me.ToolBarCmdLogOut = New System.Windows.Forms.ToolBarButton
        Me.ToolBarcmdNewEmployee = New System.Windows.Forms.ToolBarButton
        Me.ToolBarcmdFind = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtoncmdPhieuthu = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtoncmddaily = New System.Windows.Forms.ToolBarButton
        Me.ToolBarCmdPhieuChi = New System.Windows.Forms.ToolBarButton
        Me.ToolBarcmdListReceipts = New System.Windows.Forms.ToolBarButton
        Me.ToolBarreport = New System.Windows.Forms.ToolBarButton
        Me.ContextMenuReports = New System.Windows.Forms.ContextMenu
        Me.MnuBaocaoQuy = New System.Windows.Forms.MenuItem
        Me.MnuBaocaoThuChi = New System.Windows.Forms.MenuItem
        Me.MnuBaocaoTongHopThu = New System.Windows.Forms.MenuItem
        Me.MenuItemBaoCaoNgay = New System.Windows.Forms.MenuItem
        Me.MnuBaoCaoChiNop = New System.Windows.Forms.MenuItem
        Me.MnuTonhHopGNT = New System.Windows.Forms.MenuItem
        Me.ToolBarcmdcacl = New System.Windows.Forms.ToolBarButton
        Me.ToolBarcmdWindowsVersion = New System.Windows.Forms.ToolBarButton
        Me.ToolBarcmdhelp = New System.Windows.Forms.ToolBarButton
        Me.IconList = New System.Windows.Forms.ImageList(Me.components)
        Me.TimerMain = New System.Windows.Forms.Timer(Me.components)
        Me.MainMN = New System.Windows.Forms.MainMenu
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.StatusBarPanel4 = New System.Windows.Forms.StatusBarPanel
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusBarMain
        '
        Me.StatusBarMain.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StatusBarMain.Dock = System.Windows.Forms.DockStyle.None
        Me.StatusBarMain.Location = New System.Drawing.Point(0, 640)
        Me.StatusBarMain.Name = "StatusBarMain"
        Me.StatusBarMain.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel4, Me.StatusBarPanel2, Me.StatusBarPanel3})
        Me.StatusBarMain.ShowPanels = True
        Me.StatusBarMain.Size = New System.Drawing.Size(864, 22)
        Me.StatusBarMain.TabIndex = 1
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.None
        Me.StatusBarPanel1.Text = "Chương Trình Quản Lý CTV "
        Me.StatusBarPanel1.Width = 250
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.Text = "OCX"
        Me.StatusBarPanel2.Width = 300
        '
        'StatusBarPanel3
        '
        Me.StatusBarPanel3.Width = 200
        '
        'ToolBarMain
        '
        Me.ToolBarMain.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarcmdExit, Me.ToolBarCmdLogin, Me.ToolBarCmdLogOut, Me.ToolBarcmdNewEmployee, Me.ToolBarcmdFind, Me.ToolBarButtoncmdPhieuthu, Me.ToolBarButtoncmddaily, Me.ToolBarCmdPhieuChi, Me.ToolBarcmdListReceipts, Me.ToolBarreport, Me.ToolBarcmdcacl, Me.ToolBarcmdWindowsVersion, Me.ToolBarcmdhelp})
        Me.ToolBarMain.ButtonSize = New System.Drawing.Size(23, 22)
        Me.ToolBarMain.DropDownArrows = True
        Me.ToolBarMain.ImageList = Me.IconList
        Me.ToolBarMain.Location = New System.Drawing.Point(0, 0)
        Me.ToolBarMain.Name = "ToolBarMain"
        Me.ToolBarMain.ShowToolTips = True
        Me.ToolBarMain.Size = New System.Drawing.Size(864, 28)
        Me.ToolBarMain.TabIndex = 2
        '
        'ToolBarcmdExit
        '
        Me.ToolBarcmdExit.ImageIndex = 0
        Me.ToolBarcmdExit.ToolTipText = "Thoát Khỏi Ứng Dụng"
        '
        'ToolBarCmdLogin
        '
        Me.ToolBarCmdLogin.ImageIndex = 23
        Me.ToolBarCmdLogin.ToolTipText = "Đăng Nhập Hệ Thống"
        '
        'ToolBarCmdLogOut
        '
        Me.ToolBarCmdLogOut.ImageIndex = 24
        Me.ToolBarCmdLogOut.ToolTipText = "Đăng Xuất Hệ Thống"
        '
        'ToolBarcmdNewEmployee
        '
        Me.ToolBarcmdNewEmployee.ImageIndex = 11
        Me.ToolBarcmdNewEmployee.ToolTipText = "Nhập Mới Cộng Tác Viên"
        '
        'ToolBarcmdFind
        '
        Me.ToolBarcmdFind.ImageIndex = 1
        Me.ToolBarcmdFind.ToolTipText = "Danh Sách Đại Lý PSTN"
        '
        'ToolBarButtoncmdPhieuthu
        '
        Me.ToolBarButtoncmdPhieuthu.ImageIndex = 2
        Me.ToolBarButtoncmdPhieuthu.ToolTipText = "Nhập Phiếu Thu"
        '
        'ToolBarButtoncmddaily
        '
        Me.ToolBarButtoncmddaily.ImageIndex = 17
        Me.ToolBarButtoncmddaily.ToolTipText = "Thanh toán đại lý"
        '
        'ToolBarCmdPhieuChi
        '
        Me.ToolBarCmdPhieuChi.Enabled = False
        Me.ToolBarCmdPhieuChi.ImageIndex = 16
        Me.ToolBarCmdPhieuChi.ToolTipText = "Nhập phiếu chi"
        '
        'ToolBarcmdListReceipts
        '
        Me.ToolBarcmdListReceipts.ImageIndex = 13
        Me.ToolBarcmdListReceipts.ToolTipText = "Lập bảng kê nộp tiền"
        '
        'ToolBarreport
        '
        Me.ToolBarreport.DropDownMenu = Me.ContextMenuReports
        Me.ToolBarreport.ImageIndex = 4
        Me.ToolBarreport.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.ToolBarreport.ToolTipText = "Báo Cáo Quỹ"
        '
        'ContextMenuReports
        '
        Me.ContextMenuReports.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuBaocaoQuy, Me.MnuBaocaoThuChi, Me.MnuBaocaoTongHopThu, Me.MenuItemBaoCaoNgay, Me.MnuBaoCaoChiNop, Me.MnuTonhHopGNT})
        '
        'MnuBaocaoQuy
        '
        Me.MnuBaocaoQuy.Index = 0
        Me.MnuBaocaoQuy.Text = "Báo Cáo Quỹ"
        '
        'MnuBaocaoThuChi
        '
        Me.MnuBaocaoThuChi.Index = 1
        Me.MnuBaocaoThuChi.Text = "Báo Cáo Thu - Chi"
        '
        'MnuBaocaoTongHopThu
        '
        Me.MnuBaocaoTongHopThu.Index = 2
        Me.MnuBaocaoTongHopThu.Text = "Báo CáoTổng Hợp Thu"
        '
        'MenuItemBaoCaoNgay
        '
        Me.MenuItemBaoCaoNgay.Index = 3
        Me.MenuItemBaoCaoNgay.Text = "Báo Cáo Ngày - Tháng"
        '
        'MnuBaoCaoChiNop
        '
        Me.MnuBaoCaoChiNop.Index = 4
        Me.MnuBaoCaoChiNop.Text = "Báo Cáo Chi - Nộp Tiền"
        '
        'MnuTonhHopGNT
        '
        Me.MnuTonhHopGNT.Index = 5
        Me.MnuTonhHopGNT.Text = "Tổng Hợp GNT NH"
        '
        'ToolBarcmdcacl
        '
        Me.ToolBarcmdcacl.ImageIndex = 5
        Me.ToolBarcmdcacl.ToolTipText = "Calculator"
        '
        'ToolBarcmdWindowsVersion
        '
        Me.ToolBarcmdWindowsVersion.ImageIndex = 7
        Me.ToolBarcmdWindowsVersion.ToolTipText = "Windows Version"
        '
        'ToolBarcmdhelp
        '
        Me.ToolBarcmdhelp.ImageIndex = 8
        Me.ToolBarcmdhelp.ToolTipText = "Help"
        '
        'IconList
        '
        Me.IconList.ImageSize = New System.Drawing.Size(16, 16)
        Me.IconList.ImageStream = CType(resources.GetObject("IconList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.IconList.TransparentColor = System.Drawing.Color.Transparent
        '
        'TimerMain
        '
        '
        'StatusBarPanel4
        '
        Me.StatusBarPanel4.Width = 200
        '
        'FrmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(864, 662)
        Me.Controls.Add(Me.ToolBarMain)
        Me.Controls.Add(Me.StatusBarMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMN
        Me.Name = "FrmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản Lý CTV"
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim c As CMenuItem
    Dim handlerFile As EventHandler = New EventHandler(AddressOf MenuItemClick)
    Dim Meni1 As MenuItem
    Dim IMGLst As New ImageList
    Friend WithEvents ThePrintDocument As System.Drawing.Printing.PrintDocument
    Sub MenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
        Select Case CType(sender, MenuItem).Text
            Case "Danh Sách Đại Lý PSTN   "
                Call_Program(1)
            Case "Nhập Cộng Tác Viên         "
                Call_Program(2)
            Case "Nhập Phiếu Thu                "
                Call_Program(3)
            Case "Nhập Phiếu Chi                  "
                Call_Program(4)
            Case "Calculator                 "
                Shell("calc.exe", vbNormalFocus)
            Case "Thiết Lập Máy In                 "
                Dim value
                PrintDialog1.Document = ThePrintDocument
                value = PrintDialog1.ShowDialog()
                If (value = vbOK) Then
                    strPrinterName = PrintDialog1.PrinterSettings.PrinterName
                End If
            Case "Trung Tâm Thu Cước         "
                Call Call_Program(11)
            Case "Đơn Vị Thu Cước                "
                Call Call_Program(12)
            Case "Danh Mục Ngân Hàng       "
                Call Call_Program(24)
            Case "Danh Mục Tài Khoản         "
                Call Call_Program(25)
            Case "Xuất Sang Excel        "
                Call Call_Program(5)
            Case "Contents              "
            Case "Lập Bảng Kê Nộp Tiền      "
                Call Call_Program(6)
            Case "Cập Nhật Giấy Nộp Tiền    "
                Call Call_Program(8)
            Case "Báo Cáo Quỹ                            "
                Call Call_Program(5)
            Case "Báo Cáo Thu - Chi                    "
                Call Call_Program(9)
            Case "Báo Cáo Chi-Nộp Tiền              "
                Call Call_Program(10)
            Case "Tổng Hợp GNT NH                  "
                Call Call_Program(13)
            Case "Điều Chỉnh Thu          "
                Call Call_Program(14)
            Case "In Hóa Đơn                "
                Call Call_Program(15)
            Case "Báo Cáo Ngày - Tháng             "
                Call Call_Program(16)
            Case "Giao HĐ-TBC Đầu Kỳ"
                Call Call_Program(17)
            Case "Báo Cáo Tổng Hợp Thu           "
                Call Call_Program(18)
            Case "Báo Cáo Thu Theo Kỳ Cước    "
                Call Call_Program(26)
            Case "Đăng Nhập Hệ Thống         "
                Call Call_Program(19)
            Case "Thêm Người dùng               "
                Call Call_Program(20)
            Case "Thay Đổi Password             "
                Call Call_Program(21)
            Case ("Import Dữ liệu            ")
                Call Call_Program(22)
            Case ("Báo Cáo Tỷ Lệ Thu Cước        ")
                Call Call_Program(23)
            Case "Đăng Xuất Hệ Thống          "
                DisableMenu()
                Call Call_Program(19)
            Case "About..                 "
                MsgBox("Chương trình Quản Lý Thu Chi", "Quản Lý CTV")
            Case "Thoát Khỏi Ứng Dụng         "
                Me.Close()
                Application.Exit()
        End Select
    End Sub
    Sub Init()
        MainMN.MenuItems.Add("&Hệ Thống ")
        MainMN.MenuItems.Add("&Chức Năng ")
        MainMN.MenuItems.Add("&Công Cụ ")
        MainMN.MenuItems.Add("&Cửa Sổ Hiện Hành ")
        MainMN.MenuItems(3).MdiList = True
        MainMN.MenuItems.Add("&Trợ Giúp ")
        c.SetImageList = IconList

        'Set meniitem system
        Meni1 = MainMN.MenuItems(0)
        With Meni1
            .MenuItems.Add(New CMenuItem(23, "Đăng Nhập Hệ Thống         ", handlerFile, Shortcut.CtrlL, False))
            .MenuItems.Add(New CMenuItem(24, "Đăng Xuất Hệ Thống          ", handlerFile, Shortcut.CtrlO, False))
            .MenuItems.Add(New CMenuItem(25, "Thay Đổi Password             ", handlerFile, Shortcut.CtrlO, False))
            .MenuItems.Add(New CMenuItem(22, "Thêm Người dùng               ", handlerFile, Shortcut.CtrlN, False))
            .MenuItems.Add(New CMenuItem(14, "Thiết Lập Máy In                 ", handlerFile, Shortcut.CtrlShiftP, False))
            .MenuItems.Add(New CMenuItem(19, "Danh Mục                           ", handlerFile, Shortcut.CtrlD, False))
            .MenuItems.Add(New CMenuItem(0, "Thoát Khỏi Ứng Dụng         ", handlerFile, Shortcut.CtrlShiftX, False))
        End With

        Meni1 = MainMN.MenuItems(0).MenuItems(5)
        With Meni1
            .MenuItems.Add(New CMenuItem(19, "Trung Tâm Thu Cước         ", handlerFile, Shortcut.CtrlShiftA, False))
            .MenuItems.Add(New CMenuItem(19, "Đơn Vị Thu Cước                ", handlerFile, Shortcut.CtrlShiftB, False))
            .MenuItems.Add(New CMenuItem(19, "Danh Mục Ngân Hàng       ", handlerFile, Shortcut.CtrlH, False))
            .MenuItems.Add(New CMenuItem(19, "Danh Mục Tài Khoản         ", handlerFile, Shortcut.CtrlK, False))
        End With
        'Set menuitem Tools
        Meni1 = MainMN.MenuItems(1)
        With Meni1
            .MenuItems.Add(New CMenuItem(11, "Nhập Cộng Tác Viên         ", handlerFile, Shortcut.CtrlShiftV, False))
            .MenuItems.Add(New CMenuItem(1, "Danh Sách Đại Lý PSTN   ", handlerFile, Shortcut.CtrlShiftD, False))
            .MenuItems.Add(New CMenuItem(2, "Nhập Phiếu Thu                ", handlerFile, Shortcut.CtrlShiftT, False))
            .MenuItems.Add(New CMenuItem(16, "Nhập Phiếu Chi                  ", handlerFile, Shortcut.CtrlShiftC, False))
            .MenuItems(3).Enabled = False
            .MenuItems.Add(New CMenuItem(13, "Lập Bảng Kê Nộp Tiền      ", handlerFile, Shortcut.CtrlShiftL, False))
            .MenuItems.Add(New CMenuItem(18, "Cập Nhật Giấy Nộp Tiền    ", handlerFile, Shortcut.CtrlShiftU, False))
        End With

        Meni1 = MainMN.MenuItems(2)
        With Meni1
            .MenuItems.Add(New CMenuItem(5, "Calculator                   ", handlerFile, Shortcut.CtrlShiftM, False))
            .MenuItems.Add(New CMenuItem(16, "Giao HĐ-TBC Đầu Kỳ", handlerFile, Shortcut.CtrlShiftG, False))
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo - Thống Kê  ", handlerFile, Shortcut.CtrlShiftR, False))
            .MenuItems.Add(New CMenuItem(20, "Điều Chỉnh Thu          ", handlerFile, Shortcut.CtrlShift6, False))
            .MenuItems.Add(New CMenuItem(14, "In Hóa Đơn                ", handlerFile, Shortcut.CtrlShift7, False))
            .MenuItems.Add(New CMenuItem(16, "Import Dữ liệu            ", handlerFile, Shortcut.CtrlShift8, False))

        End With

        Meni1 = MainMN.MenuItems(2).MenuItems(2)
        With Meni1
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo Quỹ                            ", handlerFile, Shortcut.CtrlShift1, False))
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo Thu - Chi                    ", handlerFile, Shortcut.CtrlShift2, False))
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo Tổng Hợp Thu           ", handlerFile, Shortcut.CtrlShift3, False))
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo Thu Theo Kỳ Cước    ", handlerFile, Shortcut.CtrlShift4, False))
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo Ngày - Tháng             ", handlerFile, Shortcut.CtrlShift5, False))
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo Chi-Nộp Tiền              ", handlerFile, Shortcut.CtrlShift6, False))
            .MenuItems.Add(New CMenuItem(4, "Tổng Hợp GNT NH                  ", handlerFile, Shortcut.CtrlShift7, False))
            .MenuItems.Add(New CMenuItem(4, "Báo Cáo Tỷ Lệ Thu Cước        ", handlerFile, Shortcut.CtrlShift8, False))
        End With

        Meni1 = MainMN.MenuItems(4)
        With Meni1
            .MenuItems.Add(New CMenuItem(9, "Nội Dung             ", handlerFile, Shortcut.CtrlH, False))
            .MenuItems.Add(New CMenuItem(10, "About..                 ", handlerFile, Shortcut.CtrlA, False))
        End With

        TimerMain.Start()
    End Sub
    Private Sub FrmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Init()
        DisableMenu()
        Call_Program(19)
    End Sub

    Private Sub ToolBarMain_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBarMain.ButtonClick
        Select Case e.Button.ToolTipText
            Case "Danh Sách Đại Lý PSTN"
                Call Call_Program(1)
            Case "Thoát Khỏi Ứng Dụng"
                Me.Close()
                Application.Exit()
            Case "Nhập Mới Cộng Tác Viên"
                Call Call_Program(2)
            Case "Nhập Phiếu Thu"
                Call Call_Program(3)
            Case "Nhập phiếu chi"
                Call Call_Program(4)
            Case "Xuất Sang Excel"
                Call Call_Program(5)
            Case "Lưu"
                Call Call_Program(3)
            Case "Báo Cáo Quỹ"
                Call Call_Program(5)
            Case "Lập bảng kê nộp tiền"
                Call Call_Program(6)
            Case "Thanh toán đại lý"
                Call Call_Program(7)
            Case "Đăng Xuất Hệ Thống"
                DisableMenu()
                Call Call_Program(19)
            Case "Đăng Nhập Hệ Thống"
                Call Call_Program(19)
            Case "Danh Mục Ngân Hàng       "
                Call Call_Program(24)
            Case "Calculator"
                Shell("calc.exe", vbNormalFocus)
            Case "Help"
                MsgBox("Chương trình Quản Lý Thu Chi", , "Quản Lý Thu Chi")
            Case "Windows Version"
                Try
                    Process.Start("winver.exe")
                Catch
                End Try
        End Select
    End Sub
    Private Sub Call_Program(ByVal index As Short)
        Select Case index
            Case 0
                MsgBox("Byte!!!")
            Case 1
                If (Not CheckIfOpen("frmDaily")) Then
                    Dim frm As frmDaily = New frmDaily
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 2
                If (Not CheckIfOpen("frmEmployee")) Then
                    Dim frm As New frmEmployee
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 3
                If (Not CheckIfOpen("frmReceipts")) Then
                    Dim frm As New frmReceipts
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 4
                If (Not CheckIfOpen("frmExpenses")) Then
                    Dim frm As New frmExpenses
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 5
                'frmReports
                
                    'Dim table As DataTable
                    'frm = Me.ActiveMdiChild
                    'table = frm.GetTable()
                    'Export_ToExcel(table)

                If (Not CheckIfOpen("frmReports")) Then
                    Dim frm As New frmReports
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 6
                If (Not CheckIfOpen("frmlistReceipts")) Then
                    Dim frm As New frmlistReceipts
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 7
                If (Not CheckIfOpen("frmThanhtoanDL")) Then
                    Dim frm As New frmThanhtoanDL
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 8
                If (Not CheckIfOpen("frmCapnhatphieuthu")) Then
                    Dim frm As New frmCapnhatphieuthu
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 9
                If (Not CheckIfOpen("frmBaoCaoThuchi")) Then
                    Dim frm As New frmBaoCaoThuchi
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 10
                If (Not CheckIfOpen("frmBaoCaoChiNop")) Then
                    Dim frm As New frmBaoCaoChiNop
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 11
                If (Not CheckIfOpen("frmCountry")) Then
                    Dim frm As New frmCountry
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 12
                If (Not CheckIfOpen("frmTram")) Then
                    Dim frm As New frmTram
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 13
                If (Not CheckIfOpen("frmTongHopGNT")) Then
                    Dim frm As New frmTongHopGNT
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 14
                If (Not CheckIfOpen("frmEditReceipt")) Then
                    Dim frm As New frmEditReceipt
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 15
                If (Not CheckIfOpen("frmPrintOrder")) Then
                    Dim frm As New frmPrintOrder
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 16
                If (Not CheckIfOpen("frmBaoCaoNgayThang")) Then
                    Dim frm As New frmBaoCaoNgayThang
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 17
                If (Not CheckIfOpen("frmGiaoDauKy")) Then
                    Dim frm As New frmGiaoDauKy
                    frm.MdiParent = Me
                    frm.Show()
                End If
            Case 18
                If (Not CheckIfOpen("frmBaoCaoTonghopthu")) Then
                    Dim frm As New frmBaoCaoTonghopthu
                    frm.MdiParent = Me
                    frm.Show()
                End If

            Case 19
                If (Not CheckIfOpen("frmconnection")) Then
                    Dim frm As New frmconnection
                    frm.ShowDialog()
                    If (oledbcon.State = ConnectionState.Open) Then
                        oledbcon.Close()
                        EnableMenu()
                        StatusBarMain.Panels(2).Text = strinfor
                        StatusBarMain.Panels(1).Text = "Người sử dụng :" & UserName
                    Else
                        DisableMenu()
                        StatusBarMain.Panels(2).Text = "Chưa kết nối CSDL"
                        StatusBarMain.Panels(1).Text = "Người sử dụng :"
                    End If
                End If
            Case 20
                If (Not CheckIfOpen("frmNewUser")) Then
                    Dim frm As New frmNewUser
                    frm.ShowDialog()
                End If

            Case 21
                If (Not CheckIfOpen("frmSetpasswordUser")) Then
                    Dim frm As New frmSetpasswordUser
                    frm.ShowDialog()
                End If
            Case 22
                If (Not CheckIfOpen("frmImport")) Then
                    Dim frm As New frmImport
                    frm.ShowDialog()
                End If
            Case 23
                If (Not CheckIfOpen("frmBaoCaoThuCuoc")) Then
                    Dim frm As New frmBaoCaoThuCuoc
                    frm.ShowDialog()
                End If
            Case 24
                If (Not CheckIfOpen("frmBanks")) Then
                    Dim frm As New frmBanks
                    frm.ShowDialog()
                End If
            Case 25
                If (Not CheckIfOpen("frmBankAccounts")) Then
                    Dim frm As New frmBankAccounts
                    frm.ShowDialog()
                End If

            Case 26
                If (Not CheckIfOpen("frmTongHopThuTheoKyCuoc")) Then
                    Dim frm As New frmTongHopThuTheoKyCuoc
                    frm.ShowDialog()
                End If
        End Select
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMain.Tick
        StatusBarMain.Panels(3).Text = Date.Now
    End Sub

    Private Sub FrmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        SaveSetting("QuanLyCTV", "QuanLyCTV", "PrinterName", strPrinterName)
        Application.Exit()
    End Sub
    Private Sub CloseFormIsOpenning()
        Dim frm As Form
        For Each frm In Me.MdiChildren
            frm.Close()
        Next
    End Sub
    Private Function CheckIfOpen(ByVal frmName As String) As Boolean
        Dim frm As Form
        For Each frm In Me.MdiChildren
            If frm.Name = frmName Then
                frm.Focus()
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function
    'Private Sub cmdUpdate(ByVal frmName As String)
    '    'Dim frm As Form
    '    If (CheckIfOpen(frmName)) Then
    '        Dim frmSta As frmStatistics
    '        frmSta = Me.ActiveMdiChild
    '        frmSta.cmdUpdate()
    '    End If
    'End Sub

    Private Sub MnuBaocaoQuy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuBaocaoQuy.Click
        Call Call_Program(5)
    End Sub

    Private Sub MnuBaoCaoChiNop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuBaoCaoChiNop.Click
        Call Call_Program(10)
    End Sub

    Private Sub MnuBaocaoThuChi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuBaocaoThuChi.Click
        Call Call_Program(9)
    End Sub

    Private Sub MnuTonhHopGNT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTonhHopGNT.Click
        Call Call_Program(13)
    End Sub

    Private Sub MenuItemBaoCaoNgay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemBaoCaoNgay.Click
        Call Call_Program(16)
    End Sub

    Private Sub MnuBaocaoTongHopThu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuBaocaoTongHopThu.Click
        Call Call_Program(18)
    End Sub

    Private Sub EnableMenu()
        MainMN.MenuItems(0).MenuItems(0).Enabled = False
        MainMN.MenuItems(0).MenuItems(1).Enabled = True
        MainMN.MenuItems(0).MenuItems(2).Enabled = True
        If (UserName = "sa") Then
            MainMN.MenuItems(0).MenuItems(3).Enabled = True
        End If

        MainMN.MenuItems(0).MenuItems(4).Enabled = True
        MainMN.MenuItems(0).MenuItems(5).Enabled = True
        MainMN.MenuItems(1).Enabled = True
        MainMN.MenuItems(2).Enabled = True
        ToolBarCmdLogin.Enabled = False
        ToolBarCmdLogOut.Enabled = True
        ToolBarcmdNewEmployee.Enabled = True
        ToolBarButtoncmddaily.Enabled = True
        ToolBarButtoncmdPhieuthu.Enabled = True
        ToolBarcmdListReceipts.Enabled = True
        ToolBarcmdFind.Enabled = True
        ToolBarreport.Enabled = True

    End Sub

    Private Sub DisableMenu()
        MainMN.MenuItems(0).MenuItems(0).Enabled = True
        MainMN.MenuItems(0).MenuItems(1).Enabled = False
        MainMN.MenuItems(0).MenuItems(2).Enabled = False
        MainMN.MenuItems(0).MenuItems(3).Enabled = False
        MainMN.MenuItems(0).MenuItems(4).Enabled = False
        MainMN.MenuItems(1).Enabled = False
        MainMN.MenuItems(2).Enabled = False
        ToolBarCmdLogin.Enabled = True
        ToolBarCmdLogOut.Enabled = False
        ToolBarcmdNewEmployee.Enabled = False
        ToolBarButtoncmddaily.Enabled = False
        ToolBarButtoncmdPhieuthu.Enabled = False
        ToolBarcmdListReceipts.Enabled = False
        ToolBarcmdFind.Enabled = False
        ToolBarreport.Enabled = False
    End Sub

End Class
