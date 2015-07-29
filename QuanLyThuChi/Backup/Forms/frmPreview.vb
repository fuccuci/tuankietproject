Public Class frmPreview
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
    Friend WithEvents CrystalReportViewerReceipts As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CrystalReportViewerReceipts = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'CrystalReportViewerReceipts
        '
        Me.CrystalReportViewerReceipts.ActiveViewIndex = -1
        Me.CrystalReportViewerReceipts.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrystalReportViewerReceipts.DisplayGroupTree = False
        Me.CrystalReportViewerReceipts.Location = New System.Drawing.Point(8, 8)
        Me.CrystalReportViewerReceipts.Name = "CrystalReportViewerReceipts"
        Me.CrystalReportViewerReceipts.ReportSource = Nothing
        Me.CrystalReportViewerReceipts.ShowGotoPageButton = False
        Me.CrystalReportViewerReceipts.ShowGroupTreeButton = False
        Me.CrystalReportViewerReceipts.ShowTextSearchButton = False
        Me.CrystalReportViewerReceipts.Size = New System.Drawing.Size(808, 520)
        Me.CrystalReportViewerReceipts.TabIndex = 0
        '
        'frmPreview
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(816, 534)
        Me.Controls.Add(Me.CrystalReportViewerReceipts)
        Me.Name = "frmPreview"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmPreview"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmPreview_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CrystalReportViewerReceipts.Zoom(75)
    End Sub
End Class
