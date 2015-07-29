Imports System.ComponentModel
Imports System.Reflection
Imports System.Resources
Imports System.Drawing.Text

Public Class CMenuItem
    Inherits MenuItem

#Region "Class Members"
    Private mFont As Font = New Font("MS Sans Serif", 8)
    Public Shared mImageList As New ImageList()
    Shared mBojaPozadine As Brush = Brushes.White
    Shared mBojaIvice As Color = SystemColors.Control
    Shared mBojaSelekcije As Color = Color.SlateGray
    Dim sf As New StringFormat()
    Dim mSeparator As Boolean
    Dim mIcon As Image
#End Region

#Region "Constructors"

    Sub New(ByVal text As String, ByVal onClick As EventHandler, Optional ByVal AddSeparator As Boolean = False)
        MyBase.New(text, onClick)
        Me.OwnerDraw = True
        mSeparator = AddSeparator
    End Sub
    Sub New(ByVal text As String, ByVal onClick As EventHandler, ByVal shortcut As Shortcut, Optional ByVal AddSeparator As Boolean = False)
        MyBase.New(text, onClick, shortcut)
        Me.OwnerDraw = True
        mSeparator = AddSeparator
    End Sub

    Sub New(ByVal IndexIcon As Byte, ByVal Text As String, ByVal onClick As EventHandler, Optional ByVal AddSEparator As Boolean = False)
        MyBase.New(Text, onClick)
        Try
            mIcon = mImageList.Images(IndexIcon)
            Me.OwnerDraw = True
            Me.Text = Text
            mSeparator = AddSEparator
        Catch exp As System.ArgumentOutOfRangeException
            MessageBox.Show("Index Icone Ne Postoji !!!", "Pogresno Iniciranje Ikone u meniju", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Sub New(ByVal IndexIcon As Byte, ByVal Text As String, ByVal onClick As EventHandler, ByVal shortcut As Shortcut, Optional ByVal AddSEparator As Boolean = False)
        MyBase.New(Text, onClick, shortcut)
        Try
            mIcon = mImageList.Images(IndexIcon)
            Me.OwnerDraw = True
            Me.Text = Text
            mSeparator = AddSEparator
        Catch exp As System.ArgumentOutOfRangeException
            MessageBox.Show("Index Icone Ne Postoji !!!", "Pogresno Iniciranje Ikone u meniju", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#End Region

#Region "Property"
    Public Shared Property SetImageList() As ImageList
        Get
            Return mImageList
        End Get
        Set(ByVal Value As ImageList)
            If mImageList Is Nothing Then
                mImageList = New ImageList()
            End If
            mImageList = Value
        End Set
    End Property
    Public Shared Property BackColorMenu() As Brush
        Get
            Return mBojaPozadine
        End Get
        Set(ByVal Value As Brush)
            mBojaPozadine = Value
        End Set
    End Property
    Public Shared Property LeftSideBackColor() As Color
        Get
            Return mBojaIvice
        End Get
        Set(ByVal Value As Color)
            mBojaIvice = Value
        End Set
    End Property
#End Region

#Region "Procedure SUB"
    Protected Overrides Sub OnDrawItem(ByVal e As System.Windows.Forms.DrawItemEventArgs)
        MyBase.OnDrawItem(e)
        Dim Br As Brush
        Dim Y As Decimal
        Dim sf As StringFormat = New StringFormat()
        sf.HotkeyPrefix = HotkeyPrefix.Show
        Dim RectPozadine As Rectangle
        Dim RectIvice As Rectangle = e.Bounds
        RectIvice.Width = 20
        RectPozadine = e.Bounds
        Y = e.Bounds.Top + e.Bounds.Height - 2
        MyBase.OnDrawItem(e)
        Y = e.Bounds.Top + e.Bounds.Height - 2
        Dim RecIvice As Rectangle = RectPozadine
        Dim RecSelect As Rectangle = e.Bounds
        RecSelect.Y = e.Bounds.Y + 1
        RecSelect.Height = e.Bounds.Height - 5
        RecSelect.Width = e.Bounds.Width - 1
        RecIvice.X = RecIvice.X + 22
        If Not mIcon Is Nothing Then
            e.Graphics.DrawImage(mIcon, e.Bounds.Left + 2, e.Bounds.Top + 2)
        End If
        If Me.Enabled = False Then
            e.Graphics.FillRectangle(Brushes.Gainsboro, RectPozadine) 'Bojenje Pozadine
            e.Graphics.FillRectangle(Brushes.White, RecIvice) 'Bojenje Pozadine
            Dim r As Rectangle
            r.Size = mIcon.Size
            Dim img As Image
            ControlPaint.DrawImageDisabled(e.Graphics, mIcon, e.Bounds.Left + 2, e.Bounds.Top + 2, Color.Snow)
            e.Graphics.DrawString(GetRealText(), mFont, Brushes.Silver, e.Bounds.Left + 25, e.Bounds.Top + 2, sf)
            Exit Sub
        End If
        If CBool(e.State And DrawItemState.Selected) Then
            e.Graphics.FillRectangle(Brushes.Gainsboro, RectPozadine) 'Bojenje Pozadine
            e.Graphics.FillRectangle(Brushes.White, RecIvice) 'Bojenje Pozadine
            e.Graphics.FillRectangle(Brushes.LightSteelBlue, RecSelect) 'Bijenje Pozadine
            e.Graphics.DrawRectangle(New Pen(Color.Black), RecSelect)
            If mIcon Is Nothing = False Then
                e.Graphics.DrawImage(mIcon, e.Bounds.Left + 2, e.Bounds.Top + 2) 'CrtanjeIkone
                e.Graphics.DrawImage(mIcon, e.Bounds.Left + 3, e.Bounds.Top + 2) 'CrtanjeIkone
            End If
        Else
            e.DrawBackground()
            e.Graphics.FillRectangle(Brushes.Gainsboro, RectPozadine) 'Bojenje Pozadine
            e.Graphics.FillRectangle(mBojaPozadine, RecIvice) 'Bojenje Pozadine
            If mIcon Is Nothing = False Then
                e.Graphics.DrawImage(mIcon, e.Bounds.Left + 2, e.Bounds.Top + 2) 'CrtanjeIkone
            End If
        End If
        Br = New SolidBrush(e.ForeColor.Black)
        e.Graphics.DrawString(GetRealText(), mFont, Brushes.Black, e.Bounds.Left + 25, e.Bounds.Top + 2, sf)
        If mSeparator = True Then
            e.Graphics.DrawLine(New Pen(Color.Silver, 1), e.Bounds.X + 27, Y, e.Bounds.X + e.Bounds.Width, Y)
        End If
    End Sub
    Protected Overrides Sub OnMeasureItem(ByVal e As MeasureItemEventArgs)
        sf.HotkeyPrefix = HotkeyPrefix.Show
        MyBase.OnMeasureItem(e)
        e.ItemHeight = 24
        e.ItemWidth = CInt(e.Graphics.MeasureString(GetRealText(), mFont, 10000, sf).Width) + 20
    End Sub
    
#End Region

#Region "Procedure Function"
    Private Function GetRealText() As String
        Dim s As String = Me.Text
        If ShowShortcut And Shortcut <> Shortcut.None Then
            Dim k As Keys = CType(Shortcut, Keys)
            s = s & TypeDescriptor.GetConverter(GetType(Keys)).ConvertToString(k)
        End If
        Return s
    End Function

#End Region

End Class
