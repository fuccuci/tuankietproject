Public Class SplitNumbers
    Private strvalue As String
    Private strnumber As String
    Public Sub New()
    End Sub
    Public Sub New(ByVal _strnumber As String)
        strnumber = _strnumber
    End Sub
    Public Property strnumbers() As String
        Get
            Return strnumber
        End Get
        Set(ByVal Value As String)
            strnumber = Value
        End Set
    End Property

    Public Function Splitnumer(ByVal delimiter As String) As String
        Dim len As Integer
        Dim strtam As String
        Dim i As Integer
        Dim sodu As Integer
        Dim count As Integer = 0
        strtam = strnumber
        len = strtam.Length
        sodu = len Mod 3
        While (sodu > 0)
            strtam = "0" + strtam
            len = strtam.Length
            sodu = len Mod 3
            count += 1
        End While
        len = strtam.Length
        i = 3
        While (i < len)
            strtam = strtam.Insert(i, delimiter)
            i += 4
        End While
        If (count > 0) Then
            strtam = strtam.Remove(0, count)
        End If
        Return strtam
    End Function
End Class
