Imports System.Data.OleDb
Public Class ClassCreateUser
    Private strUser As String
    Private strPass As String
    Private OleAdapter As OleDbDataAdapter
    Private ds As DataSet
    Public Sub New()
        ds = New DataSet
    End Sub

    Public Sub New(ByVal _strUser As String, ByVal _strPass As String)
        strUser = _strUser
        strPass = _strPass
        ds = New DataSet
    End Sub

    Public Property UserName() As String
        Get
            Return strUser
        End Get
        Set(ByVal Value As String)
            strUser = Value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return strPass
        End Get
        Set(ByVal Value As String)
            strPass = Value
        End Set
    End Property

    Public Sub CreateAccount(ByVal _strUser As String, ByVal _strPass As String)
        strUser = _strUser
        strPass = _strPass
        Select Case CheckFillData()
            Case 1
                MsgBox("UserName chưa được nhập!!!")
            Case 2
                MsgBox("Password chưa được nhập!!!")
            Case 3
                Dim sql As String
                sql = "INSERT INTO Tbl_UserPass(UserName,Passwrd) VALUES('" & strUser & "','" & strPass & "')"
                Try
                    oledbcon.Open()
                    Dim cmd As New OleDbCommand(sql, oledbcon)
                    cmd.ExecuteNonQuery()
                    MsgBox("Tên truy cập :" & strUser & " đã được tạo.", MsgBoxStyle.Exclamation, "Thông báo")
                    oledbcon.Close()
                Catch ex As Exception
                    MsgBox("Lổi rồi người ơi!!!")
                End Try
        End Select

    End Sub

    Public Sub CreateAccount()
        Select Case CheckFillData()
            Case 1
                MsgBox("UserName chưa được nhập!!!")
            Case 2
                MsgBox("Password chưa được nhập!!!")
            Case 3
                Dim sql As String
                sql = "INSERT INTO Tbl_UserPass(UserName,Passwrd) VALUES('" & strUser & "','" & strPass & "')"
                Try
                    oledbcon.Open()
                    Dim cmd As New OleDbCommand(sql, oledbcon)
                    cmd.ExecuteNonQuery()
                    MsgBox("Tên truy cập " & strUser & " đã được tạo.", MsgBoxStyle.Exclamation, "Thông báo")
                    oledbcon.Close()
                Catch ex As Exception
                    MsgBox("Lổi rồi người ơi!!!")
                End Try
        End Select

    End Sub

    Public Sub DeleteAccount()
        Dim sql As String
        sql = "DELETE FROM Tbl_UserPass WHERE UserName ='" & strUser & "'"
        Try
            oledbcon.Open()
            Dim cmd As New OleDbCommand(sql, oledbcon)
            cmd.ExecuteNonQuery()
            MsgBox("Account " & strUser & " đã được xóa.", MsgBoxStyle.Exclamation, "Thông báo")
            oledbcon.Close()
        Catch ex As Exception
            MsgBox("Lổi rồi người ơi!!!")
        End Try

    End Sub

    Public Sub SetPassword()
        Dim sql As String
        sql = "UPDATE Tbl_UserPass SET Passwrd ='" & strPass & "'  WHERE UserName ='" & strUser & "'"
        Try
            oledbcon.Open()
            Dim cmd As New OleDbCommand(sql, oledbcon)
            cmd.ExecuteNonQuery()
            MsgBox("Password " & strUser & " đã được thay đổi.", MsgBoxStyle.Exclamation, "Thông báo")
            oledbcon.Close()
        Catch ex As Exception
            MsgBox("Lổi rồi người ơi!!!")
        End Try

    End Sub

    Public Function SelectUsers() As DataSet
        Dim strQuery As String
        strQuery = "SELECT * FROM Tbl_UserPass WHERE UserName <>'sa'"
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            OleAdapter = New OleDbDataAdapter(cmd)
            OleAdapter.Fill(ds, "UserPass")
            OleAdapter.Dispose()
            cmd.Dispose()
        Catch ex As Exception
            MsgBox("Lổi rồi người ơi!!!")
        End Try
        Return ds
    End Function

    Public Function CheckUserPass() As Boolean
        Dim strQuery As String
        strQuery = "SELECT * FROM Tbl_UserPass WHERE UserName ='" & strUser & "'"
        Try
            Dim cmd As New OleDbCommand(strQuery, oledbcon)
            Dim dsCheck As New DataSet
            OleAdapter = New OleDbDataAdapter(cmd)
            OleAdapter.Fill(dsCheck, "UserPass")
            OleAdapter.Dispose()
            cmd.Dispose()
            If (dsCheck.Tables(0).Rows.Count > 0) Then
                Return False
            End If
        Catch ex As Exception
            MsgBox("Lổi rồi người ơi!!!")
        End Try
        Return True
    End Function

    Private Function CheckFillData() As Short
        If (Trim$(strUser) = "") Then
            Return 1
        End If

        If (Trim$(strPass) = "") Then
            Return 2
        End If
        Return 3
    End Function

End Class
