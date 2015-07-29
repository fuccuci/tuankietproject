#Region "Oledb Namespace"
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
#End Region
Module [Global]
    Public cur As New Form
    Public strSQL As String
    Public strinfor As String
    Public strPrinterName As String
    Public UserName As String
    Public Password As String
    Public con As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Data\QLCTV.mdb ;Jet OLEDB:Database Password='##^^&&**~`!!$-+/%%an';Persist Security Info=False;"
    'Public con As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Data\QLCTV.mdb;Persist Security Info=False;"
    Public oledbcon As New OleDbConnection(con)
    Public da As OleDbDataAdapter
    Public i As Integer

    Public Sub ExcuxeSQL(ByVal strQuery As String)
        Try
            oledbcon.Open()
            Dim olecommand As New OleDbCommand
            olecommand.CommandText = strQuery
            olecommand.CommandType = CommandType.Text
            olecommand.Connection = oledbcon
            olecommand.ExecuteNonQuery()
            olecommand.Dispose()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        oledbcon.Close()
    End Sub

    Public Function GetStringName(ByVal dsSoure As DataSet, ByVal strAccountCode As String, ByVal strTablename As String, ByVal strColumCode As String, ByVal strColumName As String) As String
        Dim i As Integer
        Dim strresult As String
        For i = 0 To dsSoure.Tables(strTablename).Rows.Count - 1
            strresult = dsSoure.Tables(strTablename).Rows(i).Item(strColumCode)
            If (strresult.Equals(strAccountCode)) Then
                strresult = dsSoure.Tables(strTablename).Rows(i).Item(strColumName)
                Exit For
            End If
        Next
        Return strresult
    End Function

    Public Function CheckReciept_No(ByVal strQuery As String) As Boolean
        Dim value As Boolean = False
        Dim olecommand As OleDbCommand
        Dim oleread As OleDbDataReader
        Try
            oledbcon.Open()
            olecommand = New OleDbCommand
            olecommand.Connection = oledbcon
            olecommand.CommandType = CommandType.Text
            olecommand.CommandText = strQuery
            oleread = olecommand.ExecuteReader
            If (oleread.Read) Then
                value = True
            End If
            olecommand.Dispose()
        Catch ex As Exception
        End Try
        oledbcon.Close()
        Return value
    End Function

    Public Sub SetParameter(ByVal paramDef As ParameterFieldDefinitions, ByVal paramName As String, ByVal paramValue As String)
        Dim crParameterFieldDefinition As ParameterFieldDefinition = paramDef.Item(paramName)
        Dim crParameterValues As ParameterValues = crParameterFieldDefinition.CurrentValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = paramValue
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub

    Public Function SetParaFieldsReportHopDong(ByVal Value As String) As CrystalReport_Receipts
        Dim rpt As New CrystalReport_Receipts
        Try
            Dim paramFields As ParameterFieldDefinitions
            paramFields = rpt.DataDefinition.ParameterFields
            SetParameter(paramFields, "@MaTB", Value)
        Catch ex As Exception
            MessageBox.Show("Lỗi Trong Quá Trình Thiết Lập Tham Số Report: " & ex.Message)
        End Try
        Return rpt
    End Function

    Public Sub Export_ToExcel(ByVal table As DataTable)
        Dim index As Integer
        Dim MyXLApp As Excel.Application
        Dim MyXLBook As Excel.Workbook
        Dim MyXLWorksheet As Excel.Worksheet
        Try
            MyXLApp = CType(CreateObject("Excel.Application"), Excel.Application)
            MyXLBook = CType(MyXLApp.Workbooks.Add, Excel.Workbook)
            MyXLWorksheet = CType(MyXLBook.Worksheets(1), Excel.Worksheet)
            MyXLApp.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        With MyXLWorksheet

            .Columns().ColumnWidth = 15
            .Range("A1").ColumnWidth = 3
            .Range("A1").Value = "CÔNG TY THU CƯỚC VÀ DỊCH VỤ VIETTEL"
            .Range("B1").ColumnWidth = 50
            .Range("A1", "B1").Merge()
            .Range("A2").Value = "Trung tâm thu cước Hồ Chí Minh"
            .Range("A2", "B2").Merge()
            .Range("A3").Value = "Đơn vị thu cước.........."
            .Range("A3", "B3").Merge()

            '.Range("B1").Font.Bold = True
            '.Range("B1").  = Alignment.HorizontalCenterAlign
            '.Range("A1", "C1"). 
            '.Range("B2").Font.Bold = True
            '.Range("B2", "C2").Merge()

            .Range("C1").Value = "Mẫu số BC-10/BKNTNH"
            .Range("C1", "D1").Merge()
            .Range("D1").ColumnWidth = 18
            .Range("C1", "D1").HorizontalAlignment = Alignment.HorizontalCenterAlign

            .Range("E1").ColumnWidth = 20

            .Range("A5").Value = "BẢNG KÊ NỘP TIỀN CƯỚC VÀO TÀI KHOẢN NGÂN HÀNG KỲ CƯỚC......./.........."
            .Range("A5", "E5").Merge()
            .Range("A5", "E5").Font.Bold = True
            .Range("A5", "E5").HorizontalAlignment = Alignment.HorizontalCenterAlign
            .Range("A6").Value = "Ngày ................Tháng ……. Năm………"
            .Range("A6", "E6").Merge()
            'For index = 0 To table.Rows.Count - 1
            '    .Range("A" & (index + 2).ToString).Value = table.Rows(index).Item("SoCV")
            '    .Range("B" & (index + 2).ToString).Value = table.Rows(index).Item("ISDN")
            '    .Range("C" & (index + 2).ToString).Value = table.Rows(index).Item("Customer_Name")
            '    .Range("D" & (index + 2).ToString).Value = table.Rows(index).Item("Customer_Address")
            '    .Range("E" & (index + 2).ToString).Value = table.Rows(index).Item("Reci_Date1")
            '    .Range("F" & (index + 2).ToString).Value = table.Rows(index).Item("Reci_Date2")
            '    .Range("G" & (index + 2).ToString).Value = table.Rows(index).Item("Content_ComPlain")
            '    .Range("H" & (index + 2).ToString).Value = table.Rows(index).Item("ComName")
            '    .Range("I" & (index + 2).ToString).Value = table.Rows(index).Item("StaticName")
            'Next index
        End With

    End Sub
End Module