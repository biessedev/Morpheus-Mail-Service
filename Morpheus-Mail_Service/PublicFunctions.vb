Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Data.SqlClient
Imports System.Globalization

Module PublicFunctions
    Public MySQLConnectionString As String
    Public DBName As String
    Public ConnectionString As String
    Public strFtpServerAdd As String
    Public strFtpServerUser As String
    Public strFtpServerPsw As String
    Public tblError As DataTable
    Public DsError As New DataSet
    Public CultureInfo_ja_JP As New CultureInfo("ja-JP", False)

    Public Function OpenConnectionMySql(ByVal strHost As String, ByVal strDatabase As String, ByVal strUserName As String, ByVal strPassword As String)
        Dim conn = "host=" & strHost & ";" & "username=" & strUserName & ";" & "password=" & strPassword & ";" & "database=" & strDatabase & ";Connect Timeout=120;allow zero datetime=true;charset=utf8; "
        Try
            Dim mysqlconn = New MySqlConnection(conn)
            mysqlconn.Open()
            Return mysqlconn
        Catch ae As MySqlException
            Return New MySqlConnection()
        End Try
    End Function

    Sub CloseConnectionMySql()
        Try
            If MySqlconnection.State = ConnectionState.Open Then
                MySqlconnection.Close()
            End If
        Catch ae As MySqlException
        End Try
    End Sub

    Function ParameterTableWrite(ByVal param As String, ByVal value As String) As String
        ParameterTableWrite = "KO"
        Try
            Dim sql As String = "UPDATE `" & DBName & "`.`parameterset` SET `value` ='" & value & "' where name = '" & param & "'"
            Dim cmd = New MySqlCommand(sql, MySqlconnection)
            cmd.ExecuteNonQuery()
            ParameterTableWrite = "OK"
        Catch ex As Exception
            'MsgBox("Parametric Write error!   " & ex.Message)
        End Try
    End Function

    'Write and get the time of server.
    Function ParameterTable(ByVal param As String) As String
        Try
            Dim Adapter As New MySqlDataAdapter("SELECT * FROM parameterset", MySqlconnection)
            Dim tbl As DataTable
            Dim Ds As New DataSet, resultRow As DataRow()
            Adapter.Fill(Ds, "parameterset")
            tbl = Ds.Tables("parameterset")
            resultRow = tbl.Select("name = '" & param & "'")
            If resultRow.Length > 0 Then
                ParameterTable = resultRow(0).Item("value").ToString()
            End If
            Adapter.Dispose()
            Ds.Dispose()
        Catch ex As Exception
            'MsgBox("Error: " & ex.Message)
        End Try
    End Function

    Function string_to_date(ByVal Indate As String) As Date
        If Len(Indate) >= 8 Then string_to_date = DateTime.Parse(Indate, CultureInfo_ja_JP.DateTimeFormat)
    End Function

    Function date_to_string(ByVal Indate As Date) As String
        date_to_string = Indate.Year & "/" & Mid("0" & Indate.Month, Len(Trim(Str(Indate.Month))), 2) & "/" & Mid("0" & Indate.Day, Len(Trim(Str(Indate.Day))), 2)
    End Function
End Module
