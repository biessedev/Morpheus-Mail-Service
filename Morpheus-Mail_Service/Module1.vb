Imports System.Configuration
Imports System.Threading
Imports MySql.Data.MySqlClient

Public Module Module1
    Dim thread As New Thread(AddressOf Start)
    Public MySqlconnection As MySqlConnection

    Sub Main()
        thread.Start()
    End Sub

    Public Sub Start()
        Dim builder As New Common.DbConnectionStringBuilder()
        builder.ConnectionString = ConfigurationManager.ConnectionStrings("Morpheus").ConnectionString
        Dim timer As New TimerECR(builder("host"), builder("database"), builder("username"), builder("password"))
        timer.TimerECR_Tick()
    End Sub

    Sub Stp()
        thread.Abort()
    End Sub
End Module