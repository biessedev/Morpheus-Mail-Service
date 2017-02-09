Imports System.Configuration
Imports System.Threading

Public Module Module1

    Sub Main()
        Dim thread As New Thread(AddressOf Start)
        thread.Start()

    End Sub

    Private Sub Start()

        Dim builder As New Common.DbConnectionStringBuilder()
        builder.ConnectionString = ConfigurationManager.ConnectionStrings("Morpheus").ConnectionString
        Dim test As String = builder("host") + builder("database") + builder("username") + builder("password")
        Dim timer As New TimerECR(builder("host"), builder("database"), builder("username"), builder("password"))
        timer.TimerECR_Tick()
        'Dim timer As New TimerECR("127.0.0.1", "srvdoc")
        'timer.TimerECR_Tick()
    End Sub
End Module