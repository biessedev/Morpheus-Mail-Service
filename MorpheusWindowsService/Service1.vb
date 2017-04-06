Imports System.Configuration
Imports Morpheus_Console

Public Class Service1

     Protected Overrides Sub OnStart(ByVal args() As String)
        'System.Diagnostics.Debugger.Launch()
        Module1.Main()
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
    End Sub


    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Module1.Stp()

    End Sub

End Class
