Imports System
Imports System.Net
Imports System.IO

Public Class ftp

#Region "varibili"

    ''' <summary>
    ''' Host FTP
    ''' </summary>
    Private _Host As String
    ''' <summary>
    ''' Username
    ''' </summary>
    Private _UserName As String
    ''' <summary>
    ''' Password
    ''' </summary>
    Private _Password As String

    ''' <summary>
    ''' Richiesta FTP
    ''' </summary>
    Private _FtpRequest As FtpWebRequest
    ''' <summary>
    ''' Risposta ftp
    ''' </summary>
    Private _FtpResponse As FtpWebResponse

    ''' <summary>
    ''' Flag SSL
    ''' </summary>
    Private _UseSSL As Boolean = False

#End Region

#Region "Propriet?"

    ''' <summary>
    ''' FTP Host URI
    ''' </summary>
    Public Property Host As String
        Get
            Return _Host
        End Get
        Set
            _Host = Value
        End Set
    End Property

    ''' <summary>
    ''' FTP User Name
    ''' </summary>
    Public Property UserName As String
        Get
            Return _UserName
        End Get
        Set
            _UserName = Value
        End Set
    End Property

    ''' <summary>
    ''' FTP Password
    ''' </summary>
    Public Property Password As String
        Get
            Return _Password
        End Get
        Set
            _Password = Value
        End Set
    End Property

    ''' <summary>
    ''' Flag utilizzo SSL
    ''' </summary>
    Public Property UseSSL As Boolean
        Get
            Return _UseSSL
        End Get
        Set
            _UseSSL = Value
        End Set
    End Property

#End Region


    Public Function RemoveDir(ByVal Path As String) As String

        _FtpRequest = CType(WebRequest.Create("ftp://" + _Host + Path), FtpWebRequest)
        _FtpRequest.KeepAlive = False
        _FtpRequest.Credentials = New NetworkCredential(_UserName, _Password)
        _FtpRequest.Method = WebRequestMethods.Ftp.RemoveDirectory
        _FtpRequest.EnableSsl = _UseSSL
        _FtpRequest.Proxy = Nothing
        Try
            _FtpResponse = CType(_FtpRequest.GetResponse(), FtpWebResponse)
            Dim _ResponseStream As Stream = _FtpResponse.GetResponseStream()
            _ResponseStream.Close()
            RemoveDir = "OK"
        Catch ex As Exception
            RemoveDir = ex.Message
        End Try

    End Function

    Public Function ListDirectory(ByVal path As String, ByRef strList As String) As String
        If (path = Nothing Or path = "") Then
            path = "/"
        End If

        _FtpRequest = CType(WebRequest.Create("ftp://" + _Host + path), FtpWebRequest)

        _FtpRequest.Credentials = New NetworkCredential(_UserName, _Password)
        _FtpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails
        _FtpRequest.KeepAlive = False
        _FtpRequest.EnableSsl = _UseSSL
        _FtpRequest.Proxy = Nothing

        Try

            _FtpResponse = CType(_FtpRequest.GetResponse(), FtpWebResponse)

            Dim liststring = ""

            Try
                Dim sr = New StreamReader(_FtpResponse.GetResponseStream(), Text.Encoding.UTF8)
                liststring = sr.ReadToEnd()
                sr.Close()
                _FtpResponse.Close()
                ListDirectory = "5000"
                strList = liststring
            Catch ex As WebException
                Console.WriteLine(CType(ex.Response, FtpWebResponse).StatusDescription)
                _FtpResponse.Close()
                ListDirectory = "0010"
                strList = liststring
            End Try


        Catch ex As Exception
            ListDirectory = "0010" ' Error in list file directory
        End Try

    End Function

    Public Function DownloadFile(ByVal Path As String, ByVal LocalPath As String, ByVal Name As String) As String

        Dim _fileName As String = LocalPath + "\" + Name
        _fileName = Replace(_fileName, "\\", "\")
        Dim _File = New FileInfo(_fileName)
        Dim _FileStream As FileStream

        Try
            _FtpRequest = CType(WebRequest.Create("ftp://" + _Host + Path + Name), FtpWebRequest)

            _FtpRequest.Credentials = New NetworkCredential(_UserName, _Password)
            _FtpRequest.Method = WebRequestMethods.Ftp.DownloadFile
            _FtpRequest.KeepAlive = False
            _FtpRequest.EnableSsl = _UseSSL
            _FtpRequest.Proxy = Nothing

            _FileStream = New FileStream(_File.FullName, FileMode.Create, FileAccess.Write)

            _FtpResponse = CType(_FtpRequest.GetResponse(), FtpWebResponse)

            Dim _ResponseStream As Stream = _FtpResponse.GetResponseStream()

            Dim buffer(1024) As Byte

            Dim bytesRead As Integer = _ResponseStream.Read(buffer, 0, 1024)

            While (bytesRead <> 0)
                _FileStream.Write(buffer, 0, bytesRead)
                bytesRead = _ResponseStream.Read(buffer, 0, 1024)
            End While

            _FileStream.Close()
            _ResponseStream.Close()
            DownloadFile = "5062"  ' all ok
        Catch ex As Exception
            DownloadFile = "0059" ' File Download error
        End Try



    End Function

    Public Function CreateDir(ByVal Path As String) As String
        _FtpRequest = CType(WebRequest.Create("ftp://" + _Host + Path), FtpWebRequest)
        _FtpRequest.Method = WebRequestMethods.Ftp.MakeDirectory
        _FtpRequest.UsePassive = True
        'If behind an HTTP Proxy, MKD command is not available. It is necessary to set an empty proxy for the WebRequest        
        _FtpRequest.Proxy = Nothing
        _FtpRequest.Credentials = New NetworkCredential(_UserName, _Password)
        _FtpRequest.KeepAlive = False
        _FtpRequest.EnableSsl = _UseSSL
        Try
            _FtpResponse = CType(_FtpRequest.GetResponse(), FtpWebResponse)
            Dim _ResponseStream As Stream = _FtpResponse.GetResponseStream()
            _ResponseStream.Close()
            CreateDir = "5000"
        Catch ex As Exception
            CreateDir = "0003" ' Error in create directory
        End Try

    End Function

    Public Function UploadFile(ByVal Path As String, ByVal LocalPath As String, ByVal Name As String) As String
        Dim _fileName As String = LocalPath + "\" + Name
        Dim _File = New FileInfo(_fileName)

        _FtpRequest = CType(WebRequest.Create("ftp://" + _Host + Path + Name), FtpWebRequest)
        _FtpRequest.KeepAlive = False
        _FtpRequest.Credentials = New NetworkCredential(_UserName, _Password)
        _FtpRequest.Method = WebRequestMethods.Ftp.UploadFile
        _FtpRequest.EnableSsl = _UseSSL
        _FtpRequest.Proxy = Nothing

        Dim _fileContents(_File.Length) As Byte

        Dim fr As FileStream = _File.OpenRead()

        fr.Read(_fileContents, 0, Convert.ToInt32(_File.Length))

        fr.Close()
        Try
            Dim writer As Stream = _FtpRequest.GetRequestStream()

            writer.Write(_fileContents, 0, _fileContents.Length)
            writer.Close()
            UploadFile = "5000"
        Catch ex As Exception
            UploadFile = "0004" ' Error in file upload
        End Try

    End Function

    Public Function DeleteFile(ByVal Path As String, ByVal Name As String) As String
        _FtpRequest = CType(WebRequest.Create("ftp://" + _Host + Path + Name), FtpWebRequest)
        _FtpRequest.Credentials = New NetworkCredential(_UserName, _Password)
        _FtpRequest.Method = WebRequestMethods.Ftp.DeleteFile
        _FtpRequest.EnableSsl = _UseSSL
        _FtpRequest.KeepAlive = False
        _FtpRequest.Proxy = Nothing
        Try
            _FtpResponse = CType(_FtpRequest.GetResponse(), FtpWebResponse)
            Dim _ResponseStream As Stream = _FtpResponse.GetResponseStream()
            _ResponseStream.Close()
            DeleteFile = "5000"
        Catch ex As Exception
            DeleteFile = "0005" ' error in delete file
        End Try

    End Function

    Public Sub ResumeDownloadFile(ByVal Path As String, ByVal LocalPath As String, ByVal Name As String)
        Dim _fileName As String = LocalPath + "\" + Name

        Dim _File = New FileInfo(_fileName)
        Dim _FileStream As FileStream

        _FtpRequest = CType(WebRequest.Create("ftp://" + _Host + Path + Name), FtpWebRequest)
        _FtpRequest.KeepAlive = False
        _FtpRequest.Credentials = New NetworkCredential(_UserName, _Password)
        _FtpRequest.Proxy = Nothing
        _FtpRequest.Method = WebRequestMethods.Ftp.DownloadFile

        _FtpRequest.EnableSsl = _UseSSL


        If (_File.Exists) Then
            _FileStream = New FileStream(_File.FullName, FileMode.Append, FileAccess.Write)
        Else
            _FileStream = New FileStream(_File.FullName, FileMode.Create, FileAccess.Write)
        End If

        _FtpResponse = CType(_FtpRequest.GetResponse(), FtpWebResponse)

        Dim _ResponseStream As Stream = _FtpResponse.GetResponseStream()

        Dim buffer(1024) As Byte

        Dim bytesRead As Integer = _ResponseStream.Read(buffer, 0, 1024)

        While (bytesRead <> 0)
            _FileStream.Write(buffer, 0, bytesRead)
            bytesRead = _ResponseStream.Read(buffer, 0, 1024)
        End While

        _FileStream.Close()
        _ResponseStream.Close()

    End Sub

End Class
