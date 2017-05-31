Imports MySql.Data.MySqlClient
Imports System.Net.Mail
Imports System.Net
Imports System.Globalization
Imports System.Windows.Forms


Public Class TimerECR

    Public host As String
    Public database As String
    Public userName As String
    Public password As String
    Dim AdapterDoc As New MySqlDataAdapter("SELECT * FROM doc", MySqlconnection)
    Dim AdapterDocType As New MySqlDataAdapter("SELECT * FROM Doctype", MySqlconnection)
    Dim AdapterEcr As New MySqlDataAdapter("SELECT * FROM Ecr", MySqlconnection)
    Dim AdapterProd As New MySqlDataAdapter("SELECT * FROM product", MySqlconnection)
    Dim Adaptermail As New MySqlDataAdapter("SELECT * FROM mail", MySqlconnection)
    Dim tblDoc As DataTable, tblDocType As DataTable, tblEcr As DataTable, tblProd As DataTable, tblmail As DataTable
    Dim DsDoc As New DataSet, DsDocType As New DataSet, DsEcr As New DataSet, DsProd As New DataSet, Dsmail As New DataSet
    Dim cmd As New MySqlCommand()
    Dim MailSent As Boolean
    Dim dep As New List(Of String)
    Dim RichTextBoxConv As New RichTextBox()
    Dim w As IO.StreamWriter

    Sub New(host As String, database As String, userName As String, password As String)
        Me.host = host
        Me.database = database
        Me.userName = userName
        Me.password = password
        DBName = database
        MySqlconnection = OpenConnectionMySql(host, database, userName, password)
        strFtpServerUser = ParameterTable("MorpheusFtpUser")
        strFtpServerPsw = ParameterTable("MorpheusFtpPsw")
        strFtpServerAdd = ParameterTable("PathDocument") & DBName & "/"
        dep.Add("U")
        dep.Add("L")
        dep.Add("C")
        dep.Add("B")
        dep.Add("E")
        dep.Add("N")
        dep.Add("P")
        dep.Add("Q")
        dep.Add("F")
        dep.Add("S")
    End Sub

    Public Sub TimerECR_Tick()

        AdapterEcr.SelectCommand = New MySqlCommand("SELECT * FROM ecr;", MySqlconnection)
        AdapterEcr.Fill(DsEcr, "ecr")
        tblEcr = DsEcr.Tables("ecr")

        AdapterDoc.SelectCommand = New MySqlCommand("SELECT * FROM DOC", MySqlconnection)
        AdapterDoc.Fill(DsDoc, "doc")
        tblDoc = DsDoc.Tables("doc")

        AdapterProd.SelectCommand = New MySqlCommand("SELECT * FROM product;", MySqlconnection)
        AdapterProd.Fill(DsProd, "product")
        tblProd = DsProd.Tables("product")

        Adaptermail.SelectCommand = New MySqlCommand("SELECT * FROM mail;", MySqlconnection)
        Adaptermail.Fill(Dsmail, "mail")
        tblmail = Dsmail.Tables("mail")

        ParameterTableWrite("SYSTEM_SCHEDULE", "RUN")

        If Now.DayOfWeek <> DayOfWeek.Saturday And Now.DayOfWeek <> DayOfWeek.Sunday Then
            OpenConnectionMySql(host, database, userName, password)
            UpdateEcrTable()
            EcrMailScheduler()
            ecrDocConfirm()
            ecrDocApprove()
            ecrDocSign()
        End If

        If Now.DayOfWeek <> DayOfWeek.Saturday And Now.DayOfWeek <> DayOfWeek.Sunday Then
            OpenConnectionMySql(host, database, userName, password)
            TCRMailScheduler()
        End If

        ' DOC
        If Now.DayOfWeek <> DayOfWeek.Saturday And Now.DayOfWeek <> DayOfWeek.Sunday Then
            OpenConnectionMySql(host, database, userName, password)
            DocMailScheduler()
        End If

        ' Status
        If Now.DayOfWeek <> DayOfWeek.Saturday And Now.DayOfWeek <> DayOfWeek.Sunday Then
            OpenConnectionMySql(host, database, userName, password)
            StatusMailScheduler()
        End If

        Adaptermail.SelectCommand = New MySqlCommand("SELECT * FROM mail;", MySqlconnection)
        Adaptermail.Fill(Dsmail, "mail")
        tblmail = Dsmail.Tables("mail")

        Dim RowSearch As DataRow(), i As Integer, j As Integer
        RowSearch = tblmail.Select("name like '*'")
        For Each row In RowSearch
            j = Len(row("freq").ToString)
            If j > 1000 Then
                i = InStrRev(row("freq").ToString, "]", j - 1000, CompareMethod.Text)
                If i > 1 Then
                    WriteField("freq", Mid(row("freq").ToString, i + 1), row("id").ToString)
                End If
            End If
        Next

        ParameterTableWrite("LAST_AUTOMATIC_SCHEDULER", date_to_string(Today))
        ParameterTableWrite("SYSTEM_SCHEDULE", "HOLD")
        CloseConnectionMySql()
    End Sub

    Sub UpdateEcrTable()

        Dim RowEcr As DataRow(), pos As Integer
        Dim EcrN As Integer, sql As String, filename As String
        Dim RowSearchDoc As DataRow()

        RowSearchDoc = tblDoc.Select("header = '" & ParameterTable("plant") & "R_PRO_ECR'")

        For Each row In RowSearchDoc
            AdapterEcr.SelectCommand = New MySqlCommand("SELECT * FROM ecr;", MySqlconnection)
            tblEcr.Clear()
            DsEcr.Clear()
            AdapterEcr.Fill(DsEcr, "ecr")
            tblEcr = DsEcr.Tables("ecr")

            pos = InStr(1, row("filename").ToString, "-", CompareMethod.Text)
            EcrN = Val(Mid(row("filename").ToString, 1, pos))
            RowEcr = tblEcr.Select("number=" & EcrN)
            If EcrN > 0 And RowEcr.Length = 0 And InStr(row("filename").ToString, "template", CompareMethod.Text) <= 0 Then
                Try
                    filename = row("filename").ToString & "_" & row("rev").ToString & "." & row("extension").ToString
                    sql = "INSERT INTO `" & DBName & "`.`ecr` (`nnote` ,`number` ,`description` ,`date`,`Usign`,`nsign`,`Lsign`,`Asign`,`Qsign`,`Esign`,`Rsign`,`Psign`,`Bsign`,`Ssign`,`DocInvalid`,`IdDoc`,`CLCV`) VALUES (" &
                    Replace("'{\rtf1\fbidis\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}{\f1\fswiss\fprq2\fcharset0 Calibri;}}{\colortbl ;\red23\green54\blue93;}\viewkind4\uc1\pard\ltrpar\sl360\slmult1\cf1\lang1040\f0\fs22\par\par\par\par\ul\b\i\f1 Confirmation AREA\par\lang1033\ulnone\b0\i0 Time and First serial number / Fiche:\par\par\par\parOther Annotation:\f0\par\pard\ltrpar\cf0\lang1040\fs24\par\par\par\par}', ", "\", "\\") _
                    & EcrN & ", '" & filename & "', '" & "01/01/2000" & "', 'NOT CHECKED' , 'NOT CHECKED', 'NOT CHECKED', 'System[automatic]', 'NOT CHECKED', 'NOT CHECKED', 'NOT CHECKED', 'NOT CHECKED', 'NOT CHECKED','NOT CHECKED', 'NO', " & row("id").ToString & ",'NO');"
                    cmd = New MySqlCommand(sql, MySqlconnection)
                    cmd.ExecuteNonQuery()

                Catch ex As Exception
                    ComunicationLog("5050") ' Mysql update query error, check if bitron p/n is already in db
                End Try
            End If
        Next
    End Sub

    Sub ComunicationLog(ByVal ComCode As String)
        Dim rsResult As DataRow()
        rsResult = tblError.Select("code='" & ComCode & "'")
        If rsResult.Length = 0 Then
            ComCode = "0051"
            rsResult = tblError.Select("code='" & ComCode & "'")
        End If
        Using w As IO.StreamWriter = IO.File.AppendText("errorLog.txt")
            w.Write(vbCrLf + Date.Now + "-> Error : " + rsResult(0).Item("en").ToString)
        End Using
    End Sub

    Private Function getAllDepartmentInitialsForAutomaticSrvDocMessage() As String
        Return "ARULBENPQS"
    End Function

    Sub EcrMailScheduler()
        Dim refresh = True
        Dim RowSearchEcr As DataRow() = tblEcr.Select("")
        For Each row In RowSearchEcr
            If readDocSign(row("iddoc").ToString, refresh) = "" Then
                If row("ecrcheck").ToString <> "YES" Then
                    mailSender("ECR_" & "VerifyTo", "ECR_" & "VerifyCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf & "Please VERIFY the ECR: " & " " & row("description").ToString, "ECR Check Request " & " " & row("description").ToString, row("number").ToString)
                End If
                Dim us As String = getAllDepartmentInitialsForAutomaticSrvDocMessage()
                For Each c As String In us
                    If ((row(c & "sign").ToString = "NOT CHECKED") And (row("ecrcheck").ToString = "YES")) Then
                        mailSender("ECR_" & c & "_SignTo", "ECR_" & c & "_SignCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf & "Please CHECK the ECR: " & " " & row("description").ToString, "ECR Check Request " & " " & row("description").ToString, row("number").ToString)
                    End If
                Next
                Dim dt As Date = string_to_date((row("date").ToString))

                If row("Rsign").ToString <> "NOT CHECKED" And
                row("Usign").ToString <> "NOT CHECKED" And
                row("Lsign").ToString <> "NOT CHECKED" And
                row("Bsign").ToString <> "NOT CHECKED" And
                row("Esign").ToString <> "NOT CHECKED" And
                row("Nsign").ToString <> "NOT CHECKED" And
                row("Psign").ToString <> "NOT CHECKED" And
                row("Qsign").ToString <> "NOT CHECKED" And
                row("Ssign").ToString <> "NOT CHECKED" Then

                    us = getAllDepartmentInitialsForAutomaticSrvDocMessage()
                    For Each c As String In us
                        If row(c & "sign").ToString = "CHECKED" Then
                            mailSender("ECR_" & c & "_SignTo", "ECR_" & c & "_SignCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf & "Please APPROVE the Ecr: " & " " & row("description").ToString, "ECR Approval Request " & row("description").ToString, row("number").ToString & "A")
                        End If
                    Next
                End If
                If InStr(row("Rsign").ToString & row("Usign").ToString & row("Lsign").ToString & row("Bsign").ToString & row("Esign").ToString & row("Nsign").ToString & row("Psign").ToString & row("Qsign").ToString & row("Asign").ToString & row("Ssign").ToString, "CHECKED", CompareMethod.Text) <= 0 Then

                    us = getAllDepartmentInitialsForAutomaticSrvDocMessage()
                    For Each c As String In us
                        If row(c & "sign").ToString = "APPROVED" Then
                            mailSender("ECR_" & c & "_SignTo", "ECR_" & c & "_SignCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf & "Please SIGN the Ecr: " & " " & row("description").ToString, "ECR Sign Request " & row("description").ToString, row("number").ToString & "S")
                        End If
                    Next
                End If
            End If
            refresh = False
        Next
    End Sub

    Function readDocSign(ByVal docId As Long, ByVal refresh As Boolean) As String
        Dim tblDoc As DataTable
        Dim DsDoc As New DataSet

        AdapterDoc.SelectCommand = New MySqlCommand("SELECT * FROM DOC", MySqlconnection)
        AdapterDoc.Fill(DsDoc, "doc")
        tblDoc = DsDoc.Tables("doc")

        Dim Res As DataRow() = tblDoc.Select("id = " & docId)
        If Res.Length > 0 Then
            readDocSign = Res(0).Item("sign").ToString
        End If

    End Function

    Function mailSender(ByVal AddlistTo As String, ByVal AddlistCopy As String, ByVal bodyText As String, ByVal SubText As String, ByVal Necr As String, Optional ByVal freq As Boolean = True, Optional ByVal ATTACH As String = "") As Boolean
        Dim freqTo = ""
        Dim dt As Date = Now
        tblmail.Clear()
        Dsmail.Clear()
        mailSender = False
        Adaptermail.SelectCommand = New MySqlCommand("SELECT * FROM mail;", MySqlconnection)
        Adaptermail.Fill(Dsmail, "mail")
        tblmail = Dsmail.Tables("mail")

        Dim client As New SmtpClient(ParameterTable("SMTP"), ParameterTable("SMTP_PORT"))
        client.EnableSsl = IIf(ParameterTable("MAIL_SSL") = "YES", True, False)
        client.Credentials = New NetworkCredential(ParameterTable("MAIL_SENDER_CREDENTIAL_USER"), ParameterTable("MAIL_SENDER_CREDENTIAL_PSW"))

        Dim msg As New MailMessage(ParameterTable("MAIL_SENDER_CREDENTIAL_MAIL"), ParameterTable("MAIL_SENDER_CREDENTIAL_MAIL"))

        Dim RowSearchMail As DataRow() = tblmail.Select("list = '" & AddlistTo & "'")
        msg.To.Clear()
        msg.CC.Clear()

        For Each row In RowSearchMail
            msg.To.Add(row("name").ToString)
            freqTo = row("freq").ToString
        Next

        RowSearchMail = tblmail.Select("list = '" & AddlistCopy & "'")
        For Each row In RowSearchMail
            msg.CC.Add(row("name").ToString)
        Next

        If ATTACH <> "" Then
            Dim Allegato = New Attachment(ATTACH)
            If My.Computer.FileSystem.GetFileInfo(ATTACH).Length < Val(ParameterTable("MAX_SIZE_FILE_MAIL")) Then
                msg.Attachments.Add(Allegato)
                msg.Body = bodyText
            Else
                msg.Body = "ATTENTION... FILE NOT SENT BY MAIL FOR EXCESSIVE DIMENSION. PLEASE DOWNLOAD FROM SERVER!!!" & vbCrLf & vbCrLf & bodyText
            End If
        Else
            msg.Body = bodyText
        End If

        msg.Subject = SubText

        If freq = False Then
            freqTo = ""
        End If

        Try
            If DayOfWeek.Saturday <> dt.DayOfWeek And DayOfWeek.Sunday <> dt.DayOfWeek And (dt.Hour > 8 And dt.Hour < 20) Then
                'Dim rowEcr As DataRow() = tblEcr.Select("number = '" & Necr & "'")
                'Dim parsedDate
                'Dim dateList As New List(Of DateTime)
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("date"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateR"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateU"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateL"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateB"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateE"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateN"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateQ"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Add(If(DateTime.TryParse(rowEcr(0).Item("dateP"), parsedDate), parsedDate, Date.Parse("1/1/2012 12:00:00 AM")))
                'dateList.Sort()
                'dateList.Reverse()

                If (InStr(freqTo, "[" & Necr & "]", CompareMethod.Text) <= 0) Then
                    'Or                   ((InStr(freqTo, "[" & Necr & "]", CompareMethod.Text) > 0) And (DateDiff(DateInterval.Day, dateList.ElementAt(0), DateTime.Now) >= 3)) Then
                    client.Send(msg)
                    MailSent = True
                    Console.WriteLine("E mail sent: " & SubText & "  " & Mid(msg.To.Item(0).ToString, 1, 45) & " ....")
                    mailSender = True
                    Application.DoEvents()
                    Application.DoEvents()
                    RowSearchMail = tblmail.Select("list = '" & AddlistTo & "'")
                    For Each row In RowSearchMail
                        WriteField("freq", row("freq").ToString & "[" & Necr & "]", row("id").ToString)
                    Next
                    RowSearchMail = tblmail.Select("list = '" & AddlistCopy & "'")
                    For Each row In RowSearchMail
                        WriteField("freq", row("freq").ToString & "[" & Necr & "]", row("id").ToString)
                    Next
                End If
            End If

        Catch ex As Exception
            Console.WriteLine("Mail not sent...!!!")
        End Try
        Application.DoEvents()
    End Function

    Sub WriteField(ByVal field As String, ByVal v As String, ByVal list As String)
        Try
            Dim SQL As String = "UPDATE `" & DBName & "`.`mail` SET `" & field & "` = '" & v & "' WHERE `mail`.`id` = " & list & " ;"
            cmd = New MySqlCommand(SQL, MySqlconnection)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            ComunicationLog("0052") 'db operation error
        End Try
    End Sub

    Sub ecrDocConfirm()
        Dim sql As String, refresh = True
        Dim RowSearchEcr As DataRow() = tblEcr.Select("docInvalid = 'NO'", "number")
        For Each row In RowSearchEcr
            If InStr(row("Rsign").ToString & row("Lsign").ToString & row("Usign").ToString & row("Bsign").ToString & row("Esign").ToString & row("Nsign").ToString & row("Psign").ToString & row("Qsign").ToString & row("Asign").ToString & row("Ssign").ToString, "APPROVED", CompareMethod.Text) <= 0 And readDocSign(row("iddoc").ToString, refresh) <> "" And
                            row("confirm").ToString = "CONFIRMED" Then

                Dim fileOpen As String = downloadFileWinPath(ParameterTable("plant") & "R_PRO_ECR_" & row("DESCRIPTION").ToString, ParameterTable("plant") & "R/" & ParameterTable("plant") & "R_PRO_ECR/")
                Try
                    If mailSender("Status_SignTo", "Status_SignCopy", "Automatic SrvDoc Message:" & vbCrLf &
                               vbCrLf & row("description").ToString & " -- > (Result: Confirmation of ECR Introduction) " & vbCrLf & vbCrLf &
                               vbCrLf & "Validate Data :" & row("date").ToString & " (yyyy/mm/dd)" & vbCrLf &
                               vbCrLf & vbCrLf & "Quality Note: " & rtfTrans(row("nnote").ToString) & vbCrLf &
                               vbCrLf & vbCrLf & vbCrLf & "For all detailed info please download ECR from server SrvDoc.", "ECR - Confirmation of Introduction:   " & " " & row("description").ToString, "C" & row("number").ToString, False, fileOpen) Then
                        sql = "UPDATE `" & DBName & "`.`ECR` SET `confirm` = 'SENT_CONFIRMED' WHERE `ECR`.`id` = " & row("id").ToString & " ;"
                        cmd = New MySqlCommand(sql, MySqlconnection)
                        cmd.ExecuteNonQuery()
                    End If
                    cmd = New MySqlCommand(sql, MySqlconnection)
                    cmd.ExecuteNonQuery()

                Catch ex As Exception
                    ComunicationLog("0052") 'db operation error
                End Try
            End If
            refresh = False
        Next
    End Sub

    Function downloadFileWinPath(ByVal fileName As String, ByVal strPathFtp As String) As String
        Dim objFtp = New ftp()
        objFtp.UserName = strFtpServerUser
        objFtp.Password = strFtpServerPsw
        objFtp.Host = strFtpServerAdd
        downloadFileWinPath = ""

        If fileName <> "" Then
            Try
                ComunicationLog(objFtp.DownloadFile(strPathFtp, IO.Path.GetTempPath, fileName)) ' download successfull
                downloadFileWinPath = IO.Path.GetTempPath & fileName
            Catch ex As Exception
                ComunicationLog("0049") ' Error in ecr Download
            End Try
        Else
            ComunicationLog("5061") ' fill path
        End If

    End Function

    Function rtfTrans(ByVal rtf As String) As String
        Try
            RichTextBoxConv.Rtf = rtf
            rtfTrans = RichTextBoxConv.Text
        Catch ex As Exception
            rtfTrans = ""
        End Try
    End Function

    Sub ecrDocApprove()
        Dim RowSearchEcr As DataRow() = tblEcr.Select("docInvalid = 'NO'", "number")
        For Each row In RowSearchEcr
            Dim i As Integer
            i = Int(row("number").ToString)
            If InStr(row("Rsign").ToString & row("Usign").ToString & row("Lsign").ToString & row("Bsign").ToString & row("Esign").ToString & row("Nsign").ToString & row("Psign").ToString & row("Qsign").ToString & row("Asign").ToString & row("Ssign").ToString, "CHECKED", CompareMethod.Text) <= 0 And row("approve").ToString = "" Then
                Try
                    Dim fileOpen As Object = downloadFileWinPath(ParameterTable("plant") & "R_PRO_ECR_" & row("DESCRIPTION").ToString, ParameterTable("plant") & "R/" & ParameterTable("plant") & "R_PRO_ECR/")
                    If mailSender("ECR_SignTo", "ECR_SignCopy", "Automatic SrvDoc Message:" & vbCrLf &
                               vbCrLf & "R&D LT: " & row("leadTimeR").ToString & If(row("leadTimeR").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Purchasing LT: " & row("leadTimeU").ToString & If(row("leadTimeU").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Logistic LT: " & row("leadTimeL").ToString & If(row("leadTimeL").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Process Engineering LT: " & row("leadTimeB").ToString & If(row("leadTimeB").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Testing Engineering LT: " & row("leadTimeE").ToString & If(row("leadTimeE").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Quality LT: " & row("leadTimeN").ToString & If(row("leadTimeN").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Production LT: " & row("leadTimeP").ToString & If(row("leadTimeP").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Time & Methods LT: " & row("leadTimeQ").ToString & If(row("leadTimeQ").ToString.Equals("1"), " week", " weeks") &
                               vbCrLf & "Environment & Safety LT: " & row("leadTimeS").ToString & If(row("leadTimeS").ToString.Equals("1"), " week", " weeks") & vbCrLf &
                               vbCrLf & row("description").ToString & " -- > (Result: Approved) " &
                               vbCrLf & "Approval Data : " & row("date").ToString & "( yyyy/mm/dd )" & vbCrLf &
                               vbCrLf & vbCrLf & "R&D Note: " & rtfTrans(row("rnote").ToString) & vbCrLf &
                               vbCrLf & "Logistic Note: " & rtfTrans(row("lnote").ToString) & vbCrLf &
                               vbCrLf & "Purchasing Note: " & rtfTrans(row("unote").ToString) & vbCrLf &
                               vbCrLf & "Process Engineering Note: " & rtfTrans(row("Bnote").ToString) & vbCrLf &
                               vbCrLf & "Testing Engineering Note: " & rtfTrans(row("enote").ToString) & vbCrLf &
                               vbCrLf & "Quality Note: " & rtfTrans(row("nnote").ToString) & vbCrLf &
                               vbCrLf & "Production Note: " & rtfTrans(row("pnote").ToString) & vbCrLf &
                               vbCrLf & "Time & Methods Note: " & rtfTrans(row("qnote").ToString) & vbCrLf &
                               vbCrLf & "Admin Note: " & rtfTrans(row("anote").ToString) & vbCrLf &
                               vbCrLf & "Environment And Safety Note: " & rtfTrans(row("Snote").ToString) & vbCrLf &
                               vbCrLf & "For all details please download the ECR from server SrvDoc. ", "ECR Approval Notification:   " & " " & row("description").ToString, "SS" & row("number").ToString, False, fileOpen) Then
                        Dim sql As String = "UPDATE `" & DBName & "`.`ecr` SET `approve` = '" & "System" & "[" & date_to_string(Now) & "]" & "' WHERE `ecr`.`approve` ='' and `ecr`.`number` = '" & i & "' ;"
                        cmd = New MySqlCommand(sql, MySqlconnection)
                        cmd.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                    ComunicationLog("0052") 'db operation error
                End Try
            End If
        Next
    End Sub

    Sub ecrDocSign()
        Dim refresh = True
        Dim RowSearchEcr As DataRow() = tblEcr.Select("docInvalid = 'NO'", "number")
        For Each row In RowSearchEcr
            Dim i As Integer
            i = Int(row("number").ToString)

            If row("sign").ToString = "" And InStr(row("Rsign").ToString & row("Usign").ToString & row("Lsign").ToString & row("Bsign").ToString & row("Esign").ToString & row("Nsign").ToString & row("Psign").ToString & row("Qsign").ToString & row("Asign").ToString & row("Ssign").ToString, "APPROVED", CompareMethod.Text) <= 0 And InStr(row("Rsign").ToString & row("Lsign").ToString & row("Usign").ToString & row("Bsign").ToString & row("Esign").ToString & row("Nsign").ToString & row("Psign").ToString & row("Qsign").ToString & row("asign").ToString & row("ssign").ToString, "CHECKED", CompareMethod.Text) <= 0 And readDocSign(Int(row("iddoc").ToString), refresh) = "" Then
                Try
                    Dim fileOpen As Object = downloadFileWinPath(ParameterTable("plant") & "R_PRO_ECR_" & row("DESCRIPTION").ToString, ParameterTable("plant") & "R/" & ParameterTable("plant") & "R_PRO_ECR/")
                    If mailSender("ECR_SignTo", "ECR_SignCopy", "Automatic SrvDoc Message:" & vbCrLf &
                               vbCrLf & row("description").ToString & " -- > (Result: Signed, Released & Implemented) " &
                               vbCrLf & "Closed Data : " & row("date").ToString & "( yyyy/mm/dd )" & vbCrLf &
                               vbCrLf & vbCrLf & "R&D Note: " & rtfTrans(row("rnote").ToString) & vbCrLf &
                               vbCrLf & "Logistic Note: " & rtfTrans(row("lnote").ToString) & vbCrLf &
                               vbCrLf & "Purchasing Note: " & rtfTrans(row("unote").ToString) & vbCrLf &
                               vbCrLf & "Process Engineering Note: " & rtfTrans(row("Bnote").ToString) & vbCrLf &
                               vbCrLf & "Testing Engineering Note: " & rtfTrans(row("enote").ToString) & vbCrLf &
                               vbCrLf & "Quality Note: " & rtfTrans(row("nnote").ToString) & vbCrLf &
                               vbCrLf & "Production Note: " & rtfTrans(row("pnote").ToString) & vbCrLf &
                               vbCrLf & "Time & Methods Note: " & rtfTrans(row("qnote").ToString) & vbCrLf &
                               vbCrLf & "Administration Note: " & rtfTrans(row("anote").ToString) & vbCrLf &
                               vbCrLf & "Environment And Safety Note: " & rtfTrans(row("snote").ToString) & vbCrLf &
                               vbCrLf & "For all details please download ECR from server SrvDoc. ", "ECR Sign Notification:   " & " " & row("description").ToString, "SS" & row("number").ToString, False, fileOpen) Then
                        Dim sql As String = "UPDATE `" & DBName & "`.`ecr` SET `sign` = '" & "System" & "[" & date_to_string(Now) & "]" & "' WHERE `ecr`.`sign` ='' and `ecr`.`number` = '" & i & "' ;"
                        cmd = New MySqlCommand(sql, MySqlconnection)
                        cmd.ExecuteNonQuery()
                        sql = "UPDATE `" & DBName & "`.`doc` SET `sign` = '" & "System" & "[" & date_to_string(Now) & "]" & "' WHERE `doc`.`sign` ='' and `doc`.`id` = '" & row("iddoc").ToString & "' ;"
                        cmd = New MySqlCommand(sql, MySqlconnection)
                        cmd.ExecuteNonQuery()
                    End If

                Catch ex As Exception
                    ComunicationLog("0052") 'db operation failed
                End Try
            End If
            refresh = False
        Next
    End Sub

    Sub TCRMailScheduler()
        tblDoc.Clear()
        DsDoc.Clear()
        AdapterDoc.Fill(DsDoc, "doc")
        tblDoc = DsDoc.Tables("doc")
        Dim RowSearchDoc As DataRow() = tblDoc.Select("sign = '' and HEADER='" & ParameterTable("plant") & "R_PRO_TCR'")
        For Each row In RowSearchDoc
            Dim oi As String = Trim(Mid(row("filename").ToString, 1, InStr(row("filename").ToString, "-") - 1))
            Dim fileOpen As Object = downloadFileWinPath(ParameterTable("plant") & "R_PRO_TCR_" & row("filename").ToString & "_" & row("rev").ToString & "." & row("extension").ToString, ParameterTable("plant") & "R/" & ParameterTable("plant") & "R_PRO_TCR/")
            Try
                If mailSender("STATUS" & "_SignTo", "STATUS" & "_SignCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf &
                               "Please CHECK the TCR : " & " " & row("filename").ToString & " " & vbCrLf & vbCrLf & "Best Regard", "TCR Sign Notification  " & " " &
                               row("filename").ToString, "T_" & oi, False, fileOpen) Then
                    Dim sql As String = "UPDATE `" & DBName & "`.`doc` SET `sign` = 'System[" & date_to_string(Now) & "]' WHERE `doc`.`id` = " & row("id").ToString & " ;"
                    cmd = New MySqlCommand(sql, MySqlconnection)
                    cmd.ExecuteNonQuery()
                Else
                    'MsgBox("Error sending email for TCR!")
                End If

            Catch ex As Exception
                ComunicationLog("5050") ' Mysql update query error 
            End Try
        Next
    End Sub

    Sub DocMailScheduler()
        Dim listFile = ""
        tblDoc.Clear()
        DsDoc.Clear()
        AdapterDoc.Fill(DsDoc, "doc")
        tblDoc = DsDoc.Tables("doc")

        Dim sql As String
        Dim RowSearchDoc = From p In tblDoc.Rows
                           Where (p("header") <> (ParameterTable("plant") & "R_PRO_ECR")) And ((p("notification") = "" And p("sign") = "") Or (p("notification") = "SENT" And p("sign") = "" And (DateTime.Now.Date - DateTime.ParseExact(p("editor").Substring(p("editor").IndexOf("[") + 1, p("editor").LastIndexOf("]") - p("editor").IndexOf("[") - 1), "d/M/yyyy", CultureInfo.CurrentCulture).Date).TotalDays > 2))
                           Select p
        For Each row In RowSearchDoc
            listFile = listFile & " " & vbCrLf & row("header").ToString & "_" & row("FileName").ToString & "_" & row("rev").ToString & "." & row("Extension").ToString & " " & vbCrLf
        Next
        Try
            MailSent = False
            If listFile <> "" Then
                mailSender("STATUS" & "_SignTo", "STATUS" & "_SignCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf &
                           vbCrLf & "Please CHECK the new document/revision in the server : " & " " & vbCrLf & vbCrLf & listFile & vbCrLf & vbCrLf & "Best Regard", "File changes notification  " &
                           date_to_string(Now), date_to_string(Now), True)
            End If
        Catch ex As Exception
            ComunicationLog("5050") ' Mysql update query error 
        End Try

        For Each row In RowSearchDoc
            Try
                sql = "UPDATE `" & DBName & "`.`doc` SET `notification` = 'SENT' WHERE `notification` = '' and sign ='';"
                If MailSent = True Then
                    cmd = New MySqlCommand(sql, MySqlconnection)
                    cmd.ExecuteNonQuery()
                End If
            Catch ex As Exception
                ComunicationLog("5050") ' Mysql update query error 
            End Try
        Next

    End Sub

    Sub StatusMailScheduler()
        tblProd.Clear()
        DsProd.Clear()

        AdapterProd.Fill(DsProd, "product")
        tblProd = DsProd.Tables("product")

        Dim RowSearchProduct As DataRow(), sql As String
        RowSearchProduct = tblProd.Select("")
        For Each row In RowSearchProduct
            Dim oi As String = Replace(row("openissue").ToString, "];", "]" & vbCrLf)
            If oi = "" Then oi = "No Open Issue"

            If (row("Status").ToString = "MPA_APPROVED" Or row("Status").ToString = "MPA_STOPPED") And row("mail").ToString <> "SENT" Then
                Try
                    sql = "UPDATE `" & DBName & "`.`product` SET `mail` = 'SENT' WHERE `product`.`BitronPN` = '" & row("BITRONPN").ToString & "' ;"
                    cmd = New MySqlCommand(sql, MySqlconnection)
                    cmd.ExecuteNonQuery()
                    If mailSender("STATUS" & "_SignTo", "STATUS" & "_SignCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf &
                               "Please CHECK the Status of Product : " & " " & row("bitronpn").ToString & " " & row("name").ToString & vbCrLf &
                               vbCrLf & "Open Issue:" & vbCrLf & oi & vbCrLf & vbCrLf & "Best Regard", "Product Status Notification " & row("STATUS").ToString & " " &
                               row("bitronpn").ToString & " " & row("name").ToString, "S_" & row("bitronpn").ToString, False) Then
                        sql = "UPDATE `" & DBName & "`.`product` SET `mail` = 'SENT' WHERE `product`.`BitronPN` = '" & row("BITRONPN").ToString & "' ;"
                        cmd = New MySqlCommand(sql, MySqlconnection)
                        cmd.ExecuteNonQuery()
                    End If

                Catch ex As Exception
                    ComunicationLog("5050") ' Mysql update query error 
                End Try

                oi = Replace(row("openissue").ToString, "];", "]" & vbCrLf)
                If oi = "" Then oi = "No Open Issue at this moment"
            End If

            For Each SS In dep
                If prevStatus(SS) = row("Status").ToString Or (row("Status").ToString = "MPA_STOPPED" And SS = "N") Then
                    Try
                        mailSender("STATUS_" & SS & "_SignTo", "STATUS_" & SS & "_SignCopy", "Automatic SrvDoc Message:" & vbCrLf & vbCrLf &
                                   "Please Update the Status of Product : " & " " & row("bitronpn").ToString & " " & row("name").ToString & vbCrLf &
                                   vbCrLf & "Current Status:  " & row("Status").ToString & vbCrLf &
                                   vbCrLf & "Open Issue:" & vbCrLf & vbCrLf & oi & vbCrLf & vbCrLf & "Best Regard", "Product Status Update Request " & " " &
                                   row("bitronpn").ToString & " " & row("name").ToString, SS & "_" & row("bitronpn").ToString)
                    Catch ex As Exception
                        ComunicationLog("5050") ' Mysql update query error 
                    End Try
                End If
            Next
        Next
    End Sub

    Function prevStatus(ByVal dep As String) As String
        If dep = "U" Then prevStatus = "R&D_APPROVED"
        If dep = "L" Then prevStatus = "PURCHASING_APPROVED"
        If dep = "C" Then prevStatus = "LOGISTIC_APPROVED"
        If dep = "B" Then prevStatus = "CUSTOMER_APPROVED"
        If dep = "E" Then prevStatus = "PROCESS_ENG_APPROVED"
        If dep = "P" Then prevStatus = "TESTING_ENG_APPROVED"
        If dep = "Q" Then prevStatus = "PRODUCTION_APPROVED"
        If dep = "F" Then prevStatus = "TIME&METHODS_APPROVED"
        If dep = "N" Then prevStatus = "FINANCIAL_APPROVED"
        If dep = "S" Then prevStatus = "ENVIRONMENT_AND_SAFETY"
    End Function

End Class
