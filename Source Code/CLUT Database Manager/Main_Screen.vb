Imports System.IO
Imports System.Net.Mail
Imports System.Web.Mail
Imports System.Text


Public Class Main_Screen

    Private busyworking As Boolean = False
    Private cancelled As Boolean = False
    Private AutoUpdate As Boolean = False
    Dim shownminimizetip As Boolean = False

    Private CLUTDatabase As String = ""
    Private ScheduledTime As String = ""
    Private scheduledcheck As Date

    Private mailserver1 As String = ""
    Private mailserver1port As String = ""
    Private mailserver2 As String = ""
    Private mailserver2port As String = ""
    Private webmasteraddress As String = ""
    Private webmasterdisplay As String = ""
    Private webroot As String = ""
    Private webroottranslate As String = ""
    Private LastReport As Date


    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.ToString
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
            StatusLabel.Text = "Error Reported"
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error." & vbCrLf & vbCrLf & exc.ToString, MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Handler(ByVal message As String)
        Try
            Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            dir = Nothing
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & message)
            filewriter.WriteLine("")
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
            StatusLabel.Text = "Activity Logged"
        Catch ex As Exception
            Error_Handler(ex, "Activity Handler")
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LastReport = New Date(Now.Year, Now.Month, Now.Day + 1, 0, 30, 15, 0, DateTimeKind.Local)
            Control.CheckForIllegalCrossThreadCalls = False
            Me.Text = My.Application.Info.ProductName & " (" & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
            NotifyIcon1.BalloonTipText = "You have chosen to hide " & My.Application.Info.ProductName & ". To bring it back up, simply click here."
            NotifyIcon1.BalloonTipTitle = My.Application.Info.ProductName
            NotifyIcon1.Text = "Click to bring up " & My.Application.Info.ProductName
            ProgressBar1.Value = 0
            loadSettings()
            StatusLabel.Text = "Application Loaded"
            scheduledcheck = New Date(Now.Year, Now.Month, Now.Day, Integer.Parse(Label6.Text.Substring(0, 2)), Integer.Parse(Label6.Text.Substring(3, 2)), Integer.Parse(Label6.Text.Substring(6, 2)), 0)
            Timer1.Enabled = True
            Timer1.Start()
            Control_Enabler(True)
            SendNotificationEmail("Startup")
            Activity_Handler("Application Started")
        Catch ex As Exception
            Error_Handler(ex, "Application Loading")
        End Try
    End Sub

    Private Sub loadSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then
                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)
                        If lineread.StartsWith("CLUTDatabase=") Then
                            If My.Computer.FileSystem.FileExists(variablevalue) = True Then
                                CLUTDatabase = variablevalue.Trim
                            End If
                        End If
                        If lineread.StartsWith("ScheduledTime=") Then
                            ScheduledTime = variablevalue.Trim
                        End If
                        If lineread.StartsWith("mailserver1=") Then
                            mailserver1 = variablevalue
                        End If
                        If lineread.StartsWith("mailserver1port=") Then
                            mailserver1port = variablevalue
                        End If
                        If lineread.StartsWith("mailserver2=") Then
                            mailserver2 = variablevalue
                        End If
                        If lineread.StartsWith("mailserver2port=") Then
                            mailserver2port = variablevalue
                        End If
                        If lineread.StartsWith("webmasteraddress=") Then
                            webmasteraddress = variablevalue
                        End If
                        If lineread.StartsWith("webmasterdisplay=") Then
                            webmasterdisplay = variablevalue
                        End If
                        If lineread.StartsWith("webroot=") Then
                            webroot = variablevalue
                        End If
                        If lineread.StartsWith("webroottranslate=") Then
                            webroottranslate = variablevalue
                        End If
                    End If
                End While
                reader.Close()
                reader = Nothing
            End If

            Label5.Text = CLUTDatabase
            If ScheduledTime.Length < 1 Then
                ScheduledTime = "23:45:00"
            End If
            Label6.Text = ScheduledTime
            If webroottranslate.Length < 1 Then
                webroottranslate = "http://www.commerce.uct.ac.za"
            End If
            If webroot.Length < 1 Then
                webroot = "C:\Inetpub\wwwroot"
            End If
            If webmasterdisplay.Length < 1 Then
                webmasterdisplay = "Commerce Webmaster"
            End If
            If webmasteraddress.Length < 1 Then
                webmasteraddress = "com-webmaster@uct.ac.za"
            End If
            If mailserver2.Length < 1 Then
                mailserver2 = "obe1.com.uct.ac.za"
            End If
            If mailserver2port.Length < 1 Then
                mailserver2port = "25"
            End If
            If mailserver1.Length < 1 Then
                mailserver1 = "mail.uct.ac.za"
            End If
            If mailserver1port.Length < 1 Then
                mailserver1port = "25"
            End If
            StatusLabel.Text = "Application Settings Loaded"
            Activity_Handler("Application Settings Loaded")
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            Dim writer As StreamWriter = New StreamWriter(configfile, False)
            writer.WriteLine("CLUTDatabase=" & CLUTDatabase)
            writer.WriteLine("ScheduledTime=" & ScheduledTime)
            writer.WriteLine("mailserver1=" & mailserver1)
            writer.WriteLine("mailserver1port=" & mailserver1port)
            writer.WriteLine("mailserver2=" & mailserver2)
            writer.WriteLine("mailserver2port=" & mailserver2port)
            writer.WriteLine("webmasteraddress=" & webmasteraddress)
            writer.WriteLine("webmasterdisplay=" & webmasterdisplay)
            writer.WriteLine("webroot=" & webroot)
            writer.WriteLine("webroottranslate=" & webroottranslate)
            writer.Flush()
            writer.Close()
            writer = Nothing
            StatusLabel.Text = "Application Settings Saved"
            Activity_Handler("Application Settings Saved")
        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub

    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            StatusLabel.Text = "Application Shutting Down"
            SendNotificationEmail("Shutdown")
            SaveSettings()
            If AutoUpdate = True Then
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")) = True Then
                    Dim startinfo As ProcessStartInfo = New ProcessStartInfo
                    startinfo.FileName = (Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")
                    startinfo.Arguments = "force"
                    startinfo.CreateNoWindow = False
                    Process.Start(startinfo)
                End If
            End If
            Activity_Handler("Application Shut Down")
        Catch ex As Exception
            Error_Handler(ex, "Closing Application")
        End Try
    End Sub
    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClicked
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub


    Private Sub NotifyIcon1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseClick
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub


    Private Sub NotifyIcon1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub

    Private Sub Main_Screen_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try
            If Me.WindowState = FormWindowState.Minimized Then
                Me.ShowInTaskbar = False
                NotifyIcon1.Visible = True
                If shownminimizetip = False Then
                    NotifyIcon1.ShowBalloonTip(1)
                    shownminimizetip = True
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Change Window State")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem1.Click
        Try
            HelpBox1.ShowDialog()
            StatusLabel.Text = "Help Dialog Viewed"
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub

    Private Sub AutoUpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoUpdateToolStripMenuItem.Click
        Try
            StatusLabel.Text = "AutoUpdate Requested"
            AutoUpdate = True
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "AutoUpdate")
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem1.Click
        Try
            AboutBox1.ShowDialog()
            StatusLabel.Text = "About Dialog Viewed"
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub Control_Enabler(ByVal IsEnabled As Boolean)
        Try
            Select Case IsEnabled
                Case True
                    Button1.Enabled = True
                    Button2.Enabled = False
                    Button3.Enabled = True
                    MenuStrip1.Enabled = True
                    Me.ControlBox = True
                    ProgressBar1.Enabled = False
                Case False
                    Button1.Enabled = False
                    Button2.Enabled = True
                    Button3.Enabled = False
                    MenuStrip1.Enabled = False
                    Me.ControlBox = False
                    ProgressBar1.Enabled = True
            End Select
            StatusLabel.Text = "Control Enabler Run"
        Catch ex As Exception
            Error_Handler(ex, "Control Enabler")
        End Try
    End Sub





    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            BackgroundWorker1.ReportProgress(0)
            Dim dbcopyrequired As Boolean = False
            Dim lookingfordb As String = ""
            Dim testdate As Date
            Dim formattedmonth As String
            Dim formattednowmonth As String
            formattednowmonth = Now.Month
            While formattednowmonth.Length < 2
                formattednowmonth = "0" & formattednowmonth
            End While
            If cancelled = False Then
                Activity_Handler("Database Check Requested")
                StatusLabel.Text = "Checking if Database Copy is Required"
                Dim finfo As FileInfo = New FileInfo(CLUTDatabase)
                Dim dinfo As DirectoryInfo = New DirectoryInfo(finfo.DirectoryName)
                testdate = New Date(Now.Year, Now.Month, 1)
                testdate = testdate.AddMonths(-1)

                formattedmonth = testdate.Month
                While formattedmonth.Length < 2
                    formattedmonth = "0" & formattedmonth
                End While
                lookingfordb = "Lab_Usage_Tracker_" & testdate.Year & formattedmonth & ".mdb"
                lookingfordb = (dinfo.FullName & "\" & lookingfordb).Replace("\\", "\")
                If My.Computer.FileSystem.FileExists(lookingfordb) = False Then
                    dbcopyrequired = True
                End If
                dinfo = Nothing
                finfo = Nothing
            End If
            If dbcopyrequired = True Then
                If cancelled = False Then
                    BackgroundWorker1.ReportProgress(25)
                    StatusLabel.Text = "Duplicating Database"
                    Activity_Handler("Database Duplication Initiated")
                    My.Computer.FileSystem.CopyFile(CLUTDatabase, lookingfordb, True)
                    StatusLabel.Text = "Database Duplication Complete"
                    Try
                        StatusLabel.Text = "Cleaning Up Database Records: " & "Lab_Usage_Tracker_" & testdate.Year & formattedmonth & ".mdb"
                        Dim dbconnection As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & lookingfordb & """")
                        dbconnection.Open()
                        Dim dbsql As OleDb.OleDbCommand = New OleDb.OleDbCommand
                        dbsql.Connection = dbconnection
                        dbsql.CommandText = "delete * from Lab_Usage_Tracker where Time_Stamp not like '" & testdate.Year & formattedmonth & "%'"
                        dbsql.ExecuteNonQuery()
                        dbsql.Dispose()
                        dbsql = Nothing
                        dbconnection.Close()
                        dbconnection.Dispose()
                        dbconnection = Nothing
                        StatusLabel.Text = "Database Records Cleaned: " & "Lab_Usage_Tracker_" & testdate.Year & formattedmonth & ".mdb"
                    Catch ex As Exception
                        Error_Handler(ex, "Error in Cleaning Up Database Records: " & "Lab_Usage_Tracker_" & testdate.Year & formattedmonth & ".mdb")
                    End Try
                    Try
                        StatusLabel.Text = "Cleaning Up Database Records: " & "Lab_Usage_Tracker.mdb"
                        Dim dbconnection As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & CLUTDatabase & """")
                        dbconnection.Open()
                        Dim dbsql As OleDb.OleDbCommand = New OleDb.OleDbCommand
                        dbsql.Connection = dbconnection
                        dbsql.CommandText = "delete * from Lab_Usage_Tracker where Time_Stamp not like '" & Now.Year & formattednowmonth & "%'"
                        dbsql.ExecuteNonQuery()
                        dbsql.Dispose()
                        dbsql = Nothing
                        dbconnection.Close()
                        dbconnection.Dispose()
                        dbconnection = Nothing
                        StatusLabel.Text = "Database Records Cleaned: " & "Lab_Usage_Tracker.mdb"
                    Catch ex As Exception
                        Error_Handler(ex, "Error in Cleaning Up Database Records: " & "Lab_Usage_Tracker_" & testdate.Year & formattedmonth & ".mdb")
                    End Try
                End If
            End If
            BackgroundWorker1.ReportProgress(100)
        Catch ex As Exception
            Error_Handler(ex, "Copy Operation")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Control_Enabler(True)
            If cancelled = False And e.Cancelled = False And e.Error Is Nothing Then
                StatusLabel.Text = "Check Complete"
                Activity_Handler("Database Check Complete")
            Else
                If Not e.Error Is Nothing Then
                    StatusLabel.Text = "Check Failed"
                    Activity_Handler("Database Check Failed")
                Else
                    StatusLabel.Text = "Check Cancelled"
                    Activity_Handler("Database Check Cancelled")
                End If
            End If
            busyworking = False
        Catch ex As Exception
            Error_Handler(ex, "Check Complete")
        End Try
    End Sub


    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            If e.ProgressPercentage < 0 Then
                ProgressBar1.Value = 0
            Else
                ProgressBar1.Value = e.ProgressPercentage
            End If
        Catch ex As Exception
            Error_Handler(ex, "Progress Changed (" & e.ProgressPercentage & ")")
        End Try

    End Sub

    Private Sub runworker()
        Try
            If busyworking = False Then
                If My.Computer.FileSystem.FileExists(CLUTDatabase) = True Then
                    busyworking = True
                    cancelled = False
                    Control_Enabler(False)
                    ProgressBar1.Value = 0
                    StatusLabel.Text = "Initializing Check"
                    BackgroundWorker1.RunWorkerAsync()
                Else
                    StatusLabel.Text = "Database file cannot be located"
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Run Worker")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            runworker()
        Catch ex As Exception
            Error_Handler(ex, "Start Button Click")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            BackgroundWorker1.CancelAsync()
            cancelled = True
        Catch ex As Exception
            Error_Handler(ex, "Cancel Button Click")
        End Try
    End Sub




    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            If busyworking = False Then
                Dim TimeSpan1 As TimeSpan
                TimeSpan1 = scheduledcheck.Subtract(Now)
                If TimeSpan1.Seconds <= 0 Then
                    scheduledcheck = scheduledcheck.AddDays(1)
                    TimeSpan1 = scheduledcheck.Subtract(Now)
                End If

                StatusLabel.Text = "Awaiting Scheduled Check"

                Dim breakhour, breakminute, breaksecond As String
                breakhour = TimeSpan1.Hours.ToString
                While breakhour.Length < 2
                    breakhour = "0" & breakhour
                End While
                breakminute = TimeSpan1.Minutes.ToString
                While breakminute.Length < 2
                    breakminute = "0" & breakminute
                End While
                breaksecond = TimeSpan1.Seconds.ToString
                While breaksecond.Length < 2
                    breaksecond = "0" & breaksecond
                End While
                Label8.Text = breakhour & ":" & breakminute & ":" & breaksecond

                If Label8.Text = "00:00:00" Or TimeSpan1.Seconds < 0 Then
                    runworker()
                End If

                If TimeSpan1.Seconds <= 0 Then
                    scheduledcheck = scheduledcheck.AddDays(1)
                End If
            End If
            If busyworking = False Then
                Dim dt As Date = New Date(Now.Year, Now.Month, Now.Day, 2, 5, 15, 0, DateTimeKind.Local)
                If dt > LastReport Then
                    Timer1.Stop()
                    Activity_Handler("Sending Daily Notification")
                    LastReport = dt
                    Send_Report(Now, Format(LastReport, "yyyyMMdd"))
                    Timer1.Start()
                End If
            End If
        Catch ex As Exception
            Timer1.Stop()
            Error_Handler(ex, "Timer 1 Tick")
            Timer1.Start()
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        busyworking = True
        Try
            StatusLabel.Text = "Selecting CLUT Database"
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                If My.Computer.FileSystem.FileExists(OpenFileDialog1.FileName) = True Then
                    CLUTDatabase = OpenFileDialog1.FileName
                    Label5.Text = CLUTDatabase
                    StatusLabel.Text = "New CLUT Database Selected"
                End If
            Else
                StatusLabel.Text = "CLUT Database Selection Cancelled"
            End If
        Catch ex As Exception
            Error_Handler(ex, "Select CLUT database")
        End Try
        busyworking = False
    End Sub

    Private Sub SendNotificationEmail(ByVal StartOrClose As String)
        Try
            Dim obj As SmtpClient
            If mailserver1port.Length > 0 Then
                obj = New SmtpClient(mailserver1, mailserver1port)
            Else
                obj = New SmtpClient(mailserver1)
            End If

            Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

            If StartOrClose = "Startup" Then
                msg.Subject = My.Application.Info.ProductName & ": Application Startup"
                StatusLabel.Text = "Sending Startup Notification"
            Else
                msg.Subject = My.Application.Info.ProductName & ": Application Shutdown"
                StatusLabel.Text = "Sending Shutdown Notification"
            End If

            Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
            msg.From = fromaddress
            msg.ReplyTo = fromaddress
            msg.To.Add(fromaddress)

            msg.IsBodyHtml = False

            Dim body As String
            If StartOrClose = "Startup" Then
                body = "This is just a notification message to inform you that " & My.Application.Info.ProductName & " has been successfully started up."
            Else
                body = "This is just a notification message to inform you that " & My.Application.Info.ProductName & " has been shutdown."
            End If

            body = body & vbCrLf & vbCrLf & "******************************" & vbCrLf & vbCrLf & "This is an auto-generated email submitted from " & My.Application.Info.ProductName & " at " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & ", running on:"
            body = body & vbCrLf & vbCrLf & "Machine Name: " + Environment.MachineName
            body = body & vbCrLf & "OS Version: " & Environment.OSVersion.ToString()
            body = body & vbCrLf & "User Name: " + Environment.UserName
            msg.Body = body

            obj.DeliveryMethod = SmtpDeliveryMethod.Network
            obj.EnableSsl = False
            obj.UseDefaultCredentials = True


            obj.Send(msg)
            obj = Nothing

            If StartOrClose = "Startup" Then
                StatusLabel.Text = "Startup Notification Sent"
                Activity_Handler("Startup Notification Sent")
            Else
                StatusLabel.Text = "Shutdown Notification Sent"
                Activity_Handler("Shutdown Notification Sent")
            End If

        Catch ex As Exception
            Error_Handler(ex, "Send Startup/Shutdown Email")
        End Try
    End Sub

    Private Function Send_Report(ByVal dt As Date, ByVal FileNamePrefix As String) As Boolean
        Dim objMail As System.Web.Mail.MailMessage = New System.Web.Mail.MailMessage
        Try
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & FileNamePrefix & "_Error_Log.txt") = True Then
                My.Computer.FileSystem.CopyFile((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & FileNamePrefix & "_Error_Log.txt", (Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & FileNamePrefix & "_Error_Log_Copy.txt", True)
            End If
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log.txt") = True Then
                My.Computer.FileSystem.CopyFile((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log.txt", (Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log_Copy.txt", True)
            End If
            Try
                objMail.BodyFormat = MailFormat.Text
                objMail.From = webmasteraddress
                objMail.To = webmasteraddress
                objMail.Subject = My.Application.Info.ProductName & ": Daily Report"
                Dim body As String
                body = "Any activity and error logs for " & Format(dt, "dd MM yyyy") & " have been attached with this report."
                body = body & vbCrLf & vbCrLf & "******************************" & vbCrLf & vbCrLf & "This is an auto-generated email submitted from " & My.Application.Info.ProductName & " at " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & ", running on:"
                body = body & vbCrLf & vbCrLf & "Machine Name: " + Environment.MachineName
                body = body & vbCrLf & "OS Version: " & Environment.OSVersion.ToString()
                body = body & vbCrLf & "User Name: " + Environment.UserName
                objMail.Body = body

                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log_Copy.txt") = True Then
                    Dim objAttach As MailAttachment = New MailAttachment((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log_Copy.txt")
                    objMail.Attachments.Add(objAttach)
                    objAttach = Nothing
                End If
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & FileNamePrefix & "_Error_Log_Copy.txt") = True Then
                    Dim objAttach As MailAttachment = New MailAttachment((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & FileNamePrefix & "_Error_Log_Copy.txt")
                    objMail.Attachments.Add(objAttach)
                    objAttach = Nothing
                End If

                SmtpMail.SmtpServer = mailserver1
                SmtpMail.Send(objMail)
                objMail = Nothing
                Send_Report = True
            Catch ex As Exception
                Send_Report = False
                objMail.Attachments.Clear()
                objMail = Nothing
                Error_Handler(ex, "Send Report Mail: " & webmasteraddress & " (Send Report)")
            End Try
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log_Copy.txt") = True Then
                My.Computer.FileSystem.DeleteFile((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log_Copy.txt")
            End If
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & FileNamePrefix & "_Error_Log_Copy.txt") = True Then
                My.Computer.FileSystem.DeleteFile((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & FileNamePrefix & "_Error_Log_Copy.txt")
            End If
            Activity_Handler("Report Mail Sent")
        Catch ex As Exception
            Error_Handler(ex, "Send Report Mail: " & webmasteraddress & " (Send Report)")
        End Try
    End Function
End Class
