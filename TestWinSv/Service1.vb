Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Text

Public Class Service1
    Private Count As Integer = 0
    Public con As String = "Data Source=192.168.0.15;UID=parinya;Password=Passw0rd@1;Database=CSSC;Connect Timeout=600;Max Pool Size=400;"
    Private SV As String = "Data Source=192.168.0.15;UID=parinya;Password=Passw0rd@1;Database=CSSC;Connect Timeout=600;Max Pool Size=400;"
    Dim NameApprove As StringBuilder
    Dim MailApprove As StringBuilder
    Private STA As Boolean = False
    Private CountClear As Integer
    Public Sub OnDebug()
        Me.OnStart(Nothing)
    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Dim startPath As String = AppDomain.CurrentDomain.BaseDirectory + "Log.txt"
        Dim lines() As String = {"Start time : XX : " + DateTime.Now.ToString()}
        System.IO.File.AppendAllLines(startPath, lines)

    End Sub
    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "Log.txt"
        Dim lines() As String = {"Stop time : " + DateTime.Now.ToString()}
        System.IO.File.AppendAllLines(strPath, lines)
    End Sub
    Private Sub Timer1_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        Count += 1
        Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "Log.txt"
        Dim lines() As String = {Count & "..calling time :" + DateTime.Now.ToString()}
        System.IO.File.AppendAllLines(strPath, lines)
        If Count = 50 Then
            'SELECTDATA()
            ' SENDMAILREQUESTAPPROVEOTHERDEPARTMENT("DOCNO", "ITEM", "ITEMNAME", "REQUEST", "PARINYA")
            STA = True
            CHECKSENDMAIL()
        End If
        ' Timer1.Enabled = False
    End Sub
    'Private Sub SELECTDATA()
    '    Dim tb As New DataTable
    '    Using conr As New SqlConnection(con)
    '        conr.Open()
    '        Using cmd As New SqlCommand("SELECT TOP 10 * FROM YSSMODIFYREQUEST ORDER BY ID ASC", conr)
    '            Using dtr As SqlDataReader = cmd.ExecuteReader
    '                If dtr.HasRows Then
    '                    Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
    '                    tb.Load(dtr)
    '                    For i As Integer = 0 To tb.Rows.Count - 1
    '                        Dim lines() As String = {"DETAIL : " & Count & ":" & tb.Rows(i)(0) & " : " & tb.Rows(i)(1) & " : " & tb.Rows(i)(2) & " : " & tb.Rows(i)(3)}
    '                        System.IO.File.AppendAllLines(strPath, lines)
    '                    Next
    '                End If
    '            End Using
    '        End Using
    '    End Using
    'End Sub
    Public Function SendEmail(ByVal sj As String, ByVal mailTo As String, ByVal mailFrom As String, mailBody As StringBuilder, mailCC As String) As String
        ssl()
        Dim mail As New MailMessage
        mail.Subject = sj
        mail.To.Add(mailTo)
        mail.From = New MailAddress(mailFrom)
        If ValidateEmail(mailCC) = True Then
            mail.CC.Add(mailCC)
        End If
        mail.Body = mailBody.ToString
        mail.IsBodyHtml = True
        Dim smtp As New SmtpClient("192.168.0.251", 587)
        smtp.Credentials = New System.Net.NetworkCredential("IT@yss.com", "1234")
        smtp.EnableSsl = True
        smtp.UseDefaultCredentials = True
        smtp.DeliveryMethod = SmtpDeliveryMethod.Network
        Try
            smtp.Send(mail)
            Return "Send complete."
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function ValidateEmail(emailaddress As String) As Boolean
        Dim smail() As String = Split(emailaddress, ",")
        If smail.Count > 0 Then
            Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
            Dim emailAddressMatch As Match = Regex.Match(smail(0), pattern)
            If emailAddressMatch.Success Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function
    Public Function validateCerts(ByVal sender As Object, ByVal certificate As X509Certificate, _
                           ByVal chain As X509Chain, ByVal sspPolicyErrors As SslPolicyErrors) As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Sub ssl()
        Try
            Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf validateCerts
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub SENDMAILREQUESTAPPROVEOTHERDEPARTMENT(docno As String, item As String, itemname As String, fomrType As String, UserName As String)
        Dim sb As New StringBuilder 'ประกาศใช้งานตัวแปร StringBuilder
        sb.AppendLine("Dear " & NameApprove.ToString & "  <br><br>")
        'sb.AppendLine(MailApprove.ToString)
        sb.AppendLine("<br><br>")
        sb.AppendLine("approve document 1 item.<br><br>")
        sb.AppendLine("Doc.No : " & docno & "<br>")
        sb.AppendLine("Item Code : " & item & "<br>")
        sb.AppendLine("Item Name : " & itemname & "<br>")
        sb.AppendLine("Request from : " & UserName & "<br>")
        sb.AppendLine("Detail : Request for approve new item code<br><br>")
        sb.AppendLine("&nbsp;&nbsp;Please go to the REQUEST NEW ITEM program to approve.<br>")
        sb.AppendLine("\\192.168.0.10\Software\ApplicationRequest\setup.exe")
        sb.AppendLine("<br><br>")
        sb.AppendLine("(It is only an automated notification email.)<br><br>")
        sb.AppendLine("Thank you.")
        'Dim mess As String = SendEmail("Request Item System", stMail, Environment.UserName & "@yss.com", sb, Environment.UserName & "@yss.com") 'เรียกใช้งานฟังก์ชั่น [ Subject ,MailTo ,MailFrom ] CheckEmailAddress(S_id
        Dim mess As String = SendEmail("REQUEST NEW ITEM : NOTICE  APPROVED DOC NO : '" & docno & "'", "parinya@yss.com", "IT@yss.com", sb, "parinya@yss.com")
        ' MsgBox(mess) 'แจ้งเตือนส่งเมล์'"sudakan@yss.com"
        Console.WriteLine(mess)
    End Sub
    Private Sub GETUSEREMAILFORAPPROVEREQUEST(ft As String)
        Dim tbApp As New DataTable
        NameApprove = New StringBuilder
        MailApprove = New StringBuilder
        Using con As New SqlConnection(SV)
            con.Open()
            Dim sql As String
            sql = "SELECT distinct a.[NAME],a.EMAIL FROM db_EMPL as a INNER JOIN YSSTB_USEMENU as b ON a.EMPLID = b.EMPLID" & _
" WHERE b.FORMID in (" & ft & ") AND a.DPID != '2200' AND a.EMAIL is not null"
            Using cmd As New SqlCommand(sql, con)
                Using dtr As SqlDataReader = cmd.ExecuteReader
                    If dtr.HasRows Then
                        tbApp.Load(dtr)
                        For i As Integer = 0 To tbApp.Rows.Count - 1
                            NameApprove.Append(tbApp.Rows(i)("NAME").ToString)
                            MailApprove.Append(tbApp.Rows(i)("EMAIL").ToString)
                            If i < (tbApp.Rows.Count - 1) Then
                                NameApprove.Append(",")
                                MailApprove.Append(",")
                            End If
                        Next
                    End If
                End Using
            End Using
        End Using
    End Sub
    Private Sub CHECKSENDMAIL()
        Dim sql As String = "SELECT [ID],[RunNo],[ITEMID],[ITEMNAME],[Request_by],[frm],[Item_type],[SpringItem]" & _
            " FROM [dbo].[YSSREQUESTNEWDOC] WHERE STATUSFORM = 1"
        Dim tb As New DataTable
        Using con As New SqlConnection(SV)
            con.Open()
            Using cmd As New SqlCommand(sql, con)
                Using dtr As SqlDataReader = cmd.ExecuteReader
                    If dtr.HasRows Then
                        tb.Load(dtr)
                        For i As Integer = 0 To tb.Rows.Count - 1
                            Threading.Thread.Sleep(12000)
                            SENDMAIL(tb.Rows(i)(1).ToString, tb.Rows(i)("frm").ToString, tb.Rows(i)("ITEMID").ToString, tb.Rows(i)("ITEMNAME").ToString, tb.Rows(i)("Request_by").ToString, tb.Rows(i)("Item_type").ToString)
                        Next i
                        'aTime.Start()
                        Count = 0
                        Timer1.Enabled = False
                    End If
                End Using
            End Using
        End Using
    End Sub
    Private Sub SENDMAIL(Id As String, idf As String, item As String, itemname As String, Users As String, typebom As String)
        Dim sql As String = "SELECT [RunNo],[PlanningApprove],[PurchaseApprove],[WarehouseApprove],[AccountCostApprove],[AccountFinanceApprove]" & _
            " FROM [CSSC].[dbo].[YSSPROCESSAPPROVE] WHERE RunNo = '" & Id & "'"
        Using con As New SqlConnection(SV)
            con.Open()
            Using cmd As New SqlCommand(sql, con)
                Using dtr As SqlDataReader = cmd.ExecuteReader
                    If dtr.HasRows Then
                        While dtr.Read
                            If dtr.IsDBNull(dtr.GetOrdinal("PlanningApprove") And idf <> "1") Then
                                GETUSEREMAILFORAPPROVEREQUEST("6,7")
                                SENDMAILREQUESTAPPROVEOTHERDEPARTMENT(Id, item, itemname, idf, Users)
                                Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
                                Dim lines() As String = {"Sendmail to PlanningApprove from " & Id & ":" & MailApprove.ToString() & " : Complete"}
                                System.IO.File.AppendAllLines(strPath, lines)
                                Threading.Thread.Sleep(10000)
                              
                                ' Console.WriteLine("Sendmail to PlanningApprove from  '" & Id & "' , " & MailApprove.ToString())
                            End If
                            If dtr.IsDBNull(dtr.GetOrdinal("PurchaseApprove") And idf <> "2") Then
                                GETUSEREMAILFORAPPROVEREQUEST("8,9")
                                SENDMAILREQUESTAPPROVEOTHERDEPARTMENT(Id, item, itemname, idf, Users)
                                Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
                                Dim lines() As String = {"Sendmail to PlanningApprove from " & Id & ":" & MailApprove.ToString() & " : Complete"}
                                System.IO.File.AppendAllLines(strPath, lines)
                                Threading.Thread.Sleep(10000)
                                '  Console.WriteLine("Sendmail ot PurchaseApprove from  '" & Id & "' , " & MailApprove.ToString())
                            End If
                            If dtr.IsDBNull(dtr.GetOrdinal("WarehouseApprove")) Then
                                GETUSEREMAILFORAPPROVEREQUEST("12,13")
                                SENDMAILREQUESTAPPROVEOTHERDEPARTMENT(Id, item, itemname, idf, Users)
                                Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
                                Dim lines() As String = {"Sendmail to PlanningApprove from " & Id & ":" & MailApprove.ToString() & " : Complete"}
                                System.IO.File.AppendAllLines(strPath, lines)
                                Threading.Thread.Sleep(10000)
                                ' Console.WriteLine("Sendmail to WarehouseApprove from  '" & Id & "' , " & MailApprove.ToString())
                            End If
                            If dtr.IsDBNull(dtr.GetOrdinal("AccountCostApprove")) Then
                                GETUSEREMAILFORAPPROVEREQUEST("10,11")
                                SENDMAILREQUESTAPPROVEOTHERDEPARTMENT(Id, item, itemname, idf, Users)
                                Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
                                Dim lines() As String = {"Sendmail to PlanningApprove from " & Id & ":" & MailApprove.ToString() & " : Complete"}
                                System.IO.File.AppendAllLines(strPath, lines)
                                Threading.Thread.Sleep(10000)
                                'Console.WriteLine("Sendmail to AccountCostApprove from  '" & Id & "' , " & MailApprove.ToString())
                            End If
                            If typebom = "BOM" Then
                                If CHECKBOMACTIVE(idf) = False And CHECKROUTEACTIVE(idf) = False Then
                                    GETUSEREMAILFORAPPROVEREQUEST("22,23")
                                    REQUESTCREATEBOM(Id, item, itemname, Users)
                                    Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
                                    Dim lines() As String = {"Sendmail for BOM ROUTE"}
                                    System.IO.File.AppendAllLines(strPath, lines)
                                    Threading.Thread.Sleep(10000)
                                ElseIf CHECKBOMACTIVE(idf) = True And CHECKROUTEACTIVE(idf) = False Then
                                    GETUSEREMAILFORAPPROVEREQUEST("22,23")
                                    REQUESTCREATEBOM(Id, item, itemname, Users)
                                    Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
                                    Dim lines() As String = {"Sendmail for ROUTE"}
                                    System.IO.File.AppendAllLines(strPath, lines)
                                    Threading.Thread.Sleep(10000)
                                ElseIf CHECKBOMACTIVE(idf) = False And CHECKROUTEACTIVE(idf) = True Then
                                    GETUSEREMAILFORAPPROVEREQUEST("22,23")
                                    REQUESTCREATEBOM(Id, item, itemname, Users)
                                    Dim strPath As String = AppDomain.CurrentDomain.BaseDirectory + "LogPro.txt"
                                    Dim lines() As String = {"Sendmail for BOM"}
                                    System.IO.File.AppendAllLines(strPath, lines)
                                    Threading.Thread.Sleep(10000)
                                End If
                                'Console.WriteLine("Sendmail to RD from  '" & Id & "' , " & MailApprove.ToString())
                            End If
                            'If dtr.IsDBNull(dtr.GetOrdinal("AccountFinanceApprove")) And idf = "2" Then
                            '    GETUSEREMAILFORAPPROVEREQUEST("17,19")
                            '    SENDMAILREQUESTAPPROVEOTHERDEPARTMENT(Id, item, itemname, idf, Users)
                            '    Threading.Thread.Sleep(10000)
                            '    Console.WriteLine("Sendmail to AccountFinanceApprove from  '" & Id & "' , " & MailApprove.ToString())
                            'End If
                        End While
                    End If
                End Using
            End Using
        End Using
    End Sub
    Public Sub REQUESTCREATEBOM(docno As String, item As String, itemname As String, UserName As String)
        Dim sb As New StringBuilder 'ประกาศใช้งานตัวแปร StringBuilde
        sb.AppendLine("Dear " & NameApprove.ToString & " <br><br>")
        'sb.AppendLine(MailApprove.ToString)
        sb.AppendLine("<br><br>")
        sb.AppendLine("approve document 1 item.<br><br>")
        sb.AppendLine("Doc.No : " & docno & "<br>")
        sb.AppendLine("Item Code : " & item & "<br>")
        sb.AppendLine("Item Name : " & itemname & "<br>")
        sb.AppendLine("Approve from : " & UserName & "<br>")
        sb.AppendLine("Detail : Request for approve new item code<br><br>")
        sb.AppendLine("&nbsp;&nbsp;Please go to the REQUEST NEW ITEM program to create BOM and ROUTE or Active and Approve.<br>")
        sb.AppendLine("\\192.168.0.10\Software\ApplicationRequest\setup.exe")
        sb.AppendLine("<br><br>")
        sb.AppendLine("(It is only an automated notification email.)<br><br>")
        sb.AppendLine("Thank you.")
        'Dim mess As String = SendEmail("Request Item System", stMail, Environment.UserName & "@yss.com", sb, Environment.UserName & "@yss.com") 'เรียกใช้งานฟังก์ชั่น [ Subject ,MailTo ,MailFrom ] CheckEmailAddress(S_id
        Dim mess As String = SendEmail("REQUEST NEW ITEM : NOTICE  APPROVED DOC NO : '" & docno & "'", "Parinya@yss.com", "IT@yss.com", sb, "parinya@yss.com")
        ' MsgBox(mess) 'แจ้งเตือนส่งเมล์
        Console.WriteLine(mess)
    End Sub
    Public Function CHECKROUTEACTIVE(id As Integer) As Boolean
        Using con As New SqlConnection(SV)
            con.Open()
            Using cmd As New SqlCommand("SELECT * FROM YSSROUTEVERSION WHERE ID = " & id & " AND ACTIVE = 1", con)
                Using dtr As SqlDataReader = cmd.ExecuteReader
                    If dtr.HasRows Then
                        Return True
                    Else
                        Return False
                    End If
                End Using
            End Using
        End Using
    End Function
    Public Function CHECKBOMACTIVE(runid As String) As Boolean
        Using con As New SqlConnection(SV)
            con.Open()
            Dim Tchekc As New DataTable
            Using cmd As New SqlCommand("SELECT * FROM YSSBOMVERSION WHERE RunNo = '" & runid & "' AND ACTIVE = 1", con)
                Using dtr As SqlDataReader = cmd.ExecuteReader
                    If dtr.HasRows Then
                        Tchekc.Load(dtr)
                        For i As Integer = 0 To Tchekc.Rows.Count - 1
                            If Tchekc.Rows(i)("ACTIVE") = 1 Then
                                Return True
                                Exit For
                            End If
                        Next
                    Else
                        Return False
                    End If
                End Using
            End Using
        End Using
    End Function
End Class
