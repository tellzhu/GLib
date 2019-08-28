Imports WinSCP
Imports System.IO
Imports System.Net
Imports System.Net.Mail

Namespace net
    Public Class NetMaster

        ''' <summary>
        ''' 获取本地计算机的IP地址。
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLocalHostIPAddress() As String
            Return GetHostIPAddress(Dns.GetHostName)
        End Function

        Private Shared Function GetHostIPAddress(ByVal HostName As String) As String
            Dim IP As String = Nothing
            Dim addrList As IPAddress() = Dns.GetHostEntry(HostName).AddressList
            For i As Integer = 0 To addrList.Length - 1
                If addrList(i).AddressFamily = Sockets.AddressFamily.InterNetwork Then
                    IP = addrList(i).ToString
                    Exit For
                End If
            Next
            Array.Clear(addrList, 0, addrList.Length)
            addrList = Nothing
            Return IP
        End Function

        ''' <summary>
        ''' 从SFTP服务器下载文件。
        ''' </summary>
        ''' <param name="SFTPServer">SFTP服务器名称。</param>
        ''' <param name="UserName">服务器用户名称。</param>
        ''' <param name="Password">服务器登录密码。</param>
        ''' <param name="KeyFingerprint">服务器密钥指纹</param> 
        ''' <param name="RemoteFiles">服务器上待下载的文件列表所在路径。</param>
        ''' <param name="LocalPath">下载文件后的本地路径。</param>
        Public Shared Sub DownloadSFTPFiles(SFTPServer As String, UserName As String, Password As String, KeyFingerprint As String,
                                        RemoteFiles As String(), LocalPath As String)
            Dim sessionOptions As SessionOptions = New SessionOptions()
            With sessionOptions
                .Protocol = Protocol.Sftp
                .HostName = SFTPServer
                .UserName = UserName
                .Password = Password
                .SshHostKeyFingerprint = KeyFingerprint
            End With

            Using session As New Session
                session.Open(sessionOptions)
                Dim transferOptions As New TransferOptions
                transferOptions.TransferMode = TransferMode.Binary

                Dim transferResult As TransferOperationResult
                For i As Integer = 0 To RemoteFiles.Length - 1
                    transferResult = session.GetFiles(RemoteFiles(i), LocalPath, False, transferOptions)
                    transferResult.Check()
                Next
                session.Close()
            End Using
            sessionOptions = Nothing
        End Sub

        ''' <summary>
        ''' 从FTP服务器下载文件。
        ''' </summary>
        ''' <param name="FtpServer">FTP服务器名称。</param>
        ''' <param name="UserName">FTP用户名称。</param>
        ''' <param name="Password">FTP登录密码。</param>
        ''' <param name="RemotePath">FTP文件所在的服务器路径。</param>
        ''' <param name="LocalPath">下载文件后的本地路径。</param>
        Public Shared Sub DownloadFTPFile(FtpServer As String, UserName As String, Password As String,
                                        RemotePath As String, LocalPath As String)
            If RemotePath.StartsWith("/") Or RemotePath.StartsWith("\") Then
                RemotePath = RemotePath.Substring(1)
            End If
            If FtpServer.ToLower.StartsWith("ftp://") Then
                FtpServer = FtpServer.Substring(6)
            End If
            If FtpServer.EndsWith("/") Then
                FtpServer = FtpServer.Substring(0, FtpServer.Length - 1)
            End If
            Dim remoteFile As String = "ftp://" + FtpServer + "/" + RemotePath
            Dim response As FtpWebResponse = Nothing
            Dim responseStream As Stream = Nothing
            Dim outputStream As FileStream = Nothing
            Dim bufferSize As Integer = 65536
            Dim buffer(bufferSize - 1) As Byte
            Try
                Dim request As FtpWebRequest = CType(WebRequest.Create(New Uri(remoteFile)), FtpWebRequest)
                request.Method = WebRequestMethods.Ftp.DownloadFile
                request.UseBinary = True
                request.KeepAlive = False
                request.Timeout = Threading.Timeout.Infinite
                request.Credentials = New NetworkCredential(UserName, Password)
                response = CType(request.GetResponse, FtpWebResponse)
                responseStream = response.GetResponseStream
                outputStream = New FileStream(LocalPath, FileMode.Create)
                Dim readCount As Integer = responseStream.Read(buffer, 0, bufferSize)
                While readCount > 0
                    outputStream.Write(buffer, 0, readCount)
                    readCount = responseStream.Read(buffer, 0, bufferSize)
                End While
            Catch ex As Exception
                outputStream.Flush()
                outputStream.Close()
                Array.Clear(buffer, 0, buffer.Length)
                responseStream.Close()
                response.Close()
                Throw ex
            End Try
            If outputStream IsNot Nothing Then
                outputStream.Flush()
                outputStream.Close()
            End If
            If buffer IsNot Nothing Then
                Array.Clear(buffer, 0, buffer.Length)
            End If
            If responseStream IsNot Nothing Then
                responseStream.Close()
            End If
            If response IsNot Nothing Then
                response.Close()
            End If
        End Sub

        ''' <summary>
        ''' 从FTP服务器批量下载符合特定名称的文件。
        ''' </summary>
        ''' <param name="FtpServer">FTP服务器名称。</param>
        ''' <param name="UserName">FTP用户名称。</param>
        ''' <param name="Password">FTP登录密码。</param>
        ''' <param name="RemotePath">FTP文件所在的服务器路径。</param>
        ''' <param name="LocalPath">下载文件后的本地路径。</param>
        ''' <param name="Keyword">特定名称。</param>
        Public Shared Function DownloadFTPFiles(FtpServer As String, UserName As String, Password As String,
                                        RemotePath As String, LocalPath As String, Keyword As String) As Integer
            If RemotePath.EndsWith("/") Or RemotePath.EndsWith("\") Then
                RemotePath = RemotePath.Substring(0, RemotePath.Length - 1)
            End If
            If LocalPath.EndsWith("/") Or LocalPath.EndsWith("\") Then
                LocalPath = LocalPath.Substring(0, LocalPath.Length - 1)
            End If
            Dim lst As List(Of String) = FTPPathFiles(FtpServer, UserName, Password, RemotePath)
            If lst Is Nothing Then
                Return 0
            End If
            Dim countFile As Integer = lst.Count - 1
            Dim count As Integer = 0
            For i As Integer = 0 To countFile
                If lst(i).Contains(Keyword) Then
                    DownloadFTPFile(FtpServer, UserName, Password, RemotePath + "/" + lst(i), LocalPath + "/" + lst(i))
                    count += 1
                End If
            Next
            lst.Clear()
            lst = Nothing
            countFile = Nothing
            Return count
        End Function

        Private Shared ReadOnly Property FTPPathFiles(FtpServer As String, UserName As String, Password As String,
                                        RemotePath As String) As List(Of String)
            Get
                Dim remoteFile As String = "ftp://" + FtpServer + "/" + RemotePath + "/"
                Dim request As FtpWebRequest = CType(WebRequest.Create(New Uri(remoteFile)), FtpWebRequest)
                request.Method = WebRequestMethods.Ftp.ListDirectory
                request.UseBinary = True
                request.KeepAlive = False
                request.Timeout = Threading.Timeout.Infinite
                request.Credentials = New NetworkCredential(UserName, Password)
                Dim response As FtpWebResponse = CType(request.GetResponse, FtpWebResponse)
                Dim responseStream As Stream = response.GetResponseStream
                Dim reader As StreamReader = New StreamReader(responseStream)
                Dim s As String = reader.ReadLine
                Dim lst As List(Of String) = Nothing
                If s <> Nothing Then
                    lst = New List(Of String)
                    While s IsNot Nothing
                        lst.Add(s)
                        s = reader.ReadLine
                    End While
                End If
                reader.Close()
                reader = Nothing
                responseStream.Close()
                responseStream = Nothing
                response.Close()
                response = Nothing
                request = Nothing
                remoteFile = Nothing
                s = Nothing
                Return lst
            End Get
        End Property

        ''' <summary>
        ''' 通过SMTP服务发送邮件。
        ''' </summary>
        ''' <param name="SmtpHost">SMTP服务器地址。</param>
        ''' <param name="SmtpPort">SMTP服务端口号。</param>
        ''' <param name="SenderAddress">发件人邮件地址。</param>
        ''' <param name="DisplayName">发件人显示名称。</param>
        ''' <param name="Receivers">收件人邮件地址。</param>
        ''' <param name="CC">抄送人邮件地址。</param>
        ''' <param name="BCC">密送人邮件地址。</param>
        ''' <param name="Subject">邮件主题。</param>
        ''' <param name="Attachments">附件所在文件的全路径名称。</param>
        Public Shared Sub SendMail(SmtpHost As String, SmtpPort As Integer, SenderAddress As String, DisplayName As String,
                                 Receivers As String, CC As String, BCC As String, Subject As String, Optional Attachments() As String = Nothing)
            SetSmtpService(SmtpHost, SmtpPort, SenderAddress, DisplayName)
            SendMail(Receivers, CC, BCC, Subject, Attachments)
        End Sub

        ''' <summary>
        ''' 通过SMTP服务发送邮件。
        ''' </summary>
        ''' <param name="Receivers">收件人邮件地址。</param>
        ''' <param name="CC">抄送人邮件地址。</param>
        ''' <param name="BCC">密送人邮件地址。</param>
        ''' <param name="Subject">邮件主题。</param>
        ''' <param name="Attachments">附件所在文件的全路径名称。</param>
        Public Shared Sub SendMail(Receivers As String, CC As String, BCC As String, Subject As String, Optional Attachments() As String = Nothing)
            Dim mail As MailMessage = GetNewMailMessage(Receivers, CC, BCC, Subject)
            If Attachments IsNot Nothing Then
                For i As Integer = 0 To Attachments.Length - 1
                    mail.Attachments.Add(New Attachment(Attachments(i)))
                Next
            End If
            Dim smtpClt As SmtpClient = New SmtpClient(m_SmtpHost, m_SmtpPort)
            Dim password As String = String.Empty
            smtpClt.Credentials = New NetworkCredential(mail.From.Address, password)
            smtpClt.DeliveryMethod = SmtpDeliveryMethod.Network
            smtpClt.Send(mail)
            For Each att As Attachment In mail.Attachments
                att.Dispose()
            Next
            smtpClt = Nothing
            mail = Nothing
            password = Nothing
        End Sub

        Private Shared m_SmtpHost As String = Nothing
        Private Shared m_SmtpPort As Integer = Nothing
        Private Shared m_SenderAddress As String = Nothing
        Private Shared m_DisplayName As String = Nothing

        ''' <summary>
        ''' 设置SMTP服务邮件发送参数。
        ''' </summary>
        ''' <param name="SmtpHost">SMTP服务器地址。</param>
        ''' <param name="SmtpPort">SMTP服务端口号。</param>
        ''' <param name="SenderAddress">发件人邮件地址。</param>
        ''' <param name="DisplayName">发件人显示名称。</param>
        Public Shared Sub SetSmtpService(SmtpHost As String, SmtpPort As Integer, SenderAddress As String, DisplayName As String)
            m_SmtpHost = SmtpHost
            m_SmtpPort = SmtpPort
            m_SenderAddress = SenderAddress
            m_DisplayName = DisplayName
        End Sub

        ''' <summary>
        ''' 通过SMTP服务发送邮件。
        ''' </summary>
        ''' <param name="SmtpHost">SMTP服务器地址。</param>
        ''' <param name="SmtpPort">SMTP服务端口号。</param>
        ''' <param name="SenderAddress">发件人邮件地址。</param>
        ''' <param name="DisplayName">发件人显示名称。</param>
        ''' <param name="Receivers">收件人邮件地址。</param>
        ''' <param name="CC">抄送人邮件地址。</param>
        ''' <param name="BCC">密送人邮件地址。</param>
        ''' <param name="Subject">邮件主题。</param>
        ''' <param name="Body">邮件正文。</param>
        ''' <param name="IsHTML">正文是否为HTML格式。</param>
        Public Shared Sub SendMail(SmtpHost As String, SmtpPort As Integer, SenderAddress As String, DisplayName As String,
                                 Receivers As String, CC As String, BCC As String, Subject As String, Body As String, IsHTML As Boolean)
            SetSmtpService(SmtpHost, SmtpPort, SenderAddress, DisplayName)
            SendMail(Receivers, CC, BCC, Subject, Body, IsHTML)
        End Sub

        Private Shared Function GetNewMailMessage(Receivers As String, CC As String, BCC As String, Subject As String) As MailMessage
            Dim mail As MailMessage = New MailMessage
            If Receivers <> Nothing Then
                mail.To.Add(Receivers)
            End If
            If CC <> Nothing Then
                mail.CC.Add(CC)
            End If
            If BCC <> Nothing Then
                mail.Bcc.Add(BCC)
            End If
            If Subject <> Nothing Then
                mail.Subject = Subject
            End If
            mail.From = New MailAddress(m_SenderAddress, m_DisplayName)
            Return mail
        End Function

        ''' <summary>
        ''' 通过SMTP服务发送邮件。
        ''' </summary>
        ''' <param name="Receivers">收件人邮件地址。</param>
        ''' <param name="CC">抄送人邮件地址。</param>
        ''' <param name="BCC">密送人邮件地址。</param>
        ''' <param name="Subject">邮件主题。</param>
        ''' <param name="Body">邮件正文。</param>
        ''' <param name="IsHTML">正文是否为HTML格式。</param>
        Public Shared Sub SendMail(Receivers As String, CC As String, BCC As String, Subject As String, Body As String,
                                 IsHTML As Boolean)
            Dim mail As MailMessage = GetNewMailMessage(Receivers, CC, BCC, Subject)
            mail.IsBodyHtml = IsHTML
            If Body <> Nothing Then
                mail.Body = Body
            End If
            Dim smtpClt As SmtpClient = New SmtpClient(m_SmtpHost, m_SmtpPort)
            Dim password As String = String.Empty
            smtpClt.Credentials = New NetworkCredential(mail.From.Address, password)
            smtpClt.DeliveryMethod = SmtpDeliveryMethod.Network
            smtpClt.Send(mail)
            smtpClt = Nothing
            mail = Nothing
            password = Nothing
        End Sub

    End Class
End Namespace
