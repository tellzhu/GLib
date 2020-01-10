Imports Microsoft.Office .Interop .Outlook 

Namespace office
    Public Class Outlookist

        Private Shared m_app As Application = Nothing

        ''' <summary>
        ''' 获得当前运行的Outlook实例。
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function LocateCurrentApplication() As Boolean
            m_app = Nothing
            Try
                Dim o As Object = GetObject(, "Outlook.Application")
                m_app = CType(o, Application)
                o = Nothing
            Catch ex As system.Exception
                m_app = Nothing
            End Try
            Return m_app IsNot Nothing
        End Function

        ''' <summary>
        ''' 邮件组件类型。
        ''' </summary>
        Public Enum MailComponent
            ''' <summary>
            '''邮件发件人。
            ''' </summary>
            SENDER_EMAIL
            ''' <summary>
            ''' 邮件标题。
            ''' </summary>
            SUBJECT
            ''' <summary>
            ''' 邮件正文。
            ''' </summary>
            BODY
        End Enum

        Private Shared m_mailList As List(Of MailItem) = Nothing

        ''' <summary>
        '''  在指定的Outlook收件箱文件夹中，获取未标记完成的邮件集合。
        ''' </summary>
        ''' <param name="FolderName">指定的文件夹名称，该文件夹需位于“收件箱”中。若为Nothing，则默认为“收件箱”。</param>
        ''' <param name="MinReceivedDate">最早的邮件收件日期。</param>
        ''' <returns></returns>
        Public Shared Function LocateUnCompleteMails(FolderName As String, MinReceivedDate As Date) As Integer
            If m_app Is Nothing Then
                Return 0
            End If
            Dim nspace As [NameSpace] = m_app.GetNamespace("MAPI")
            Dim folder As MAPIFolder = nspace.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
            If FolderName IsNot Nothing Then
                folder = folder.Folders.Item(FolderName)
            End If
            Dim mail As MailItem
            Dim cnt As Integer = folder.Items.Count
            m_mailList = New List(Of MailItem)
            For i As Integer = cnt To 1 Step -1
                mail = CType(folder.Items(i), MailItem)
                If mail.ReceivedTime.Date >= MinReceivedDate Then
                    If mail.FlagStatus = 0 Then
                        m_mailList.Add(mail)
                    End If
                ElseIf mail.ReceivedTime.Date < MinReceivedDate Then
                    Exit For
                End If
            Next
            If m_mailList.Count = 0 Then
                m_mailList = Nothing
                Return 0
            Else
                Return m_mailList.Count
            End If
        End Function

        ''' <summary>
        ''' 获取指定的邮件的特定组成部分。
        ''' </summary>
        ''' <param name="Index">邮件序号。</param>
        ''' <param name="Filter">指定邮件组件的过滤条件。</param>
        ''' <returns></returns>
        Public Shared Function GetMailMessage(Index As Integer, Filter As MailComponent) As String
            Select Case Filter
                Case MailComponent.BODY
                    Return m_mailList.Item(Index).Body
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' 将指定的邮件标记为完成。
        ''' </summary>
        ''' <param name="Index">邮件序号。</param>
        Public Shared Sub CompleteMail(Index As Integer)
            m_mailList.Item(Index).FlagStatus = OlFlagStatus.olFlagComplete
            m_mailList.Item(Index).UnRead = False
            m_mailList.Item(Index).Save()
        End Sub

        ''' <summary>
        ''' 释放定位邮件过程中分配的系统资源。
        ''' </summary>
        Public Shared Sub UnlocateMails()
            If m_mailList.Count > 0 Then
                m_mailList.Clear()
                m_mailList = Nothing
            End If
        End Sub

        ''' <summary>
        ''' 在指定的Outlook收件箱文件夹中，获取符合给定关键字标题的邮件内容。
        ''' </summary>
        ''' <param name="FolderName">指定的文件夹名称，该文件夹需位于“收件箱”中。若为Nothing，则默认为“收件箱”。</param>
        ''' <param name="Filter">用于过滤邮件的筛选条件。</param>
        ''' <param name="Keyword">过滤条件的关键字。</param>
        ''' <returns>根据邮件时间从晚到早检索，对符合条件的邮件返回其正文内容。</returns>
        Public Shared Function LocateLastMail(FolderName As String, Filter As MailComponent, Keyword As String) As String
            If m_app Is Nothing Then
                Return Nothing
            End If
            Dim nspace As [NameSpace] = m_app.GetNamespace("MAPI")
            Dim folder As MAPIFolder = nspace.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
            If FolderName IsNot Nothing Then
                folder = folder.Folders.Item(FolderName)
            End If
            Dim mail As MailItem
            Dim cnt As Integer = folder.Items.Count
            Dim s As String = Nothing
            For i As Integer = cnt To 1 Step -1
                mail = CType(folder.Items(i), MailItem)
                Select Case Filter
                    Case MailComponent.SENDER_EMAIL
                        If mail.SenderEmailAddress.ToUpper = Keyword.ToUpper Then
                            s = mail.Body
                            Exit For
                        End If
                    Case MailComponent.SUBJECT
                        If mail.Subject.Contains(Keyword) Then
                            s = mail.Body
                            Exit For
                        End If
                End Select
            Next
            Return s
        End Function

        ''' <summary>
        ''' 向指定收件人发送邮件。
        ''' </summary>
        ''' <param name="Receivers">收件人的电子邮件地址。当存在多个收件人时，以分号分隔不同的地址。</param>
        ''' <param name="CC">抄送人的电子邮件地址。当存在多个抄送人时，以分号分隔不同的地址。</param>
        ''' <param name="Subject">邮件主题。</param>
        ''' <param name="Body">邮件正文。</param>
        ''' <param name="Attachments">邮件附件所在文件的全路径名称。</param>
        Public Shared Sub SendMail(ByVal Receivers As String, CC As String,
        ByVal Subject As String, ByVal Body As String, Optional ByVal Attachments As String() = Nothing)
            Dim nspace As [NameSpace] = m_app.GetNamespace("MAPI")
            Dim mail As MailItem = CType(m_app.CreateItem(OlItemType.olMailItem), MailItem)
            With mail
                If Receivers IsNot Nothing Then
                    .To = Receivers
                End If
                If CC IsNot Nothing Then
                    .CC = CC
                End If
                .Subject = Subject
                .Body = Body
                If Attachments IsNot Nothing Then
                    For i As Integer = 0 To Attachments.Length - 1
                        .Attachments.Add(Attachments(i))
                    Next
                End If
            End With
            mail.Send()

            mail = Nothing
            nspace = Nothing
        End Sub

    End Class
End Namespace
