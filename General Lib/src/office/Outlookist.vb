Imports Microsoft.Office .Interop .Outlook 

Namespace office
    Public Class Outlookist

        Private Shared m_app As Application = Nothing

        ''' <summary>
        ''' ��õ�ǰ���е�Outlookʵ����
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
        ''' �ʼ�������͡�
        ''' </summary>
        Public Enum MailComponent
            ''' <summary>
            '''�ʼ������ˡ�
            ''' </summary>
            SENDER_EMAIL
            ''' <summary>
            ''' �ʼ����⡣
            ''' </summary>
            SUBJECT
            ''' <summary>
            ''' �ʼ����ġ�
            ''' </summary>
            BODY
        End Enum

        Private Shared m_mailList As List(Of MailItem) = Nothing

        ''' <summary>
        '''  ��ָ����Outlook�ռ����ļ����У���ȡδ�����ɵ��ʼ����ϡ�
        ''' </summary>
        ''' <param name="FolderName">ָ�����ļ������ƣ����ļ�����λ�ڡ��ռ��䡱�С���ΪNothing����Ĭ��Ϊ���ռ��䡱��</param>
        ''' <param name="MinReceivedDate">������ʼ��ռ����ڡ�</param>
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
        ''' ��ȡָ�����ʼ����ض���ɲ��֡�
        ''' </summary>
        ''' <param name="Index">�ʼ���š�</param>
        ''' <param name="Filter">ָ���ʼ�����Ĺ���������</param>
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
        ''' ��ָ�����ʼ����Ϊ��ɡ�
        ''' </summary>
        ''' <param name="Index">�ʼ���š�</param>
        Public Shared Sub CompleteMail(Index As Integer)
            m_mailList.Item(Index).FlagStatus = OlFlagStatus.olFlagComplete
            m_mailList.Item(Index).UnRead = False
            m_mailList.Item(Index).Save()
        End Sub

        ''' <summary>
        ''' �ͷŶ�λ�ʼ������з����ϵͳ��Դ��
        ''' </summary>
        Public Shared Sub UnlocateMails()
            If m_mailList.Count > 0 Then
                m_mailList.Clear()
                m_mailList = Nothing
            End If
        End Sub

        ''' <summary>
        ''' ��ָ����Outlook�ռ����ļ����У���ȡ���ϸ����ؼ��ֱ�����ʼ����ݡ�
        ''' </summary>
        ''' <param name="FolderName">ָ�����ļ������ƣ����ļ�����λ�ڡ��ռ��䡱�С���ΪNothing����Ĭ��Ϊ���ռ��䡱��</param>
        ''' <param name="Filter">���ڹ����ʼ���ɸѡ������</param>
        ''' <param name="Keyword">���������Ĺؼ��֡�</param>
        ''' <returns>�����ʼ�ʱ�������������Է����������ʼ��������������ݡ�</returns>
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
        ''' ��ָ���ռ��˷����ʼ���
        ''' </summary>
        ''' <param name="Receivers">�ռ��˵ĵ����ʼ���ַ�������ڶ���ռ���ʱ���Էֺŷָ���ͬ�ĵ�ַ��</param>
        ''' <param name="CC">�����˵ĵ����ʼ���ַ�������ڶ��������ʱ���Էֺŷָ���ͬ�ĵ�ַ��</param>
        ''' <param name="Subject">�ʼ����⡣</param>
        ''' <param name="Body">�ʼ����ġ�</param>
        ''' <param name="Attachments">�ʼ����������ļ���ȫ·�����ơ�</param>
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
