Namespace text
    Public Class StringHandler

        Friend Shared Function SerializabledString(ByRef values() As Decimal, _
                                           Optional ByVal SeparateChar As Char = ","c) As String
            Dim s As String = ""
            Dim len As Integer = values.Length - 2
            For i As Integer = 0 To len
                s = s + CStr(values(i)) + SeparateChar
            Next
            s = s + CStr(values(len + 1))
            len = Nothing
            Return s
        End Function

        ''' <summary>
        ''' �����ַ��������Զ������ַ������͵Ĺ�ϣ���ϡ�
        ''' </summary>
        ''' <param name="s">��ת�����ַ������顣</param>
        ''' <param name="IsClearArray">ת��Ϊ��ϣ���Ϻ��Ƿ��Զ�������ͷ�ԭ�ַ������顣</param>
        ''' <returns>��Ӧ���ַ�������Ĺ�ϣ���ϡ�</returns>
        ''' <remarks>�ַ����������ظ���Ԫ�ؽ�ֻ����һ����</remarks>
        Public Shared Function StringsToHashSet(ByRef s() As String, Optional ByVal IsClearArray As Boolean = True) As HashSet(Of String)
            If s Is Nothing Then
                Return Nothing
            End If
            If s.Length = 0 Then
                Return Nothing
            End If
            Dim hs As HashSet(Of String) = New HashSet(Of String)
            For i As Integer = 0 To s.Length - 1
                hs.Add(s(i))
            Next
            If IsClearArray Then
                Array.Clear(s, 0, s.Length)
                s = Nothing
            End If
            Return hs
        End Function

        Friend Shared Function NumberStringTrim(ByVal str As String) As String
            If str = Nothing Then
                Return Nothing
            End If
            If str.Length = 0 Then
                Return str
            End If
            If str.IndexOf(".") = -1 Then
                Return str
            End If
            Dim pos As Integer = str.Length
            While str.Chars(pos - 1) = "0"
                pos -= 1
            End While
            If str.Chars(pos - 1) = "." Then
                pos -= 1
            End If
            Return str.Substring(0, pos)
        End Function

        ''' <summary>
        ''' �����ض����еĹؼ��ʼ����ַ����������ذ�˳��ɸѡ������ַ�����
        ''' </summary>
        ''' <param name="s">���������ַ�����</param>
        ''' <param name="Milestone">�ؼ���ʶ�������С�</param>
        ''' <param name="Terminal">�ַ�������ֹ��ʶ����</param>
        ''' <returns>����ɸѡ�����ַ�����</returns>
        ''' <remarks></remarks>
        Public Shared Function KeySubstring(ByVal s As String, _
        ByVal Milestone() As String, Optional ByVal Terminal As String = Nothing) As String
            If s Is Nothing Then
                Return Nothing
            End If
            Dim index As Integer
            For i As Integer = 0 To Milestone.Length - 1
                index = s.IndexOf(Milestone(i))
                If index = -1 Then
                    s = Nothing
                    index = Nothing
                    Return Nothing
                Else
                    s = s.Substring(index + Milestone(i).Length)
                End If
            Next
            If Terminal Is Nothing Then
                Return s.Trim
            Else
                index = s.IndexOf(Terminal)
                If index = -1 Then
                    s = Nothing
                    index = Nothing
                    Return Nothing
                Else
                    Return s.Substring(0, index).Trim
                End If
            End If
        End Function

    End Class
End Namespace
