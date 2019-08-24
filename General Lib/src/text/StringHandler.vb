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
        ''' 根据字符串数组自动生成字符串类型的哈希集合。
        ''' </summary>
        ''' <param name="s">待转换的字符串数组。</param>
        ''' <param name="IsClearArray">转换为哈希集合后，是否自动清除并释放原字符串数组。</param>
        ''' <returns>对应于字符串数组的哈希集合。</returns>
        ''' <remarks>字符串数组中重复的元素将只保留一个。</remarks>
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
        ''' 根据特定序列的关键词检索字符串，并返回按顺序筛选后的子字符串。
        ''' </summary>
        ''' <param name="s">待分析的字符串。</param>
        ''' <param name="Milestone">关键标识符的序列。</param>
        ''' <param name="Terminal">字符串的终止标识符。</param>
        ''' <returns>经过筛选的子字符串。</returns>
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
