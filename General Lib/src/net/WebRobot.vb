Imports System.IO
Imports System.Net

Namespace net
	Public Class WebRobot

        Public Shared Function HTTPText(ByVal HttpAddress As String, Optional ByRef Headers As String() = Nothing) As String
            Try
                Dim request As WebRequest = WebRequest.Create(HttpAddress)
                If Headers IsNot Nothing Then
                    For i As Integer = 0 To Headers.Length - 1
                        request.Headers.Add(Headers(i))
                    Next
                End If
                request.Proxy = m_proxy
                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim reader As New StreamReader(response.GetResponseStream())
                Dim s As String = reader.ReadToEnd
                reader.Close()
                response.Close()
                reader = Nothing
                response = Nothing
                request = Nothing
                Return s
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Private Shared m_proxy As WebProxy = Nothing

		Public Shared Sub SetProxyServer(ByVal Host As String, ByVal Port As Integer, _
		Optional ByVal Username As String = Nothing, Optional ByVal Password As String = Nothing)
			m_proxy = New WebProxy(Host, Port)
			If Username <> Nothing Or Password <> Nothing Then
				m_proxy.Credentials = New NetworkCredential(Username, Password)
			End If
		End Sub
	End Class
End Namespace
