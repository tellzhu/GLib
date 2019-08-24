Namespace io

    Public Class Logger

        Private Shared logFile As String
        Private Shared orderId As Integer = 1

        Public Shared Sub Init(ByVal FileName As String)
            logFile = FileName
            My.Computer.FileSystem.WriteAllText(logFile, "", False)
        End Sub

		Public Shared Sub Logging(ByVal text As String)
			My.Computer.FileSystem.WriteAllText(logFile, _
				 orderId & ": " & Now() & "  " + text + vbCrLf, True)
			orderId += 1
		End Sub
    End Class
End Namespace
