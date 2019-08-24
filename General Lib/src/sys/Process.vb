Namespace sys
    Public Class Process

        ''' <summary>
        ''' 启动一个新的Windows进程。
        ''' </summary>
        ''' <param name="ProcessName">Windows进程的文件全名。</param>
        ''' <param name="Path">Windows进程文件所在的工作目录。</param>
        ''' <remarks></remarks>
        Public Shared Sub Start(ByVal ProcessName As String, Optional ByVal Path As String = Nothing)
            Dim startInfo As ProcessStartInfo = New ProcessStartInfo
            If Path = Nothing Then
                startInfo.WorkingDirectory = My.Application.Info.DirectoryPath
            Else
                startInfo.WorkingDirectory = Path
            End If
            startInfo.FileName = ProcessName
            System.Diagnostics.Process.Start(startInfo)
        End Sub

        ''' <summary>
        ''' 终止指定名称的Windows进程。
        ''' </summary>
        ''' <param name="ProcessName">Windows进程名称。</param>
        ''' <remarks>若同一进程名称具备多个实例，则终止全部实例。</remarks>
        Private Shared Sub Kill(ByVal ProcessName As String)
            Dim p As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName(ProcessName)
            If p.Length > 0 Then
                For i As Integer = 0 To p.Length - 1
                    p(i).Kill()
                Next
                Array.Clear(p, 0, p.Length)
            End If
            p = Nothing
        End Sub

        ''' <summary>
        ''' 终止指定进程ID的Windows进程。
        ''' </summary>
        ''' <param name="ProcessId">Windows进程ID。</param>
        Friend Shared Sub Kill(ProcessId As Integer)
            Try
                Dim p As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessById(ProcessId)
                If p IsNot Nothing Then
                    p.Kill()
                    p = Nothing
                End If
            Catch ex As Exception
            End Try
        End Sub
    End Class
End Namespace
