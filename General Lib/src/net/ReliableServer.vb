Imports System.Net
Imports System.Net.Sockets
Imports System.Threading

Namespace net
    Public Class ReliableServer

        Private m_Port As Integer = Nothing
        Private m_IsStopped As Boolean = False
        Private m_ListenerSocket As Socket = Nothing
        Private m_protocol As IReliableProtocol = Nothing

        Public Sub New(Port As Integer, ByRef protocol As IReliableProtocol)
            m_Port = Port
            m_protocol = protocol
        End Sub

        Private Sub WaitForConnection(ByVal state As Object)
            m_ListenerSocket.Listen(0)
            Try
                While Not m_IsStopped
                    Dim serverSocket As Socket = m_ListenerSocket.Accept
                    ThreadPool.QueueUserWorkItem(AddressOf HandleClientSocket, serverSocket)
                End While
            Catch ex As SocketException
                Throw ex
            End Try
        End Sub

        Private Sub HandleClientSocket(clnt As Object)
            Dim sock As Socket = CType(clnt, Socket)
            m_protocol.Response(sock)
            sock.Close(1)
            sock = Nothing
        End Sub

        Public Sub Start()
            m_ListenerSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            m_ListenerSocket.Bind(New IPEndPoint(IPAddress.Parse("127.0.0.1"), m_Port))
            ThreadPool.SetMaxThreads(6, 6)
            ThreadPool.QueueUserWorkItem(AddressOf WaitForConnection)
        End Sub

        Public Sub StopService()
            m_IsStopped = True
            m_ListenerSocket.Close()
            m_ListenerSocket = Nothing
        End Sub

    End Class
End Namespace
