Imports System.Net
Imports System.Net.Sockets

Namespace net
    Public Class ReliableClient
        Private m_Socket As Socket = Nothing
        Private m_Protocol As IReliableProtocol = Nothing

        Public Sub New(Host As String, Port As Integer, ByRef Protocol As IReliableProtocol)
            m_Protocol = Protocol
            m_Socket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            m_Socket.Connect(New IPEndPoint(IPAddress.Parse(Host), Port))
        End Sub

        Public Sub Request(stateinfo As Object)
            m_Protocol.Request(m_Socket, stateinfo)
        End Sub

        Public Sub Close()
            m_Socket.Close()
            m_Socket = Nothing
        End Sub

    End Class
End Namespace
