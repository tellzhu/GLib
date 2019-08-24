Imports System.Net.Sockets

Namespace net
    Public Interface IReliableProtocol
        Function Request(ByRef sock As Socket, stateinfo As Object) As Object
        Sub Response(ByRef sock As Socket)
    End Interface
End Namespace
