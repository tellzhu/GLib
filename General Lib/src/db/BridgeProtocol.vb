Imports System.Net.Sockets
Imports dotNet.net

Namespace db
    Public Class BridgeProtocol
        Implements IReliableProtocol

        Public Sub Response(ByRef sock As Socket) Implements IReliableProtocol.Response
            Throw New NotImplementedException()
        End Sub

        Public Function Request(ByRef sock As Socket, stateinfo As Object) As Object Implements IReliableProtocol.Request
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace

