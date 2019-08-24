Imports System.Net.Sockets

Namespace net
    Public Class DemoProtocol
        Implements IReliableProtocol

        Public Function Request(ByRef sock As Socket, stateinfo As Object) As Object Implements IReliableProtocol.Request
            Dim msg As String = CStr(stateinfo)
            Dim data() As Byte = System.Text.Encoding.ASCII.GetBytes(msg)
            Dim tran As Transporter = New Transporter
            tran.SendData(sock, data)
            data = tran.ReceiveData(sock)
            tran = Nothing
            msg = System.Text.Encoding.ASCII.GetString(data)
            MsgBox(msg)
            Return Nothing
        End Function

        Public Sub Response(ByRef sock As Socket) Implements IReliableProtocol.Response
            Dim trans As Transporter = New Transporter
            Dim data() As Byte = trans.ReceiveData(sock)
            Dim msg As String = System.Text.Encoding.ASCII.GetString(data, 0, data.Length)
            Console.WriteLine("Server: " + msg)
            data = System.Text.Encoding.ASCII.GetBytes(msg.ToUpper)
            trans.SendData(sock, data)
        End Sub
    End Class

End Namespace
