Imports System.Net.Sockets

Namespace net
    Friend Class Transporter
        Private Const TCPBufferSize As Integer = 8192

        Friend Sub SendData(ByRef sock As Socket, ByRef data() As Byte)
            Dim buffer(TCPBufferSize - 1) As Byte
            Dim length As Integer = data.Length
            Dim fullCount As Integer = length \ TCPBufferSize
            If fullCount > 0 Then
                For i As Integer = 1 To fullCount
                    Array.Copy(data, (i - 1) * TCPBufferSize, buffer, 0, TCPBufferSize)
                    sock.Send(buffer, TCPBufferSize, SocketFlags.None)
                Next
            End If
            fullCount *= TCPBufferSize
            If length > fullCount Then
                Array.Copy(data, fullCount, buffer, 0, length - fullCount)
                sock.Send(buffer, length - fullCount, SocketFlags.None)
            End If
            Array.Clear(buffer, 0, TCPBufferSize)
        End Sub

        Friend Function ReceiveData(ByRef sock As Socket) As Byte()
            Dim buffer(TCPBufferSize - 1) As Byte
            Dim lst As List(Of Byte()) = New List(Of Byte())
            Dim totalCount As Integer = 0
            Dim count As Integer = sock.Receive(buffer, TCPBufferSize, SocketFlags.None)
            sock.ReceiveTimeout = 5
            While count > 0
                totalCount += count
                lst.Add(CType(buffer.Clone, Byte()))
                If count < TCPBufferSize Then
                    Exit While
                End If
                count = sock.Receive(buffer, TCPBufferSize, SocketFlags.None)
            End While
            Dim data() As Byte
            ReDim data(totalCount - 1)
            count = lst.Count - 1
            For i As Integer = 0 To count - 1
                Array.Copy(lst(i), 0, data, i * TCPBufferSize, lst(i).Length)
                Array.Clear(lst(i), 0, lst(i).Length)
                lst(i) = Nothing
            Next
            Array.Copy(lst(count), 0, data, count * TCPBufferSize, totalCount - count * TCPBufferSize)
            Array.Clear(lst(count), 0, lst(count).Length)
            lst(count) = Nothing
            lst.Clear()
            Array.Clear(buffer, 0, TCPBufferSize)
            Return data
        End Function

    End Class

End Namespace
