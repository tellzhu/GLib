Namespace math
    Public Class Triangle(Of T)

        Private m_list As List(Of T) = Nothing
        Private edge As Integer = Nothing
        Private angle As Position

        Public Enum Position
            NorthWest
            NorthEast
        End Enum

        Public Sub Clear()
            m_list.Clear()
            edge = 0
        End Sub

        Friend ReadOnly Property EdgeLength() As Integer
            Get
                Return edge
            End Get
        End Property

        Friend ReadOnly Property Cell(ByVal Row As Integer, ByVal Column As Integer) As T
            Get
                If edge <= 0 Then
                    Return Nothing
                End If
                Dim o As T = Nothing
                If Row + Column <= edge + 1 Then
                    o = m_list.Item(((2 * edge - Row + 2) * (Row - 1)) \ 2 + Column - 1)
                End If
                If IsDBNull(o) Then
                    Return Nothing
                Else
                    Return o
                End If
            End Get
        End Property

        Public Sub New(ByRef data As List(Of T), Optional ByVal AnglePosition As Position = Position.NorthWest)
            m_list = New List(Of T)(data)
            edge = (CInt(System.Math.Sqrt(8 * m_list.Count + 1)) - 1) \ 2
            angle = AnglePosition
        End Sub

        Protected Overrides Sub Finalize()
            Clear()
            m_list = Nothing
            MyBase.Finalize()
        End Sub
    End Class
End Namespace
