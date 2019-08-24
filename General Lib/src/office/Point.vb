Namespace office
    Public Class Point

        Private r As Integer
        Private c As Integer

        Friend Property Row() As Integer
            Get
                Return r
            End Get
            Set(ByVal value As Integer)
                r = value
            End Set
        End Property

        Friend Property Column() As Integer
            Get
                Return c
            End Get
            Set(ByVal value As Integer)
                c = value
            End Set
        End Property
    End Class
End Namespace
