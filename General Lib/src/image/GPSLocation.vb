Imports System.Drawing.Imaging

Namespace image
    Public Class GPSLocation
        Private m_list As List(Of PropertyItem) = Nothing

        Friend Sub Add(ByRef item As PropertyItem)
            m_list.Add(item)
        End Sub

        Public Sub New()
            m_list = New List(Of PropertyItem)
        End Sub

        Friend Sub UpdateImage(ByRef img As Drawing.Image)
            Dim cnt As Integer = m_list.Count - 1
            For i As Integer = 0 To cnt
                img.SetPropertyItem(m_list.Item(i))
            Next
        End Sub

        Protected Overrides Sub Finalize()
            m_list.Clear()
            m_list = Nothing
            MyBase.Finalize()
        End Sub
    End Class
End Namespace
