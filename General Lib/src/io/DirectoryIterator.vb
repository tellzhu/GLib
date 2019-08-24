Imports Microsoft.VisualBasic.FileIO
Imports System.IO

Namespace io
    Public Class DirectoryIterator

        Private m_CurrentFullName As String = Nothing
        Private m_IsFile As Boolean = Nothing

        Public Function MoveNext() As Boolean
            If m_SearchStack.Count = 0 Then
                Return False
            Else
                m_CurrentFullName = m_SearchStack.Pop
                m_IsFile = (m_CurrentFullName.Chars(0) = "F"c)
                m_CurrentFullName = m_CurrentFullName.Substring(1)
                If Not m_IsFile And FileSystem.DirectoryExists(m_CurrentFullName) Then
                    AddChildren(m_CurrentFullName)
                End If
                Return True
            End If
        End Function

        Public ReadOnly Property Current() As String
            Get
                Return m_CurrentFullName
            End Get
        End Property

        Public ReadOnly Property CurrentName As String
            Get
                Return FileSystem.GetName(m_CurrentFullName)
            End Get
        End Property

        Public ReadOnly Property IsFile As Boolean
            Get
                Return m_IsFile
            End Get
        End Property

        Public ReadOnly Property IsEmptyDirectory As Boolean
            Get
                Return Directory.GetFileSystemEntries(m_CurrentFullName).Length = 0
            End Get
        End Property

        Public ReadOnly Property Exists As Boolean
            Get
                Return FileSystem.DirectoryExists(m_CurrentFullName) _
                    Or FileSystem.FileExists(m_CurrentFullName)
            End Get
        End Property

        Private m_SearchStack As Stack(Of String) = Nothing

        Private Sub AddChildren(ByVal DirectoryName As String)
            Try
                Dim s() As String = Directory.GetFiles(DirectoryName)
                Dim len As Integer = s.Length - 1
                For i As Integer = 0 To len
                    m_SearchStack.Push("F" + s(i))
                Next
                Array.Clear(s, 0, s.Length)
                s = Directory.GetDirectories(DirectoryName)
                len = s.Length - 1
                For i As Integer = 0 To len
                    m_SearchStack.Push("D" + s(i))
                Next
                Array.Clear(s, 0, s.Length)
                s = Nothing
                len = Nothing
            Catch ex As UnauthorizedAccessException
            End Try
        End Sub

        Public Sub New(DirectoryName As String)
            m_SearchStack = New Stack(Of String)
            AddChildren(DirectoryName)
        End Sub

        Protected Overrides Sub Finalize()
            m_SearchStack.Clear()
            m_SearchStack = Nothing
            m_CurrentFullName = Nothing
            m_IsFile = Nothing
            MyBase.Finalize()
        End Sub
    End Class
End Namespace
