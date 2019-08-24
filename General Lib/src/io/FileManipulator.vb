Imports System.IO

Namespace io
    Public Class FileManipulator
        ''' <summary>
        ''' 删除文件到回收站。
        ''' </summary>
        ''' <param name="FileName">需删除的文件名。</param>
        Public Shared Sub DeleteFile(ByVal FileName As String)
            FileIO.FileSystem.DeleteFile(FileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
        End Sub

        ''' <summary>
        ''' 将文件复制到新的位置。
        ''' </summary>
        ''' <param name="sourceFileName">要复制的文件。</param>
        ''' <param name="destinationFileName">文件应复制到的位置。</param>
        Public Shared Sub CopyFile(sourceFileName As String, destinationFileName As String)
            FileIO.FileSystem.CopyFile(sourceFileName, destinationFileName, True)
        End Sub

        Public Shared Sub DeleteDirectory(ByVal DirectoryName As String)
            FileIO.FileSystem.DeleteDirectory(DirectoryName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
        End Sub

        Public Shared Function IsHidden(ByVal FileName As String) As Boolean
            Dim fi As FileInfo = FileIO.FileSystem.GetFileInfo(FileName)
            Dim b As Boolean = (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden
            fi = Nothing
            Return b
        End Function

        Public Shared ReadOnly Property ParentPath(ByVal Path As String, Optional ByVal Level As Integer = 1) As String
            Get
                Dim s As String = Path
                For i As Integer = 1 To Level
                    s = FileIO.FileSystem.GetParentPath(s)
                Next
                Return s
            End Get
        End Property

    End Class
End Namespace
