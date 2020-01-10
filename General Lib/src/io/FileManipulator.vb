Imports System.IO
Imports System.IO.Compression

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

        Public Shared Function FileExists(fileNames As String(), IsAutoDelete As Boolean) As Boolean
            If fileNames Is Nothing Then
                Return False
            End If
            Dim FinalFileIsNotExist As Boolean = False
            For i As Integer = 0 To fileNames.Length - 1
                If Not File.Exists(fileNames(i)) Then
                    FinalFileIsNotExist = True
                    Exit For
                End If
            Next
            If FinalFileIsNotExist And IsAutoDelete Then
                For i As Integer = fileNames.Length - 1 To 0 Step -1
                    If File.Exists(fileNames(i)) Then
                        File.Delete(fileNames(i))
                    End If
                Next
            End If
            Return Not FinalFileIsNotExist
        End Function

        ''' <summary>
        ''' 将ZIP文件解压缩至当前目录。若Path参数指明的是一个目录，则解压缩该目录下的所有ZIP文件；若Path参数指明的是一个ZIP文件，则解压缩该文件。
        ''' </summary>
        ''' <param name="Path">ZIP文件所在的目录名称，或ZIP文件名称。</param>
        ''' <param name="ExcludeFileName">不需要解压缩的ZIP文件名称，该参数仅在Path参数指明为目录时有效。</param>
        Public Shared Sub ExtractZips(Path As String, Optional ExcludeFileName As String = Nothing)
            If Directory.Exists(Path) Then
                Dim destDir As DirectoryInfo = New DirectoryInfo(Path)
                For Each zipF As FileInfo In destDir.GetFiles("*.zip")
                    If zipF.Name.IndexOf(ExcludeFileName) = -1 Then
                        Try
                            ZipFile.ExtractToDirectory(Path + "\\" + zipF.Name, Path)
                        Catch ex As Exception
                        End Try
                    End If
                Next
                Return
            End If
            If File.Exists(Path) Then
                Dim file As FileInfo = New FileInfo(Path)
                Try
                    ZipFile.ExtractToDirectory(Path, file.DirectoryName)
                Catch ex As Exception
                End Try
            End If
        End Sub

        Public Shared Sub DeleteDirectory(ByVal DirectoryName As String)
            FileIO.FileSystem.DeleteDirectory(DirectoryName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
        End Sub

        Public Shared Function IsHidden(ByVal FileName As String) As Boolean
            Dim fi As FileInfo = FileIO.FileSystem.GetFileInfo(FileName)
            Dim b As Boolean = (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden
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
