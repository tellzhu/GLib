Imports System.IO
Imports dotNet.db.Admin
Imports dotNet.net.NetMaster
Imports dotNet.time.DateExpert

Namespace db
    ''' <summary>
    ''' 将字符串数组批量加载到数据库表的装载器。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class StringLoader
        Private m_MaxSize As Integer = Nothing
        Private m_StringLines() As String = Nothing
        Private m_Index As Integer = Nothing
        Private m_DataTableName As String = Nothing

        ''' <summary>
        ''' 将一个字符串添加到字符串数组的末尾。
        ''' </summary>
        ''' <param name="Str">待添加的字符串。</param>
        ''' <remarks></remarks>
        Public Sub Append(Str As String)
            If m_Index > m_MaxSize Then
                LoadMemoryToDBTable(Nothing)
                m_Index = 0
            End If
            Str = Str.Replace(vbCrLf, "")
            m_StringLines(m_Index) = Str
            m_Index += 1
        End Sub

        ''' <summary>
        ''' 设置或获取字符串数组的长度。
        ''' </summary>
        ''' <value>预设置的字符串数组长度。</value>
        ''' <returns>字符串数组的长度。</returns>
        ''' <remarks></remarks>
        Public Property Size() As Integer
            Get
                If m_StringLines Is Nothing Then
                    Return 0
                Else
                    Return m_StringLines.Length
                End If
            End Get
            Set(value As Integer)
                If m_StringLines IsNot Nothing Then
                    If m_StringLines.Length = value Then
                        m_Index = 0
                        Return
                    End If
                End If
                Clear()
                m_MaxSize = value - 1
                ReDim m_StringLines(m_MaxSize)
                m_Index = 0
            End Set
        End Property

        Private Sub Clear()
            If m_StringLines IsNot Nothing Then
                Array.Clear(m_StringLines, 0, m_StringLines.Length)
                m_StringLines = Nothing
                m_Index = -1
                m_MaxSize = Nothing
            End If
        End Sub

        ''' <summary>
        ''' 设置或获取数据库表的名称。
        ''' </summary>
        ''' <value>预设置的数据库表名称。</value>
        ''' <returns>已设置的数据库表名称。</returns>
        ''' <remarks></remarks>
        Public Property DataTableName As String
            Get
                Return m_DataTableName
            End Get
            Set(value As String)
                m_DataTableName = value
            End Set
        End Property

        ''' <summary>
        ''' 将字符串数组批量装载到数据库表中。
        ''' </summary>
        ''' <param name="DataTableName">待批量装载数据的数据库表名称。</param>
        ''' <param name="FileName">用于保存字符串数组数据的临时文件名称。</param>
        ''' <remarks></remarks>
        Public Sub Load(DataTableName As String, Optional FileName As String = Nothing)
            If m_Index = 0 Then
                Return
            End If
            If m_Index <= m_MaxSize Then
                Dim tempStr(m_Index - 1) As String
                Array.Copy(m_StringLines, tempStr, m_Index)
                ReDim m_StringLines(m_Index - 1)
                Array.Copy(tempStr, m_StringLines, m_Index)
                Array.Clear(tempStr, 0, m_Index)
            End If
            m_DataTableName = DataTableName
            LoadMemoryToDBTable(FileName)
        End Sub

        Private Sub LoadMemoryToDBTable(TempFileName As String)
            If TempFileName = Nothing Then
                TempFileName = "Data_" + Format(Now, "YYYYMMDDHHMMSS")
            End If
            Dim fullFileName As String = My.Computer.FileSystem.SpecialDirectories.Temp + "\" + TempFileName + ".csv"
            If MetaData.IPAddress <> GetLocalHostIPAddress() Then
                fullFileName = DataMover.TransitDirectory + "\" + TempFileName + ".csv"
            End If
            File.WriteAllLines(fullFileName, m_StringLines, System.Text.Encoding.Unicode)
            LoadBulkData(fullFileName, m_DataTableName)
            File.Delete(fullFileName)
            fullFileName = Nothing
        End Sub

        Protected Overrides Sub Finalize()
            Clear()
            m_DataTableName = Nothing
            MyBase.Finalize()
        End Sub

        ''' <summary>
        ''' 构建一个新的数据装载器，并同时设置其字符串数组的长度。
        ''' </summary>
        ''' <param name="Size">预设置的字符串数组长度，默认值为10000。</param>
        ''' <remarks></remarks>
        Public Sub New(Optional Size As Integer = 10000)
            Me.Size = Size
        End Sub
    End Class
End Namespace