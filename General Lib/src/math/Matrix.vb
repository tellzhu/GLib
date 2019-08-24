Imports dotNet.db.Admin

Namespace math
    ''' <summary>
    ''' 表示一个抽象意义、存储在内存中的矩阵，其矩阵元素可为任意类型数据。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Matrix(Of T)

        Private m_List As List(Of T) = Nothing
        Private rowCount As Integer = Nothing
        Private columnCount As Integer = Nothing

        Public Function Transform() As Triangle(Of T)
            If rowCount <= 1 Or columnCount <> 1 Then
                Return Nothing
            End If
            Dim n As Integer = (CInt(System.Math.Sqrt(8 * rowCount + 1)) - 1) \ 2
            If n * (n + 1) <> 2 * rowCount Then
                Return Nothing
            Else
                Return New Triangle(Of T)(m_List)
            End If
        End Function

        ''' <summary>
        ''' 获取矩阵的某个指定元素。
        ''' </summary>
        ''' <param name="Row">指定元素所在的行。</param>
        ''' <param name="Column">指定元素所在的列。</param>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Cell(ByVal Row As Integer, ByVal Column As Integer) As T
            Get
                If rowCount <= 0 Then
                    Return Nothing
                End If
                Dim o As T = m_List.Item((Row - 1) * columnCount + Column - 1)
                If IsDBNull(o) Then
                    Return Nothing
                Else
                    Return o
                End If
            End Get
            Set(ByVal value As T)
                If rowCount > 0 Then
                    m_List.Item((Row - 1) * columnCount + Column - 1) = value
                End If
            End Set
        End Property

        ''' <summary>
        ''' 返回矩阵的总列数。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ColumnsCount() As Integer
            Get
                Return columnCount
            End Get
        End Property

        ''' <summary>
        ''' 返回矩阵的总行数。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property RowsCount() As Integer
            Get
                Return rowCount
            End Get
        End Property

        ''' <summary>
        '''在当前矩阵下附加一个矩阵，形成一个新矩阵。
        ''' </summary>
        ''' <param name="mt">待附加的矩阵。</param>
        ''' <remarks>新矩阵的行数为原来两个矩阵的行数之和。</remarks>
        Public Sub AddRows(ByRef mt As Matrix(Of T))
            If mt.columnCount = columnCount Then
                m_List.AddRange(mt.m_List)
                rowCount += mt.rowCount
            End If
        End Sub

        ''' <summary>
        ''' 在当前矩阵下附加一个数组列表，形成一个新矩阵。
        ''' </summary>
        ''' <param name="lst">待附加的数组列表。</param>
        ''' <remarks>新矩阵的行数为以下二者之和：1、原矩阵的行数；2、数组元素数量整除原矩阵列数后的值。</remarks>
        Public Sub AddRows(ByRef lst As List(Of T))
            If lst.Count Mod Me.columnCount = 0 Then
                m_List.AddRange(lst)
                rowCount += lst.Count \ Me.columnCount
            End If
        End Sub

        Public Sub RemoveRow(ByVal Row As Integer)
            If Row <= rowCount Then
                For i As Integer = 1 To columnCount
                    m_List.RemoveAt(columnCount * (Row - 1))
                Next
                rowCount -= 1
            End If
        End Sub

        ''' <summary>
        ''' 利用List对象创建一个矩阵。
        ''' </summary>
        ''' <param name="Data">将矩阵元素按照从上至下、从左到右的
        ''' 顺序逐行、逐列排列并存储在其中的List对象。</param>
        ''' <param name="ColumnCount">矩阵的列数。</param>
        ''' <remarks>矩阵的行数与列数的乘积应等于List的成员数。</remarks>
        Public Sub New(ByRef Data As List(Of T), ByVal ColumnCount As Integer)
            Init(Data, ColumnCount)
        End Sub

        Private Sub Init(ByRef Data As List(Of T), ByVal ColumnCount As Integer)
            If Data Is Nothing Then
                m_List = Nothing
                Me.columnCount = 0
                Me.rowCount = 0
            Else
                m_List = Data
                Me.columnCount = ColumnCount
                If Me.columnCount > 0 Then
                    rowCount = m_List.Count \ Me.columnCount
                Else
                    rowCount = 0
                End If
            End If
        End Sub

        ''' <summary>
        ''' 利用SQL查询语句的结果集对象创建一个矩阵。
        ''' </summary>
        ''' <param name="Command">SQL查询语句。</param>
        ''' <remarks>矩阵的列数等于SQL查询语句对应的列数。</remarks>
        Public Sub New(ByVal Command As String)
            Dim data As List(Of T) = ExecuteType(Of T)(Command)
            Init(data, LastFieldCount)
        End Sub

        ''' <summary>
        ''' 清除矩阵中的所有元素并释放相应的内存。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Clear()
            If m_List IsNot Nothing Then
                m_List.Clear()
                rowCount = 0
                columnCount = 0
            End If
        End Sub

        Protected Overrides Sub Finalize()
            If m_List IsNot Nothing Then
                m_List.Clear()
                m_List = Nothing
            End If
            MyBase.Finalize()
        End Sub

    End Class
End Namespace
