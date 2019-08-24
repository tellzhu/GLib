Imports dotNet.db.Admin

Namespace math
    ''' <summary>
    ''' ��ʾһ���������塢�洢���ڴ��еľ��������Ԫ�ؿ�Ϊ�����������ݡ�
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
        ''' ��ȡ�����ĳ��ָ��Ԫ�ء�
        ''' </summary>
        ''' <param name="Row">ָ��Ԫ�����ڵ��С�</param>
        ''' <param name="Column">ָ��Ԫ�����ڵ��С�</param>
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
        ''' ���ؾ������������
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
        ''' ���ؾ������������
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
        '''�ڵ�ǰ�����¸���һ�������γ�һ���¾���
        ''' </summary>
        ''' <param name="mt">�����ӵľ���</param>
        ''' <remarks>�¾��������Ϊԭ���������������֮�͡�</remarks>
        Public Sub AddRows(ByRef mt As Matrix(Of T))
            If mt.columnCount = columnCount Then
                m_List.AddRange(mt.m_List)
                rowCount += mt.rowCount
            End If
        End Sub

        ''' <summary>
        ''' �ڵ�ǰ�����¸���һ�������б��γ�һ���¾���
        ''' </summary>
        ''' <param name="lst">�����ӵ������б�</param>
        ''' <remarks>�¾��������Ϊ���¶���֮�ͣ�1��ԭ�����������2������Ԫ����������ԭ�����������ֵ��</remarks>
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
        ''' ����List���󴴽�һ������
        ''' </summary>
        ''' <param name="Data">������Ԫ�ذ��մ������¡������ҵ�
        ''' ˳�����С��������в��洢�����е�List����</param>
        ''' <param name="ColumnCount">�����������</param>
        ''' <remarks>����������������ĳ˻�Ӧ����List�ĳ�Ա����</remarks>
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
        ''' ����SQL��ѯ���Ľ�������󴴽�һ������
        ''' </summary>
        ''' <param name="Command">SQL��ѯ��䡣</param>
        ''' <remarks>�������������SQL��ѯ����Ӧ��������</remarks>
        Public Sub New(ByVal Command As String)
            Dim data As List(Of T) = ExecuteType(Of T)(Command)
            Init(data, LastFieldCount)
        End Sub

        ''' <summary>
        ''' ��������е�����Ԫ�ز��ͷ���Ӧ���ڴ档
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
