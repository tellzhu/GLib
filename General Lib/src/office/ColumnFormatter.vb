Imports dotNet.office.AreaExpert

Namespace office
    ''' <summary>
    ''' 将行列形式的二维数组以指定格式映射到Excel单元格区域中。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ColumnFormatter

        Private m_columns As Integer() = Nothing
        Private m_Formatter As Dictionary(Of Integer, FormatType) = Nothing

        ''' <summary>
        ''' 设置矩阵列对应Excel单元格的显示格式。
        ''' </summary>
        ''' <param name="columns">矩阵的列标标号数组。</param>
        ''' <param name="formats">Excel单元格的显示格式数组。</param>
        ''' <remarks></remarks>
        Private Sub SetColumnsFormat(ByRef columns As Integer(), ByRef formats As FormatType())
            If columns.Length <> formats.Length Then
                Return
            End If
            m_columns = columns
            For i As Integer = 0 To m_columns.Length - 1
                If m_Formatter.ContainsKey(m_columns(i)) Then
                    m_Formatter.Item(m_columns(i)) = formats(i)
                Else
                    m_Formatter.Add(m_columns(i), formats(i))
                End If
            Next
        End Sub

        ''' <summary>
        ''' 获得指定矩阵列的Excel单元格格式。
        ''' </summary>
        ''' <param name="index">指定的矩阵列列号在整个格式列数组中的序号。</param>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property ColumnFormat(ByVal index As Integer) As FormatType
            Get
                Return m_Formatter.Item(m_columns(index))
            End Get
        End Property

        ''' <summary>
        ''' 获得指定矩阵列的列号。
        ''' </summary>
        ''' <param name="index">要获取的矩阵列列号在整个格式列数组中的序号。</param>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property ColumnNo(ByVal index As Integer) As Integer
            Get
                Return m_columns(index)
            End Get
        End Property

        ''' <summary>
        ''' 获得已设置矩阵列格式的数据列的数量。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property Count() As Integer
            Get
                If m_Formatter Is Nothing Then
                    Return 0
                End If
                Return m_Formatter.Count
            End Get
        End Property

        ''' <summary>
        ''' 以特定的矩阵列及Excel单元格显示格式数组初始化一个ColumnFormatter。
        ''' </summary>
        ''' <param name="columns">矩阵的列标标号数组。</param>
        ''' <param name="formats">Excel单元格的显示格式数组。</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef columns As Integer(), ByRef formats As FormatType())
            Me.New()
            SetColumnsFormat(columns, formats)
        End Sub

        ''' <summary>
        ''' 创建一个ColumnFormatter。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            m_Formatter = New Dictionary(Of Integer, FormatType)
        End Sub

        Protected Overrides Sub Finalize()
            Array.Clear(m_columns, 0, m_columns.Length)
            m_columns = Nothing
            m_Formatter.Clear()
            m_Formatter = Nothing
            MyBase.Finalize()
        End Sub
    End Class
End Namespace
