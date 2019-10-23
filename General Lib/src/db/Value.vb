Imports dotNet.db.Admin
Imports dotNet.db.MetaData
Imports dotNet.math

Namespace db
    ''' <summary>
    ''' 获取存储于数据库系统中的常见值。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Value

        ''' <summary>
        ''' 获取SQL语句查询结果对应的类型值。
        ''' </summary>
        ''' <typeparam name="T">查询结果的数据类型。</typeparam>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns>SQL语句查询结果集理论上应只包含一个元素，否则可能产生预期外的结果。</returns>
        Public Shared Function Val(Of T)(ByVal Command As String) As T
            Dim obj As Object = Execute(Command).Item(0)
            If IsDBNull(obj) Then
                Return Nothing
            Else
                Return CType(obj, T)
            End If
        End Function

        ''' <summary>
        '''  获取SQL语句查询结果对应的“键-值”对字典表。
        ''' </summary>
        ''' <typeparam name="T">键的数据类型。</typeparam>
        ''' <typeparam name="V">值的数据类型。</typeparam>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns>SQL语句的查询结果集为两列，其中第一列为字典表的键，第二列为字典表的值。</returns>
        Public Shared Function Pair(Of T, V)(Command As String) As Dictionary(Of T, V)
            Dim mt As Matrix(Of Object) = New Matrix(Of Object)(Command)
            Dim cnt As Integer = mt.RowsCount
            If cnt = 0 Then
                mt.Clear()
                Return Nothing
            End If
            Dim dict As New Dictionary(Of T, V)
            For i As Integer = 1 To cnt
                dict.Add(CType(mt.Cell(i, 1), T), CType(mt.Cell(i, 2), V))
            Next
            mt.Clear()
            Return dict
        End Function

        ''' <summary>
        ''' 获取SQL语句查询结果对应的字符串值。
        ''' </summary>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns></returns>
        ''' <remarks>SQL语句查询结果集理论上应只包含一个元素，否则可能产生预期外的结果。
        ''' </remarks>
        Public Shared Function Str(ByVal Command As String) As String
            Dim obj As Object = Execute(Command).Item(0)
            If IsDBNull(obj) Then
                Return Nothing
            Else
                Return CStr(obj)
            End If
        End Function

        ''' <summary>
        ''' 获取SQL语句查询结果对应的整数值。
        ''' </summary>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns></returns>
        ''' <remarks>SQL语句查询结果集理论上应只包含一个元素，否则可能产生预期外的结果。</remarks>
        Public Shared Function Int(ByVal Command As String) As Integer
            Dim obj As Object = Execute(Command).Item(0)
            If IsDBNull(obj) Then
                Return 0
            Else
                Return CInt(obj)
            End If
        End Function

        ''' <summary>
        ''' 获取SQL语句查询结果对应的浮点数值。
        ''' </summary>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns></returns>
        ''' <remarks>SQL语句查询结果集理论上应只包含一个元素，否则可能产生预期外的结果。</remarks>
        Public Shared Function Dbl(ByVal Command As String) As Double
            Dim obj As Object = Execute(Command).Item(0)
            If IsDBNull(obj) Then
                Return 0
            Else
                Return CDbl(obj)
            End If
        End Function

        ''' <summary>
        ''' 获取SQL语句查询结果对应的日期值。
        ''' </summary>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns></returns>
        ''' <remarks>SQL语句查询结果集理论上应只包含一个元素，否则可能产生预期外的结果。
        ''' </remarks>
        Public Shared Function Dte(ByVal Command As String) As Date
            Dim obj As Object = Execute(Command).Item(0)
            If IsDBNull(obj) Then
                Return Nothing
            Else
                Return CDate(obj)
            End If
        End Function

        ''' <summary>
        ''' 获取SQL语句查询结果对应的“键-值”对字典表。
        ''' </summary>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns></returns>
        ''' <remarks>SQL语句的查询结果集应为两列，其中第一列为字典表的键，第二列为字典表的值。</remarks>
        Public Shared Function Pair(Command As String) As Dictionary(Of String, String)
            Dim mt As Matrix(Of String) = New Matrix(Of String)(Command)
            Dim cnt As Integer = mt.RowsCount
            Dim dict As Dictionary(Of String, String) = Nothing
            If cnt > 0 Then
                dict = New Dictionary(Of String, String)
                For i As Integer = 1 To cnt
                    dict.Add(mt.Cell(i, 1), mt.Cell(i, 2))
                Next
                mt.Clear()
            End If
            Return dict
        End Function

        ''' <summary>
        ''' 获取SQL语句查询结果对应的精确小数值。
        ''' </summary>
        ''' <param name="Command">SQL语句。</param>
        ''' <returns></returns>
        ''' <remarks>SQL语句查询结果集理论上应只包含一个元素，否则可能产生预期外的结果。
        ''' </remarks>
        Public Shared Function Dec(ByVal Command As String) As Decimal
            Dim obj As Object = Execute(Command).Item(0)
            If IsDBNull(obj) Then
                Return 0
            Else
                Return CDec(obj)
            End If
        End Function

        Public Shared Function MonthKey(ByVal ColumnName As String) As String
            Select Case DatabaseType
                Case DBType.DB2
                    Return " (YEAR(" + ColumnName + ")*12+MONTH(" + ColumnName + ")) "
                Case Else
                    MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' 获得指定SQL查询条件的数据库表中的结果集行数。
        ''' </summary>
        ''' <param name="TableName">数据库表名称。</param>
        ''' <param name="Condition">SQL查询条件。</param>
        ''' <returns>返回查询结果集的行数。</returns>
        ''' <remarks></remarks>
        Public Shared Function Count(ByVal TableName As String, Optional ByVal Condition As String = Nothing) As Integer
            Dim cond As String = ""
            If Condition <> Nothing Then
                cond = " WHERE " + Condition
            End If
            Return Int("SELECT COUNT(1) FROM " + TableName + cond)
        End Function

        ''' <summary>
        ''' 获取指定数据表中的某个数据列的下一个最大整数值。
        ''' </summary>
        ''' <param name="TableName">指定的数据表名称。</param>
        ''' <param name="ColumnName">指定的数据列名称。</param>
        ''' <returns>当前数据列中的最大值加1。</returns>
        ''' <remarks>若指定的数据表为空，则返回1。</remarks>
        Public Shared Function NextSerialInteger(ByVal TableName As String, _
        ByVal ColumnName As String) As Integer
            If IsEmptyTable(TableName) Then
                Return 1
            Else
                Return Int("SELECT MAX(" + ColumnName + ") FROM " + TableName) + 1
            End If
        End Function

        ''' <summary>
        ''' 获取SQL语句查询结果对应的序列化后的字符串。
        ''' </summary>
        ''' <param name="Command">SQL语句。</param>
        ''' <param name="SeparateChar" >不同数据列之间的分隔符。</param>
        ''' <param name="IsCrLf" >不同行之间是否包含回车换行。</param>
        ''' <returns>经过序列化后的字符串。</returns>
        ''' <remarks></remarks>
        Public Shared Function SerializabledString(ByVal Command As String, _
                                                   Optional ByVal SeparateChar As Char = " "c, _
                                                   Optional ByVal IsCrLf As Boolean = True) As String
            Dim mt As Matrix(Of String) = New Matrix(Of String)(Command)
            Dim s As String = ""
            Dim row As Integer = mt.RowsCount, column As Integer = mt.ColumnsCount
            For i As Integer = 1 To row
                For j As Integer = 1 To column - 1
                    s = s + mt.Cell(i, j) + SeparateChar
                Next
                s += mt.Cell(i, column)
                If i <> row Then
                    If IsCrLf Then
                        s += vbCrLf
                    Else
                        s += SeparateChar
                    End If
                End If
            Next
            mt.Clear()
            Return s
        End Function

    End Class
End Namespace
