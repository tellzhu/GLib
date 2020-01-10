Imports System.Data

Namespace db
    Public Class DBExpert
        Friend Shared Function Split(Of T)(ByRef dt As DataTable, ColumnIndex As Integer) As Dictionary(Of T, DataTable)
            Dim rowCount As Integer = dt.Rows.Count - 1
            Dim s As T
            Dim dict As Dictionary(Of T, DataTable) = New Dictionary(Of T, DataTable)
            Dim subDt As DataTable
            For i As Integer = 0 To rowCount
                s = CType(dt.Rows(i)(ColumnIndex), T)
                If dict.ContainsKey(s) Then
                    dict.Item(s).Rows.Add(dt.Rows(i).ItemArray)
                Else
                    subDt = dt.Clone
                    subDt.Rows.Add(dt.Rows(i).ItemArray)
                    dict.Add(s, subDt)
                End If
            Next
            Return dict
        End Function

        Private Shared Function Filter(Of T)(dt As DataTable, ColumnIndex As Integer, ByRef FilterSet As HashSet(Of T)) As DataTable
            Dim rowCount As Integer = dt.Rows.Count - 1
            Dim s As T
            Dim subDt As DataTable = dt.Clone
            For i As Integer = 0 To rowCount
                s = CType(dt.Rows(i)(ColumnIndex), T)
                If FilterSet.Contains(s) Then
                    subDt.Rows.Add(dt.Rows(i).ItemArray)
                End If
            Next
            Return subDt
        End Function

        Friend Shared Function Filter(Of T)(dt As DataTable, ColumnName As String, ByRef FilterSet As HashSet(Of T)) As DataTable
            Dim columnIndex As Integer = dt.Columns.IndexOf(ColumnName)
            If columnIndex = -1 Then
                Return Nothing
            End If
            Return Filter(Of T)(dt, columnIndex, FilterSet)
        End Function

        Friend Shared Function ConvertListToSet(Of T)(ByRef lst As List(Of T)) As HashSet(Of T)
            Dim hs As HashSet(Of T) = New HashSet(Of T)
            Dim cnt As Integer = lst.Count - 1
            For i As Integer = 0 To cnt
                If Not hs.Contains(lst(i)) Then
                    hs.Add(lst(i))
                End If
            Next
            Return hs
        End Function

        Private Shared Function Filter(Of T)(ByRef dt As DataTable, ColumnIndex As Integer, ByRef FilterList As List(Of T)) As DataTable
            Return Filter(Of T)(dt, ColumnIndex, ConvertListToSet(Of T)(FilterList))
        End Function

        Private Shared Function ConvertDict(Of T, V)(ByRef dict As Dictionary(Of T, DataTable), ByRef MappingDict As Dictionary(Of T, V)) As Dictionary(Of V, DataTable)
            Dim newDict As Dictionary(Of V, DataTable) = New Dictionary(Of V, DataTable)
            Dim val As V
            For Each key As T In dict.Keys
                If MappingDict.ContainsKey(key) Then
                    val = MappingDict(key)
                    If newDict.ContainsKey(val) Then
                        newDict(val).Merge(dict(key))
                    Else
                        newDict.Add(val, dict(key).Copy)
                    End If
                End If
            Next
            Return newDict
        End Function

        Public Shared Function Split(Of T, V)(ByRef dt As DataTable, ColumnIndex As Integer, ByRef MappingDict As Dictionary(Of T, V)) As Dictionary(Of V, DataTable)
            If dt Is Nothing Or MappingDict Is Nothing Then
                Return Nothing
            End If
            Dim dict As Dictionary(Of T, DataTable) = Split(Of T)(dt, ColumnIndex)
            Dim newDict As Dictionary(Of V, DataTable) = ConvertDict(dict, MappingDict)
            dict.Clear()
            Return newDict
        End Function

        Friend Shared Function ContainsData(Of T)(ByRef dt As DataTable, ColumnName As String, ColumnValue As T) As Boolean
            If dt Is Nothing Then
                Return False
            End If
            Dim columnIndex As Integer = dt.Columns.IndexOf(ColumnName)
            If columnIndex = -1 Then
                Return False
            End If
            Dim rowCount As Integer = dt.Rows.Count - 1
            Dim s As T
            For i As Integer = 0 To rowCount
                s = CType(dt.Rows(i)(columnIndex), T)
                If s.Equals(ColumnValue) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Shared Sub AddColumnToTable(ByRef dt As DataTable, ByRef MappingDict As Dictionary(Of String, String), KeyColumnName As String, ValueColumnName As String, Offset As Integer)
            If MappingDict Is Nothing Or dt Is Nothing Then
                Return
            End If
            Dim keyColumnIndex As Integer = dt.Columns.IndexOf(KeyColumnName)
            If keyColumnIndex = -1 Then
                Return
            End If
            If Offset < 0 And -Offset > keyColumnIndex + 1 Or Offset > 0 And Offset > dt.Columns.Count - keyColumnIndex Then
                Return
            End If
            If dt.Columns.Contains(ValueColumnName) Then
                Return
            End If
            Dim valueColumnIndex As Integer
            If Offset <> 0 Then
                dt.Columns.Add(ValueColumnName, ValueColumnName.GetType)
                If Offset > 0 Then
                    valueColumnIndex = keyColumnIndex + Offset
                Else
                    valueColumnIndex = keyColumnIndex + Offset + 1
                    keyColumnIndex += 1
                End If
                dt.Columns(ValueColumnName).SetOrdinal(valueColumnIndex)
            Else
                dt.Columns(keyColumnIndex).ColumnName = ValueColumnName
            End If
            Dim s As String
            For i As Integer = 0 To dt.Rows.Count - 1
                s = CStr(dt.Rows(i).Item(keyColumnIndex))
                If MappingDict.ContainsKey(s) Then
                    If Offset <> 0 Then
                        dt.Rows(i).Item(valueColumnIndex) = MappingDict(s)
                    Else
                        dt.Rows(i).Item(keyColumnIndex) = MappingDict(s)
                    End If
                Else
                    If Offset <> 0 Then
                        dt.Rows(i).Item(valueColumnIndex) = DBNull.Value
                    End If
                End If
            Next
        End Sub
    End Class
End Namespace