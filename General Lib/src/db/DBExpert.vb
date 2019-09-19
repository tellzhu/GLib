Imports System.Data

Namespace db
    Public Class DBExpert

        Friend Shared Function Split(ByRef dt As DataTable, ColumnIndex As Integer) As Dictionary(Of String, DataTable)
            Dim rowCount As Integer = dt.Rows.Count - 1
            Dim s As String
            Dim dict As Dictionary(Of String, DataTable) = New Dictionary(Of String, DataTable)
            Dim subDt As DataTable
            For i As Integer = 0 To rowCount
                s = CStr(dt.Rows(i)(ColumnIndex))
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

        Public Shared Function Split(ByRef dt As DataTable, ColumnIndex As Integer, ByRef MappingDict As Dictionary(Of String, String)) As Dictionary(Of String, DataTable)
            If dt Is Nothing Or MappingDict Is Nothing Then
                Return Nothing
            End If
            Dim rowCount As Integer = dt.Rows.Count - 1
            Dim s As String
            Dim dict As Dictionary(Of String, DataTable) = New Dictionary(Of String, DataTable)
            Dim subDt As DataTable
            For i As Integer = 0 To rowCount
                s = CStr(dt.Rows(i)(ColumnIndex))
                If MappingDict.ContainsKey(s) Then
                    s = MappingDict(s)
                    If dict.ContainsKey(s) Then
                        dict(s).Rows.Add(dt.Rows(i).ItemArray)
                    Else
                        subDt = dt.Clone
                        subDt.Rows.Add(dt.Rows(i).ItemArray)
                        dict.Add(s, subDt)
                    End If
                End If
            Next
            Return dict
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

