Imports System.Data

Namespace db
    Friend Class DBExpert
        Friend Shared Function IndexOfColumnName(ByRef dt As DataTable, Name As String) As Integer
            Dim cnt As Integer = dt.Columns.Count - 1
            For i As Integer = 0 To cnt
                If dt.Columns(i).Caption = Name Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Friend Shared Function Split(ByRef dt As DataTable, ColumnIndex As Integer) As Dictionary(Of String, DataTable)
            Dim rowCount As Integer = dt.Rows.Count - 1
            Dim s As String = Nothing
            Dim dict As Dictionary(Of String, DataTable) = New Dictionary(Of String, DataTable)
            Dim subDt As DataTable = Nothing
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
            subDt = Nothing
            rowCount = Nothing
            s = Nothing
            Return dict
        End Function

    End Class
End Namespace

