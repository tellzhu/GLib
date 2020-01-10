Imports System.Data
Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports dotNet.db.Admin

Namespace db
    Public Class UIFiller

        ''' <summary>
        ''' 根据SQL查询语句结果填充DataGridView控件。
        ''' </summary>
        ''' <param name="DataGrid">待填充的DataGridView控件。</param>
        ''' <param name="Command">待运行的SQL查询语句。</param>
        ''' <remarks></remarks>
        Public Shared Sub FillDataGridView(ByRef DataGrid As DataGridView, ByVal Command As String)
            Dim dbAdapter As DbDataAdapter
            Select Case MetaData.DatabaseType
                Case MetaData.DBType.ODBC
                    dbAdapter = New OdbcDataAdapter(Command, MetaData.DataSource)
                Case MetaData.DBType.ACCESS
                    dbAdapter = New OleDbDataAdapter(Command, MetaData.DataSource)
                Case MetaData.DBType.SQLSERVER
                    dbAdapter = New SqlDataAdapter(Command, MetaData.DataSource)
                Case MetaData.DBType.ORACLE
                    dbAdapter = New OracleDataAdapter(Command, MetaData.DataSource)
                Case Else
                    MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                    Return
            End Select

            Dim dTable As DataTable = New DataTable()
            dbAdapter.Fill(dTable)
            DataGrid.DataSource = Nothing
            DataGrid.Columns.Clear()
            DataGrid.DataSource = dTable
        End Sub

        ''' <summary>
        ''' 根据SQL查询语句结果填充ComboBox控件。
        ''' </summary>
        ''' <param name="Box">待填充的ComboBox控件。</param>
        ''' <param name="Command">待运行的SQL查询语句。</param>
        ''' <param name="DefaultSelectedIndex">ComboBox控件默认被选中的项目索引序号。</param>
        ''' <remarks></remarks>
        Public Shared Sub FillComboBox(ByRef Box As ComboBox, ByVal Command As String, _
                                      Optional ByVal DefaultSelectedIndex As Integer = 0)
            Box.Items.Clear()
            Dim arr As List(Of String) = ExecuteType(Of String)(Command)
            If arr Is Nothing Then
                Return
            End If
            For i As Integer = 1 To arr.Count
                Box.Items.Add(arr.Item(i - 1))
            Next
            arr.Clear()
            If Box.Items.Count > DefaultSelectedIndex And DefaultSelectedIndex >= 0 Then
                Box.SelectedIndex = DefaultSelectedIndex
            ElseIf Box.Items.Count > 0 Then
                Box.SelectedIndex = 0
            End If
        End Sub
    End Class
End Namespace
