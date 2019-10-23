Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports System.Threading.Thread
Imports IBM.Data.DB2
Imports Microsoft.Office.Interop.Excel
Imports dotNet.db.MetaData
Imports dotNet.db.Value

Namespace db

    Public Class Admin

        Private Structure GeneralConnection
            Dim ADODBRowCount As Integer
            Dim ADODBConnection As ADODB.Connection
            Dim ADODBRecordset As ADODB.Recordset
            Dim DBMSConnection As DbConnection
            Dim DBMSCommand As DbCommand

            Friend ReadOnly Property IsEmptyDBMS As Boolean
                Get
                    Return DBMSConnection Is Nothing Or DBMSCommand Is Nothing
                End Get
            End Property

            Friend Function Execute(Command As String) As ArrayList
                Me.DBMSCommand.CommandText = Command
                Dim reader As DbDataReader = Me.DBMSCommand.ExecuteReader
                Dim array As ArrayList = New ArrayList()
                Dim FieldCount As Integer = reader.FieldCount
                Do While reader.Read()
                    For i As Integer = 0 To FieldCount - 1
                        array.Add(reader.GetValue(i))
                    Next
                Loop
                reader.Close()
                RecordQueryFieldCount(FieldCount)
                If array.Count > 0 Then
                    Return array
                Else
                    Return Nothing
                End If
            End Function

            Friend Function ExecuteType(Of T)(Command As String) As List(Of T)
                Me.DBMSCommand.CommandText = Command
                Dim reader As DbDataReader = Me.DBMSCommand.ExecuteReader
                Dim array As List(Of T) = New List(Of T)
                Dim FieldCount As Integer = reader.FieldCount
                Dim o As Object
                Do While reader.Read()
                    For i As Integer = 0 To FieldCount - 1
                        o = reader.GetValue(i)
                        If IsDBNull(o) Then
                            array.Add(Nothing)
                        Else
                            array.Add(CType(o, T))
                        End If
                    Next
                Loop
                reader.Close()
                RecordQueryFieldCount(FieldCount)
                If array.Count > 0 Then
                    Return array
                Else
                    Return Nothing
                End If
            End Function
        End Structure

        Private Shared m_ConnectionSet As Dictionary(Of Integer, GeneralConnection) = New Dictionary(Of Integer, GeneralConnection)
        Private Shared m_FieldCountSet As Dictionary(Of Integer, Integer) = New Dictionary(Of Integer, Integer)

        Friend Shared Sub CloseCurrentConnections()
            If m_ConnectionSet.Count > 0 Then
                For Each con As GeneralConnection In m_ConnectionSet.Values
                    CloseDBConnection(con)
                    con = Nothing
                Next
                m_ConnectionSet.Clear()
            End If
        End Sub

        Private Shared Sub OpenADODBConnection(ByRef con As GeneralConnection, ByVal Command As String)
            con.ADODBConnection = New ADODB.Connection
            con.ADODBConnection.Open(DataSourceOfRecordset)
            con.ADODBConnection.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            con.ADODBRecordset = con.ADODBConnection.Execute(Command)
            con.ADODBRowCount = con.ADODBRecordset.RecordCount
    End Sub

    Friend Shared Function LoadCommand(ByVal Command As String) As Integer
        CloseRecordsetResource()
        Dim tId As Integer = CurrentThread.ManagedThreadId
            Dim con As GeneralConnection
            If Not m_ConnectionSet.ContainsKey(tId) Then
            con = New GeneralConnection
            OpenADODBConnection(con, Command)
            m_ConnectionSet.Add(tId, con)
        Else
            con = m_ConnectionSet.Item(tId)
            If con.ADODBConnection Is Nothing Or con.ADODBRecordset Is Nothing Then
                OpenADODBConnection(con, Command)
            End If
            m_ConnectionSet.Item(tId) = con
        End If
            Return con.ADODBRowCount
        End Function

    Private Shared Sub CloseRecordsetResource()
        Dim tId As Integer = CurrentThread.ManagedThreadId
        If m_ConnectionSet.ContainsKey(tId) Then
            Dim con As GeneralConnection = m_ConnectionSet.Item(tId)
            If con.ADODBRecordset IsNot Nothing Then
                con.ADODBRecordset.Close()
                con.ADODBRecordset = Nothing
            End If
            If con.ADODBConnection IsNot Nothing Then
                con.ADODBConnection.Close()
                con.ADODBConnection = Nothing
            End If
            m_ConnectionSet.Item(tId) = con
        End If
        tId = Nothing
    End Sub

    Friend Shared Function CopyFromRecordset(ByRef R As Range) As Integer
        Dim tId As Integer = CurrentThread.ManagedThreadId
        If Not m_ConnectionSet.ContainsKey(tId) Then
            Return 0
        End If
        Dim con As GeneralConnection = m_ConnectionSet.Item(tId)
        If con.ADODBConnection Is Nothing Or con.ADODBRecordset Is Nothing Then
            CloseRecordsetResource()
            Return 0
        End If
        Dim len As Integer = R.CopyFromRecordset(con.ADODBRecordset)
        con.ADODBRecordset.Close()
        con.ADODBRecordset = Nothing
        con.ADODBConnection.Close()
        con.ADODBConnection = Nothing
        m_ConnectionSet.Item(tId) = con
        tId = Nothing
        Return len
    End Function

    ''' <summary>
    ''' 判断给定SQL查询条件后的数据库表是否存在结果集。
    ''' </summary>
    ''' <param name="TableName">数据库表名称。</param>
    ''' <param name="Condition">SQL查询条件。</param>
    ''' <returns>若查询结果存在结果集，则返回False；否则返回True。</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmptyTable(ByVal TableName As String, Optional ByVal Condition As String = Nothing) As Boolean
        Return Count(TableName, Condition) = 0
    End Function

    Public Shared Sub Aggregate(ByVal TableName As String, ByVal TargetColumn As String, _
  ByVal TargetValue As String, ByVal Groups As String, Optional ByVal Condition As String = Nothing)
        Execute("DELETE FROM " + TableName + " WHERE " _
        + TargetColumn + "=" + TargetValue)

        If IsEmptyTable(TableName, Condition) Then
            Return
        End If

        Dim group As String() = Split(Groups, ",")
        Dim command As String = "INSERT INTO " + TableName + " SELECT "
        Dim cols As String() = Columns(TableName)
        For i As Integer = 0 To cols.Length - 1
            If TargetColumn = cols(i) Then
                command += TargetValue
            ElseIf Array.IndexOf(group, cols(i)) <> -1 Then
                command += cols(i)
            Else
                command = command + "SUM(" + cols(i) + ")"
            End If
            If i = cols.Length - 1 Then
                command += " FROM " + TableName
                If Condition IsNot Nothing Then
                    command += " WHERE " + Condition
                End If
                If Groups <> "" Then
                    command += " GROUP BY " + Groups
                End If
            Else
                command += ","
            End If
        Next
        group = Nothing
        cols = Nothing

        Execute(command)
        command = Nothing
    End Sub

    Public Shared Sub DropTempTable(ByVal TableName As String)
        Execute("DROP TABLE SESSION." + TableName)
    End Sub

    Private Shared Sub RenameObject(OldName As String, NewName As String)
        Select Case DatabaseType
            Case DBType.SQLSERVER
                Execute("sp_rename '" + OldName + "','" _
                        + NewName.Substring(NewName.IndexOf(".") + 1) + "'")
        End Select
    End Sub

    ''' <summary>
    ''' 重命名一个数据表。
    ''' </summary>
    ''' <param name="OldTableName">原来的数据表名。</param>
    ''' <param name="NewTableName">新的数据表名。</param>
    ''' <remarks></remarks>
    Public Shared Sub RenameTable(OldTableName As String, NewTableName As String)
        Select Case DatabaseType
            Case DBType.SQLSERVER
                If OldTableName.IndexOf(".") = -1 Then
                    OldTableName = "dbo." + OldTableName
                End If
                If NewTableName.IndexOf(".") = -1 Then
                    NewTableName = "dbo." + NewTableName
                End If
                Dim s1 As String = OldTableName.Substring(0, OldTableName.IndexOf("."))
                Dim s2 As String = NewTableName.Substring(0, NewTableName.IndexOf("."))
                If s1 <> s2 Then
                    s1 = Nothing
                    s2 = Nothing
                    Return
                End If
                RenameObject(OldTableName, NewTableName)
                s1 = Nothing
                s2 = Nothing
        End Select
    End Sub

    ''' <summary>
    ''' 重命名一个数据视图。
    ''' </summary>
    ''' <param name="OldViewName">原来的数据视图名。</param>
    ''' <param name="NewViewName">新的数据视图名。</param>
    ''' <remarks></remarks>
    Public Shared Sub RenameView(OldViewName As String, NewViewName As String)
        Select Case DatabaseType
            Case DBType.SQLSERVER
                If OldViewName.IndexOf(".") = -1 Then
                    OldViewName = "dbo." + OldViewName
                End If
                If NewViewName.IndexOf(".") = -1 Then
                    NewViewName = "dbo." + NewViewName
                End If
                Dim s1 As String = OldViewName.Substring(0, OldViewName.IndexOf("."))
                Dim s2 As String = NewViewName.Substring(0, NewViewName.IndexOf("."))
                If s1 <> s2 Then
                    s1 = Nothing
                    s2 = Nothing
                    Return
                End If
                RenameObject(OldViewName, NewViewName)
                s1 = Nothing
                s2 = Nothing
        End Select
    End Sub

    ''' <summary>
        ''' 清空数据库表中的数据。
    ''' </summary>
        ''' <param name="TableName">数据库表的名称。</param>
        ''' <param name="Condition">用于清空数据的筛选条件。</param>
        ''' <remarks>若筛选条件为空，则清空数据表中的所有数据。</remarks>
        Public Shared Sub EmptyTable(ByVal TableName As String, Optional ByVal Condition As String = Nothing)
            If Condition = Nothing Then
                Select Case DatabaseType
                    Case DBType.DB2
                        Execute("ALTER TABLE " + TableName + " ACTIVATE NOT LOGGED INITIALLY WITH EMPTY TABLE")
                    Case DBType.SQLSERVER, DBType.ORACLE
                        Execute("TRUNCATE TABLE " + TableName)
                    Case Else
                        MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                End Select
            Else
                Select Case DatabaseType
                    Case DBType.DB2, DBType.SQLSERVER, DBType.ORACLE
                        Execute("DELETE FROM " + TableName + " WHERE " + Condition)
                    Case Else
                        MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                End Select
            End If
        End Sub

    Friend Shared Sub LoadBulkData(FileName As String, TableName As String)
        Select Case DatabaseType
            Case DBType.SQLSERVER
                Execute("BULK INSERT " + TableName + " FROM '" _
                        + FileName + "' WITH(FIELDTERMINATOR=',',ROWTERMINATOR='\n')")
            Case Else
                MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
        End Select
    End Sub


    Public Shared Sub CreateTempTable(ByVal TableName As String, ByVal Columns As String)
        Select Case DatabaseType
            Case DBType.DB2
                Execute("DECLARE GLOBAL TEMPORARY TABLE SESSION." _
                + TableName + " (" + Columns + ") ON COMMIT PRESERVE ROWS NOT LOGGED IN " + MetaData.TempTableSpace)
            Case Else
                MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
        End Select
    End Sub

    Friend Shared ReadOnly Property LastFieldCount As Integer
        Get
            Dim tId As Integer = CurrentThread.ManagedThreadId
            If m_FieldCountSet.ContainsKey(tId) Then
                Return m_FieldCountSet.Item(tId)
            Else
                tId = Nothing
                Return 0
            End If
            End Get
    End Property

    Private Shared m_ConnectionIsReused As Boolean = False
    ''' <summary>
    ''' 指明数据库适配器是否复用已经建立的数据库连接。
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>默认情况下，不复用已经建立的数据库连接。</remarks>
    Public Shared Property ConnectionIsReused As Boolean
        Get
            Return m_ConnectionIsReused
        End Get
        Set(ByVal value As Boolean)
            m_ConnectionIsReused = value
        End Set
    End Property

    Private Shared Sub CloseDBConnection(ByRef con As GeneralConnection)
        If con.DBMSCommand IsNot Nothing Then
            con.DBMSCommand = Nothing
        End If
        If con.DBMSConnection IsNot Nothing Then
            con.DBMSConnection.Close()
            con.DBMSConnection = Nothing
        End If
    End Sub

    Private Shared Sub RecordQueryFieldCount(ByVal FieldCount As Integer)
        Dim tId As Integer = CurrentThread.ManagedThreadId
        m_FieldCountSet.Item(tId) = FieldCount
        If Not m_ConnectionIsReused Then
            Dim con As GeneralConnection = m_ConnectionSet.Item(tId)
            CloseDBConnection(con)
            m_ConnectionSet.Item(tId) = Nothing
            m_ConnectionSet.Remove(tId)
            con = Nothing
        End If
        tId = Nothing
    End Sub

    Private Shared Sub InitDBConnection()
        Dim tId As Integer = CurrentThread.ManagedThreadId
        Dim con As GeneralConnection = Nothing
        If Not m_ConnectionSet.ContainsKey(tId) Then
            con = New GeneralConnection
            OpenDBConnection(con)
            m_ConnectionSet.Add(tId, con)
        Else
            con = m_ConnectionSet.Item(tId)
            If con.IsEmptyDBMS Then
                OpenDBConnection(con)
                m_ConnectionSet.Item(tId) = con
            End If
        End If
    End Sub

    Private Shared Sub OpenDBConnection(ByRef con As GeneralConnection)
        If con.IsEmptyDBMS Then
                Select Case DatabaseType
                    Case DBType.ODBC
                        con.DBMSConnection = New Data.Odbc.OdbcConnection(DataSource)
                        con.DBMSCommand = New OdbcCommand()
                    Case DBType.ACCESS
                        con.DBMSConnection = New Data.OleDb.OleDbConnection(DataSource)
                        con.DBMSCommand = New OleDbCommand()
                    Case DBType.SQLSERVER
                        con.DBMSConnection = New SqlConnection(DataSource)
                        con.DBMSCommand = New SqlCommand()
                    Case DBType.DB2
                        con.DBMSConnection = New DB2Connection(DataSource)
                        con.DBMSCommand = New DB2Command()
                    Case DBType.ORACLE
                        con.DBMSConnection = New OracleConnection(DataSource)
                        con.DBMSCommand = New OracleCommand()
                    Case Else
                        MsgBox("Error Database Type.", MsgBoxStyle.Critical)
                End Select
                con.DBMSCommand.Connection = con.DBMSConnection
                con.DBMSCommand.CommandTimeout = 0
                con.DBMSConnection.Open()
        End If
    End Sub

    Private Shared m_IsSQLTune As Boolean = False
    Public Shared Property IsSQLTune As Boolean
        Get
            Return m_IsSQLTune
        End Get
        Set(ByVal value As Boolean)
            m_IsSQLTune = value
        End Set
    End Property

    Private Shared m_IsLogAllSQL As Boolean = False
        Public Shared WriteOnly Property IsLogAllSQL As Boolean
            Set(ByVal value As Boolean)
                m_IsLogAllSQL = value
            End Set
        End Property

        Private Shared m_SQLLogFileName As String = My.Computer.FileSystem.SpecialDirectories.Desktop + "\SQL_Log.txt"
        Public Shared WriteOnly Property SQLLogFileName As String
            Set(ByVal value As String)
                m_SQLLogFileName = value
            End Set
        End Property

        Friend Shared Sub BulkCopy(ByRef sourceTable As Data.DataTable, destinationTable As String)
            Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(DataSource)
            bulkCopy.DestinationTableName = destinationTable
            bulkCopy.WriteToServer(sourceTable)
            bulkCopy.Close()
            bulkCopy = Nothing
        End Sub

        ''' <summary>
        ''' 运行一个SQL查询语句。
        ''' </summary>
        ''' <param name="Command">待运行的SQL语句。</param>
        ''' <returns>若SQL语句存在查询的结果集，则将结果集按从上至下、从左到右的
        ''' 顺序逐行、逐列输出并存储在一个ArrayList对象中；若SQL语句不存在查询结果集，
        ''' 则返回Nothing。
        ''' </returns>
        ''' <remarks></remarks>
        Public Shared Function Execute(ByVal Command As String) As ArrayList
            Dim tId As Integer = CurrentThread.ManagedThreadId
            If m_IsSQLTune Then
                If m_IsLogAllSQL Then
                    My.Computer.FileSystem.WriteAllText(m_SQLLogFileName, Now.ToString + " " + Command + vbCrLf, True)
                Else
                    If Command.Trim.ToUpper.IndexOf("SELECT") <> 0 Then
                        My.Computer.FileSystem.WriteAllText(m_SQLLogFileName, Now.ToString + " " + Command + vbCrLf, True)
                    End If
                End If
            End If
            InitDBConnection()
            Dim con As GeneralConnection = m_ConnectionSet.Item(tId)
            Return con.Execute(Command)
        End Function

        Private Shared Function GetDataAdapter(Command As String, ByRef con As GeneralConnection) As DbDataAdapter
            Select Case DatabaseType
                Case DBType.ODBC
                    Return New OdbcDataAdapter(Command, CType(con.DBMSConnection, Data.Odbc.OdbcConnection))
                Case DBType.ACCESS
                    Return New OleDbDataAdapter(Command, CType(con.DBMSConnection, Data.OleDb.OleDbConnection))
                Case DBType.SQLSERVER
                    Return New SqlDataAdapter(Command, CType(con.DBMSConnection, SqlConnection))
                Case DBType.DB2
                    Return New DB2DataAdapter(Command, CType(con.DBMSConnection, DB2Connection))
                Case DBType.ORACLE
                    Return New OracleDataAdapter(Command, CType(con.DBMSConnection, OracleConnection))
                Case Else
                    MsgBox("Error Database Type.", MsgBoxStyle.Critical)
                    Return Nothing
            End Select
        End Function
        Public Shared Function GetDataTable(Command As String) As Data.DataTable
            Dim tId As Integer = CurrentThread.ManagedThreadId
            InitDBConnection()
            Dim con As GeneralConnection = m_ConnectionSet.Item(tId)
            Dim adapter As DbDataAdapter = GetDataAdapter(Command, con)
            adapter.SelectCommand.CommandTimeout = 0
            Dim dt As Data.DataTable = New Data.DataTable
            adapter.Fill(dt)
            Return dt
        End Function

        Public Shared Function ExecuteType(Of T)(ByVal Command As String) As List(Of T)
            Dim tId As Integer = CurrentThread.ManagedThreadId
            If m_IsSQLTune Then
                If m_IsLogAllSQL Then
                    My.Computer.FileSystem.WriteAllText(m_SQLLogFileName, Now.ToString + " " + Command + vbCrLf, True)
                Else
                    If Command.Trim.ToUpper.IndexOf("Select") <> 0 Then
                        My.Computer.FileSystem.WriteAllText(m_SQLLogFileName, Now.ToString + " " + Command + vbCrLf, True)
                    End If
                End If
            End If
            InitDBConnection()
            Dim con As GeneralConnection = m_ConnectionSet.Item(tId)
            Return con.ExecuteType(Of T)(Command)
    End Function

    End Class
End Namespace
