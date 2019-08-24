Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Data.OracleClient
Imports System.IO
Imports IBM.Data.DB2
Imports dotNet.db.Admin
Imports dotNet.db.MetaData
Imports dotNet.time.DateExpert

Namespace db
    Public Class DataMover

        Private Shared m_Separator As Char = Nothing
        Private Shared m_Comments As String = Nothing
        Private Shared m_EOF As String = Nothing

        Public Shared Sub SetSourceTextFile(Separator As Char, Comments As String, Optional EOF As String = Nothing)
            m_Separator = Separator
            m_Comments = Comments
            m_EOF = EOF
            m_SourceDbType = DBType.TEXT
        End Sub

        Public Shared Sub CloseSourceTextFile()
            m_Separator = Nothing
            m_Comments = Nothing
            m_EOF = Nothing
            m_SourceDbType = Nothing
        End Sub

        ''' <summary>
        ''' 准备源数据库以便进行数据传输。
        ''' </summary>
        ''' <param name="DatabaseType">源数据库类型，目前包括SQLServer、Access、DB2、ODBC、DBF、CSV和Excel数据库。</param>
        ''' <param name="IPAddress">源数据库服务器的IP地址。</param>
        ''' <param name="DBName">源数据库名称。</param>
        ''' <param name="UserName">源数据库用户名称。</param>
        ''' <param name="Password">源数据库用户密码。</param>
        ''' <remarks></remarks>
        Public Shared Sub OpenSourceDatabase(ByVal DatabaseType As DBType, ByVal IPAddress As String,
                                        ByVal DBName As String, ByVal UserName As String,
                                        ByVal Password As String)
            m_SourceDbType = DatabaseType
            Dim m_DataSourceString As String = Nothing
            Select Case DatabaseType
                Case DBType.DB2
                    m_DataSourceString = "Server=" + IPAddress + ":50000;Database=" _
                        + DBName + ";UID=" + UserName + ";PWD=" + Password + ";"
                    m_SourceDbConnection = New DB2Connection(m_DataSourceString)
                    m_SourceDbCommand = New DB2Command()
                    m_SourceDbCommand.Connection = m_SourceDbConnection
                    m_SourceDbCommand.CommandTimeout = 0
                Case DBType.ORACLE
                    m_DataSourceString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + IPAddress _
                        + ")(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + DBName + ")));User ID=" + UserName _
                        + ";Password=" + Password
                    m_SourceDbConnection = New OracleConnection(m_DataSourceString)
                    m_SourceDbCommand = New OracleCommand
                    m_SourceDbCommand.Connection = m_SourceDbConnection
                    m_SourceDbCommand.CommandTimeout = 0
                Case DBType.DBF
                    m_DataSourceString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + IPAddress + ";Extended Properties=dBASE IV;User ID=Admin;Password=;"
                    m_SourceDbConnection = New OleDbConnection(m_DataSourceString)
                Case DBType.CSV
                    m_DataSourceString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + IPAddress _
                        + "\;Extended Properties=""Text;HDR=yes;FMT=Delimited"""
                    m_SourceDbConnection = New OleDbConnection(m_DataSourceString)
                Case DBType.EXCEL
                    m_DataSourceString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName _
               + ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
                    'Provider = Microsoft.Jet.OLEDB.4.0;Data Source=D:\MyExcel.xls;Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1""
                    m_SourceDbConnection = New OleDbConnection(m_DataSourceString)
                Case Else
                    MsgBox("Error Source Database Type")
            End Select
            m_SourceDbConnection.Open()
            m_DataSourceString = Nothing
            ConnectionIsReused = True
        End Sub

        Private Shared m_SourceDbConnection As DbConnection = Nothing
        Private Shared m_SourceDbCommand As DbCommand = Nothing
        Private Shared m_SourceDbType As DBType = Nothing

        ''' <summary>
        ''' 关闭源数据库连接。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseSourceDatabase()
            If m_DictionaryOfColumnType IsNot Nothing Then
                m_DictionaryOfColumnType.Clear()
                m_DictionaryOfColumnType = Nothing
            End If
            m_InsertStatement = Nothing
            ConnectionIsReused = False
            m_SourceDbType = Nothing
            If m_SourceDbCommand IsNot Nothing Then
                m_SourceDbCommand = Nothing
            End If
            If m_SourceDbConnection IsNot Nothing Then
                m_SourceDbConnection.Close()
                m_SourceDbConnection = Nothing
            End If
        End Sub

        Private Shared m_TargetTableName As String = Nothing

        ''' <summary>
        ''' 设置数据传输的目标表的名称。
        ''' </summary>
        ''' <value>目标表的名称。</value>
        ''' <remarks></remarks>
        Public Shared WriteOnly Property TargetTableName As String
            Set(ByVal value As String)
                m_TargetTableName = value
                InitColumnTypes(value)
            End Set
        End Property

        Private Shared m_TransitDirectory As String = Nothing
        ''' <summary>
        ''' 设置或获取数据中转目录的名称。
        ''' </summary>
        ''' <value>预设置的中转目录名称。</value>
        ''' <returns>已设置的中转目录名称。</returns>
        ''' <remarks></remarks>
        Public Shared Property TransitDirectory As String
            Get
                Return m_TransitDirectory
            End Get
            Set(value As String)
                m_TransitDirectory = value
            End Set
        End Property

        ''' <summary>
        ''' 将查询语句的结果传输至目标数据表中。
        ''' </summary>
        ''' <param name="SelectCommand">将在源数据库中执行的查询语句，通常为SELECT SQL语句。</param>
        ''' <remarks></remarks>
        Public Shared Sub MoveFrom(ByVal SelectCommand As String)
            If m_SourceDbType = DBType.DB2 And DatabaseType = DBType.SQLSERVER Then
                MoveDataDB2ToSqlServer(SelectCommand)
            ElseIf m_SourceDbType = DBType.ORACLE And DatabaseType = DBType.SQLSERVER Then
                MoveDataOracleToSqlServer(SelectCommand)
            ElseIf (m_SourceDbType = DBType.DBF Or m_SourceDbType = DBType.EXCEL Or m_SourceDbType = DBType.CSV) And DatabaseType = DBType.ORACLE Then
                MoveDataDBFExcelCSVToOracle(SelectCommand)
            ElseIf m_SourceDbType = DBType.TEXT And DatabaseType = DBType.SQLSERVER Then
                MoveDataTextToSqlServer(SelectCommand)
            ElseIf m_SourceDbType = DBType.DBF And DatabaseType = DBType.SQLSERVER Then
                MoveDataDBFToSqlServer(SelectCommand)
            Else
                MsgBox("Error Move Data Type", MsgBoxStyle.Critical)
            End If
        End Sub

        Private Shared m_DictionaryOfColumnType As Dictionary(Of Integer, String) = Nothing
        Private Shared m_InsertStatement As String = Nothing

        Private Shared Function GetColumnsString(ByVal TableOrViewName As String) As String
            Dim cols As String() = Columns(TableOrViewName)
            Dim s As String = ""
            For i As Integer = 0 To cols.Length - 1
                s += cols(i) + ","
            Next
            Array.Clear(cols, 0, cols.Length)
            cols = Nothing
            Return s.Substring(0, s.Length - 1)
        End Function

        Private Shared Sub InitColumnTypes(ByVal TableName As String)
            Dim columnsName As String = GetColumnsString(TableName)
            m_InsertStatement = "INSERT INTO " + TableName + "(" + columnsName + ") VALUES("
            Dim dict As Dictionary(Of String, String) = ColumnTypes(TableName)
            Dim s As String() = Split(columnsName, ",")
            If m_DictionaryOfColumnType Is Nothing Then
                m_DictionaryOfColumnType = New Dictionary(Of Integer, String)
            Else
                m_DictionaryOfColumnType.Clear()
            End If
            For i As Integer = 0 To s.Length - 1
                m_DictionaryOfColumnType.Add(i, dict.Item(s(i)))
            Next
            Array.Clear(s, 0, s.Length)
            s = Nothing
            dict.Clear()
            dict = Nothing
            columnsName = Nothing
        End Sub

        Private Shared Sub MoveDataDBFToSqlServer(SelectCommand As String)
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter(SelectCommand, CType(m_SourceDbConnection, OleDbConnection))
            Dim dt As Data.DataTable = New Data.DataTable
            adapter.Fill(dt)
            adapter = Nothing
            BulkCopy(dt, m_TargetTableName)
            dt.Clear()
            dt = Nothing
        End Sub

        Private Shared Sub MoveDataDBFExcelCSVToOracle(SelectCommand As String)
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter(SelectCommand, CType(m_SourceDbConnection, OleDbConnection))
            Dim dt As Data.DataTable = New Data.DataTable
            adapter.Fill(dt)
            adapter = Nothing
            Dim maxColumn As Integer = dt.Columns.Count - 1
            Dim str As String = Nothing
            Dim maxRow As Integer = dt.Rows.Count - 1
            Dim obj As Object = Nothing
            For currRow As Integer = 0 To maxRow
                str = m_InsertStatement
                For currColumn As Integer = 0 To maxColumn
                    obj = dt.Rows(currRow).Item(currColumn)
                    Select Case m_DictionaryOfColumnType.Item(currColumn)
                        Case "VARCHAR2", "CHAR"
                            If IsDBNull(obj) Then
                                str += "''"
                            Else
                                str += "'" + CStr(obj).Trim.Replace(",", "；") + "'"
                            End If
                        Case "DATE"
                            If IsDate(obj) Then
                                str += "TO_DATE('" + CDate(obj).ToString("yyyyMMdd") + "','yyyymmdd')"
                            ElseIf CStr(obj).Length = 8 Then
                                str += "TO_DATE('" + CStr(obj) + "','yyyymmdd')"
                            Else
                                str += "NULL"
                            End If
                        Case "NUMBER", "FLOAT"
                            If IsDBNull(obj) Then
                                str += "NULL"
                            Else
                                str = str & CDec(obj)
                            End If
                        Case Else
                            MsgBox("Error Column Type.", MsgBoxStyle.Critical)
                    End Select
                    If currColumn < maxColumn Then
                        str += ","
                    Else
                        str += ")"
                    End If
                Next
                Execute(str)
            Next
            obj = Nothing
            dt.Clear()
            dt = Nothing
            maxColumn = Nothing
            maxRow = Nothing
            str = Nothing
        End Sub

        Private Shared Sub MoveDataTextToSqlServer(ByVal FileName As String)
            Dim table As Data.DataTable = New Data.DataTable
            Dim maxColumn As Integer = m_DictionaryOfColumnType.Count - 1
            For i As Integer = 0 To maxColumn
                table.Columns.Add("Column" & i)
            Next
            Dim sr As StreamReader = New StreamReader(FileName)
            Dim dr As Data.DataRow = Nothing
            Dim strLine As String = Nothing
            Dim strs() As String = Nothing
            While Not sr.EndOfStream
                strLine = sr.ReadLine
                If m_EOF <> Nothing And strLine = m_EOF Then
                    Exit While
                End If
                If strLine.StartsWith(m_Comments) Then
                    Continue While
                End If
                dr = table.NewRow()
                strs = Split(strLine, m_Separator)
                For i As Integer = 0 To maxColumn
                    Select Case m_DictionaryOfColumnType.Item(i)
                        Case "VARCHAR", "NVARCHAR"
                            dr(i) = strs(i)
                        Case "DATE"
                            If strs(i).Length = 8 Then
                                dr(i) = DateValue(strs(i))
                            End If
                        Case "DECIMAL"
                            If IsNumeric(strs(i)) Then
                                dr(i) = CDec(strs(i))
                            End If
                        Case Else
                            MsgBox("Error Column Type:" + m_DictionaryOfColumnType.Item(i) + " Value:" + strs(i), MsgBoxStyle.Critical)
                    End Select
                Next
                Array.Clear(strs, 0, maxColumn + 1)
                table.Rows.Add(dr)
            End While
            strs = Nothing
            strLine = Nothing
            dr = Nothing
            sr.Close()
            sr = Nothing
            BulkCopy(table, m_TargetTableName)
            table.Clear()
            table = Nothing
            maxColumn = Nothing
        End Sub

        Private Shared Sub MoveDataDB2ToSqlServer(ByVal SelectCommand As String)
            m_SourceDbCommand.CommandText = SelectCommand
            Dim sLoader As StringLoader = New StringLoader
            sLoader.DataTableName = m_TargetTableName
            Dim reader As DB2DataReader = CType(m_SourceDbCommand, DB2Command).ExecuteReader()
            Dim maxColumn As Integer = reader.FieldCount - 1
            Dim str As String = Nothing
            Do While reader.Read
                str = ""
                For i As Integer = 0 To maxColumn
                    Select Case m_DictionaryOfColumnType.Item(i)
                        Case "VARCHAR", "NVARCHAR"
                            If IsDBNull(reader.GetValue(i)) Then
                                str += " "
                            Else
                                str += reader.GetString(i).Trim.Replace(",", "；")
                            End If
                        Case "DATE"
                            If IsDBNull(reader.GetValue(i)) Then
                                str += ""
                            Else
                                str += CStr(reader.GetDate(i))
                            End If
                        Case "DECIMAL"
                            If IsDBNull(reader.GetValue(i)) Then
                                str += "0"
                            Else
                                str = str & reader.GetDecimal(i)
                            End If
                        Case "INT"
                            str = str & CInt(reader.GetValue(i))
                        Case Else
                            MsgBox("Error Column Type.", MsgBoxStyle.Critical)
                    End Select
                    If i < maxColumn Then
                        str += ","
                    End If
                Next
                sLoader.Append(str)
            Loop
            reader.Close()
            reader = Nothing
            sLoader.Load(m_TargetTableName)
            sLoader = Nothing
            maxColumn = Nothing
            str = Nothing
        End Sub

        Private Shared Sub MoveDataOracleToSqlServer(ByVal SelectCommand As String)
            m_SourceDbCommand.CommandText = SelectCommand
            Dim sLoader As StringLoader = New StringLoader
            sLoader.DataTableName = m_TargetTableName
            Dim reader As OracleDataReader = CType(m_SourceDbCommand, OracleCommand).ExecuteReader()
            Dim maxColumn As Integer = reader.FieldCount - 1
            Dim str As String = Nothing
            Do While reader.Read
                str = ""
                For i As Integer = 0 To maxColumn
                    If reader.IsDBNull(i) Then
                        str += ""
                    Else
                        Select Case m_DictionaryOfColumnType.Item(i)
                            Case "VARCHAR", "NVARCHAR"
                                str += reader.GetString(i).Trim.Replace(",", "；")
                            Case "DATE"
                                str += CStr(reader.GetDateTime(i))
                            Case "INT"
                                str = str & CInt(reader.GetValue(i))
                            Case "BIGINT"
                                str = str & CLng(reader.GetValue(i))
                            Case "DECIMAL"
                                str = str & CDec(reader.GetValue(i))
                            Case Else
                                MsgBox("Error Column Type:" + m_DictionaryOfColumnType.Item(i), MsgBoxStyle.Critical)
                        End Select
                    End If
                    If i < maxColumn Then
                        str += ","
                    End If
                Next
                sLoader.Append(str)
            Loop
            reader.Close()
            reader = Nothing
            sLoader.Load(m_TargetTableName)
            sLoader = Nothing
            maxColumn = Nothing
            str = Nothing
        End Sub
    End Class
End Namespace
