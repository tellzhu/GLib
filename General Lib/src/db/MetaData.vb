Imports dotNet.db.Admin
Imports dotNet.db.Value
Imports dotNet.math

Namespace db

    Public Class MetaData

        Private Shared m_DataSourceOfRecordset As String = Nothing

        ''' <summary>
        ''' 直通型数据库服务器连接字符串属性。
        ''' </summary>
        ''' <value>字符串类型的直通型数据库服务器连接。</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Shared Property DataSourceOfRecordset As String
            Get
                Return m_DataSourceOfRecordset
            End Get
            Set(ByVal value As String)
                m_DataSourceOfRecordset = value
            End Set
        End Property

        Private Shared m_CurrentDBType As DBType = DBType.DB2

        ''' <summary>
        ''' 数据库服务器类型，支持IBM DB2、MS SQLServer、Oracle、MS Access、ODBC、DBF、CSV、TEXT和Excel数据库。
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum DBType
            DB2
            SQLSERVER
            ORACLE
            ACCESS
            ODBC
            DBF
            CSV
            EXCEL
            TEXT
        End Enum

        ''' <summary>
        ''' 数据库服务器类型属性。
        ''' </summary>
        ''' <value>DBType类型的数据库服务器类型值。</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Shared Property DatabaseType() As DBType
            Get
                Return m_CurrentDBType
            End Get
            Set(ByVal value As DBType)
                m_CurrentDBType = value
            End Set
        End Property

        Private Shared tempTSpace As String
        Public Shared Property TempTableSpace() As String
            Get
                Return tempTSpace
            End Get
            Set(ByVal value As String)
                tempTSpace = value
            End Set
        End Property

        Private Shared m_DataSource As String = Nothing
        ''' <summary>
        ''' 数据库查询连接字符串属性。
        ''' </summary>
        ''' <value>字符串类型的数据库查询连接字符串。</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Shared Property DataSource() As String
            Get
                Return m_DataSource
            End Get
            Set(ByVal value As String)
                m_DataSource = value
            End Set
        End Property

        Private Shared m_IPAddress As String = Nothing
        Friend Shared ReadOnly Property IPAddress As String
            Get
                Return m_IPAddress
            End Get
        End Property

        Private Shared m_ProxyServer As String = Nothing
        Private Shared m_ProxyServerPort As Integer = Nothing
        Private Shared m_ProxyIdentifyingCode As String = Nothing
        Private Shared m_IsProxy As Boolean = False

        ''' <summary>
        ''' 设置数据库代理服务器的基本信息。
        ''' </summary>
        ''' <param name="IPAddress">代理服务器的IP地址。</param>
        ''' <param name="Port">代理服务器所用端口。</param>
        ''' <param name="IdentifyingCode">代理服务器认证代码。</param>
        Public Shared Sub SetProxy(IPAddress As String, Port As Integer, IdentifyingCode As String)
            m_ProxyServer = IPAddress
            m_ProxyServerPort = Port
            m_ProxyIdentifyingCode = IdentifyingCode
            m_IsProxy = (IPAddress <> Nothing) And (Port >= 0)
        End Sub

        ''' <summary>
        ''' 设置数据库连接的基本属性信息。
        ''' </summary>
        ''' <param name="DatabaseType">数据库类型，目前包括SQLServer、Access、DB2、Oracle和ODBC数据库。</param>
        ''' <param name="IPAddress">数据库服务器的IP地址。</param>
        ''' <param name="DBName">数据库名称。</param>
        ''' <param name="UserName">数据库用户名称。</param>
        ''' <param name="Password">数据库用户密码。</param>
        ''' <remarks></remarks>
        Public Shared Sub SetDBMetaData(ByVal DatabaseType As DBType, ByVal IPAddress As String,
                                        ByVal DBName As String, ByVal UserName As String,
                                        ByVal Password As String)
            m_CurrentDBType = DatabaseType
            m_IPAddress = IPAddress
            Select Case DatabaseType
                Case DBType.SQLSERVER
                    m_DataSource = "Server=" + IPAddress + ";Database=" _
                        + DBName + ";User Id=" + UserName + ";Password=" + Password + ";Connection Timeout=0;"
                    m_DataSourceOfRecordset = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" _
                        + UserName + ";" + "Password=" + Password + ";Initial Catalog=" _
                        + DBName + ";Data Source=" + IPAddress
                Case DBType.DB2
                    m_DataSource = "Server=" + IPAddress + ":50000;Database=" _
                        + DBName + ";UID=" + UserName + ";PWD=" + Password + ";"
                    m_DataSourceOfRecordset = Nothing
                Case DBType.ACCESS
                    If IPAddress <> Nothing Then
                        MsgBox("Error Database Type")
                        Return
                    End If
                    If UserName <> Nothing Then
                        MsgBox("Error Database Type")
                        Return
                    End If
                    m_DataSource = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" + DBName _
                        + """;Jet OLEDB:Database Password=""" + Password + """"
                    m_DataSourceOfRecordset = Nothing
                Case DBType.ORACLE
                    m_DataSource = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + IPAddress _
                        + ")(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + DBName + ")));User ID=" + UserName _
                        + ";Password=" + Password
                    m_DataSourceOfRecordset = Nothing
                Case Else
                    MsgBox("Error Database Type")
            End Select
            CloseCurrentConnections()
        End Sub

        Private Shared Function IsTable(ByVal TableOrViewName As String) As Boolean
            Select Case m_CurrentDBType
                Case DBType.SQLSERVER
                    Dim index As Integer = TableOrViewName.IndexOf(".")
                    If index <> -1 Then
                        Return Str("SELECT Ltrim(Rtrim(type)) FROM sys.all_objects T1 INNER JOIN sys.schemas T2 " _
                                   + "ON T1.schema_id=T2.schema_id WHERE T1.name='" + TableOrViewName.Substring(index + 1) _
                                   + "' AND T2.name='" + TableOrViewName.Substring(0, index) + "'") = "U"
                    Else
                        Return Str("SELECT Ltrim(Rtrim(type)) FROM sys.all_objects WHERE name='" + TableOrViewName + "'") = "U"
                    End If
                Case Else
                    MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                    Return Nothing
            End Select
        End Function

        Public Shared Function GetCreateTableDDL(TableName As String, DatabaseType As DBType) As String
            If m_CurrentDBType = DBType.ORACLE And DatabaseType = DBType.SQLSERVER Then
                Dim ddl As String = "DROP TABLE " + TableName + ";" + vbCrLf _
                                    + "CREATE TABLE " + TableName + "(" + vbCrLf
                Dim index As Integer = TableName.IndexOf(".")
                If index <> -1 Then
                    TableName = TableName.Substring(index + 1)
                End If
                Dim mt As Matrix(Of Object) = New Matrix(Of Object)("select column_name,data_type,data_length,data_precision,data_scale,nullable " _
                                    + "from user_tab_columns where table_name='" + TableName.ToUpper + "' order by column_id")
                index = mt.RowsCount
                For i As Integer = 1 To index
                    ddl += CStr(mt.Cell(i, 1)) + " "
                    Select Case CStr(mt.Cell(i, 2))
                        Case "DATE"
                            ddl += "DATETIME"
                        Case "VARCHAR2"
                            ddl += "VARCHAR(" + CStr(mt.Cell(i, 3)) + ")"
                        Case "NUMBER"
                            If mt.Cell(i, 4) Is Nothing And CInt(mt.Cell(i, 5)) = 0 Then
                                ddl += "BIGINT"
                            Else
                                ddl += "DECIMAL(" + CStr(mt.Cell(i, 4)) + "," + CStr(mt.Cell(i, 5)) + ")"
                            End If
                        Case "CHAR"
                            ddl += "CHAR(" + CStr(mt.Cell(i, 3)) + ")"
                        Case "FLOAT"
                            ddl += "FLOAT"
                        Case Else
                            MsgBox("Error Data Type: " + CStr(mt.Cell(i, 2)))
                            Exit For
                    End Select
                    If CStr(mt.Cell(i, 6)) = "N" Then
                        ddl += " NOT NULL"
                    End If
                    If i = index Then
                        ddl += ");" + vbCrLf
                    Else
                        ddl += "," + vbCrLf
                    End If
                Next
                mt.Clear()
                Return ddl
            End If
            Return Nothing
        End Function


        ''' <summary>
        ''' 获取数据库表或视图中所有列的数据类型名称。
        ''' </summary>
        ''' <param name="TableOrViewName">指定的数据库表或视图的名称。</param>
        ''' <value></value>
        ''' <returns>一个建立“列名称->数据类型名称”对应关系的字典。</returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property ColumnTypes(ByVal TableOrViewName As String) As Dictionary(Of String, String)
            Get
                Select Case m_CurrentDBType
                    Case DBType.SQLSERVER
                        Dim schemaTableName As String = "views"
                        If IsTable(TableOrViewName) Then
                            schemaTableName = "tables"
                        End If
                        Dim index As Integer = TableOrViewName.IndexOf(".")
                        Dim sql As String
                        If index = -1 Then
                            sql = "SELECT T2.name,UPPER(T3.name) FROM sys." + schemaTableName + " T1 INNER JOIN sys.columns T2 " _
                              + "ON T1.object_id=T2.object_id INNER JOIN sys.types T3 ON T2.system_type_id=T3.system_type_id" _
                              + " WHERE T1.name='" + TableOrViewName + "' AND T3.name<>'sysname'"
                        Else
                            sql = "SELECT T2.name,UPPER(T3.name) FROM sys." + schemaTableName + " T1 INNER JOIN sys.columns T2 " _
  + "ON T1.object_id=T2.object_id INNER JOIN sys.types T3 ON T2.system_type_id=T3.system_type_id " _
                            + " INNER JOIN sys.schemas T4 ON T1.schema_id=T4.schema_id " _
  + " WHERE T1.name='" + TableOrViewName.Substring(index + 1) + "' AND T4.name='" _
  + TableOrViewName.Substring(0, index) + "' AND T3.name<>'sysname'"
                        End If
                        Return Pair(sql)
                    Case DBType.ORACLE
                        Dim index As Integer = TableOrViewName.IndexOf(".")
                        If index <> -1 Then
                            TableOrViewName = TableOrViewName.Substring(index + 1)
                        End If
                        Dim sql As String = "select COLUMN_NAME,DATA_TYPE from user_tab_columns " _
                                            + "where table_name='" + TableOrViewName.ToUpper + "'"
                        Return Pair(sql)
                    Case Else
                        MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                        Return Nothing
                End Select
            End Get
        End Property

        ''' <summary>
        ''' 获取数据库表或视图的所有列的名称。
        ''' </summary>
        ''' <param name="TableOrViewName">指定的数据库表或视图的名称。</param>
        ''' <value></value>
        ''' <returns>一个将表或视图的列名称存储在其中的字符串数组。</returns>
        ''' <remarks>列的名称按其在表或视图中的顺序排列。</remarks>
        Public Shared ReadOnly Property Columns(ByVal TableOrViewName As String) As String()
            Get
                Select Case m_CurrentDBType
                    Case DBType.DB2
                        Dim index As Integer = TableOrViewName.IndexOf(".")
                        Return CType(Execute("SELECT COLNAME FROM " _
                          + "SYSCAT.COLUMNS WHERE TABSCHEMA='" + TableOrViewName.Substring(0, index) _
                          + "' AND TABNAME='" + TableOrViewName.Substring(index + 1) + "' ORDER BY COLNO") _
                          .ToArray(GetType(String)), String())
                    Case DBType.SQLSERVER
                        Dim schemaTable As String = "views"
                        If IsTable(TableOrViewName) Then
                            schemaTable = "tables"
                        End If
                        Dim index As Integer = TableOrViewName.IndexOf(".")
                        Dim schemaName As String = "dbo"
                        If index <> -1 Then
                            schemaName = TableOrViewName.Substring(0, index)
                            TableOrViewName = TableOrViewName.Substring(index + 1)
                        End If
                        Return ExecuteType(Of String)("SELECT T2.name FROM sys." + schemaTable + " T1 " _
                                             + "INNER JOIN sys.columns T2 ON T1.object_id=T2.object_id " _
                                             + "INNER JOIN sys.schemas T3 ON T1.schema_id=T3.schema_id " _
                                             + "WHERE T1.name='" + TableOrViewName + "' AND T3.name='" _
                                             + schemaName + "' ORDER BY column_id").ToArray()
                    Case DBType.ORACLE
                        Dim index As Integer = TableOrViewName.IndexOf(".")
                        If index <> -1 Then
                            TableOrViewName = TableOrViewName.Substring(index + 1)
                        End If
                        Return ExecuteType(Of String)("select column_name from user_tab_columns where table_name='" + TableOrViewName.ToUpper _
                                               + "' order by column_id").ToArray()
                    Case Else
                        MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                        Return Nothing
                End Select
            End Get
        End Property

        Public Shared ReadOnly Property TableName(ByVal ViewName As String) As String
            Get
                Select Case m_CurrentDBType
                    Case DBType.DB2
                        Dim index As Integer = ViewName.IndexOf(".")
                        Return CStr(Execute("SELECT RTRIM(BSCHEMA)||'.'||BNAME FROM SYSCAT.VIEWDEP WHERE " _
                        + "VIEWSCHEMA='" + ViewName.Substring(0, index) + "' AND VIEWNAME='" _
                       + ViewName.Substring(index + 1) + "'").Item(0))
                    Case Else
                        MsgBox("Error Execution Path.", MsgBoxStyle.Critical)
                        Return Nothing
                End Select
            End Get
        End Property
    End Class
End Namespace
