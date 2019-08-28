Imports Microsoft.Office.Interop.Excel
Imports dotNet.db.Admin
Imports dotNet.db.DBExpert
Imports dotNet.math
Imports dotNet.office.Excelist
Imports dotNet.office.FormulaBuilder

Namespace office
	Public Class AreaExpert

        ''' <summary>
        ''' ����Excel������ָ����Ԫ���ע�����ݡ�
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="comment">ע�͵��ı����ݡ�</param>
        ''' <remarks></remarks>
		Public Shared Sub setComment(ByRef a As Area, ByVal comment As String)
			a.setComment(comment)
		End Sub

		Public Shared Sub setFontUnderline(ByRef a As Area, Optional ByVal isUnderline As Boolean = True)
			a.setFontUnderline(isUnderline)
		End Sub

        ''' <summary>
        ''' ����Excel������ָ����Ԫ�������Ӵ���ʾ��ʽ��
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="isBold">�Ƿ�����Ӵ֡�</param>
        ''' <remarks></remarks>
		Public Shared Sub setFontBold(ByRef a As Area, Optional ByVal isBold As Boolean = True)
			a.setFontBold(isBold)
        End Sub

        Private Shared Sub PrintRecordSet(ByRef a As Area)
            a.PrintRecordSet()
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�����������һ��SQL��ѯ��������е�����Ԫ�ء������Excel�ĵ�Ԫ��
        ''' ���������������Ԫ����һһ��Ӧ�ģ����������Ԫ�������������������ֱ���ھ����������
        ''' �������ҵ�Ԫ���������Ͻǡ����½ǻ�����һ��λ�õĵ�Ԫ������������Ͻǡ����½ǻ�
        ''' ��Ӧλ�õ�Ԫ��һһ��Ӧ��
        ''' </summary>
        ''' <param name="a">ָ��Excel��Ԫ����������Ͻǵ�Ԫ��</param>
        ''' <param name="Command">��������ݵ�SQL��ѯ��䡣</param>
        ''' <returns>�������������</returns>
        ''' <remarks></remarks>
        Public Shared Function PrintRecordSet(ByRef a As Area, ByVal Command As String) As Integer
            Dim rowCnt As Integer = LoadCommand(Command)
            a.PrintRecordSet()
            Return rowCnt
        End Function

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�������а�Ԥ�����ʽ���һ��SQL��ѯ��������е�����Ԫ�ء������Excel�ĵ�Ԫ��
        ''' ���������������Ԫ����һһ��Ӧ�ģ����������Ԫ�������������������ֱ���ھ����������
        ''' �������ҵ�Ԫ���������Ͻǡ����½ǻ�����һ��λ�õĵ�Ԫ������������Ͻǡ����½ǻ�
        ''' ��Ӧλ�õ�Ԫ��һһ��Ӧ��
        ''' </summary>
        ''' <param name="a">ָ��Excel��Ԫ����������Ͻǵ�Ԫ��</param>
        ''' <param name="Command">��������ݵ�SQL��ѯ��䡣</param>
        ''' <param name="cf">Ԥ����ľ����е�Excel��Ԫ���ʽ��</param>
        ''' <returns>�������������</returns>
        ''' <remarks></remarks>
        Public Shared Function PrintRecordSet(ByRef a As Area, ByVal Command As String, ByRef cf As ColumnFormatter) As Integer
            Dim maxNo As Integer = cf.Count - 1
            Dim baseRow As Integer = a.Row, baseColumn As Integer = a.Column
            Dim rowCnt As Integer = LoadCommand(Command)
            For i As Integer = 0 To maxNo
                setFormat(Cell(Base(baseRow, baseColumn + cf.ColumnNo(i) - 1), rowCnt - 1, 0), cf.ColumnFormat(i))
            Next
            PrintRecordSet(Cell(baseRow, baseColumn))
            maxNo = Nothing
            baseRow = Nothing
            baseColumn = Nothing
            Return rowCnt
        End Function

        Public Shared Sub setFontSize(ByRef a As Area, ByVal size As Integer)
            a.setFontSize(size)
        End Sub

        Public Shared Sub FitRows(ByRef a As Area)
            a.FitRows()
        End Sub

        ''' <summary>
        '''��Excel�������ָ������ĵ�Ԫ���Ȱ���ʵ�����ݵĿ���Զ����е�����
        ''' </summary>
        ''' <param name="a">ָ����Excel����</param>
        ''' <remarks></remarks>
        Public Shared Sub FitColumns(ByRef a As Area)
            a.FitColumns()
        End Sub

        ''' <summary>
        ''' ���Excel������ָ�������еĹ�ʽ��
        ''' </summary>
        ''' <param name="a">ָ����Excel����</param>
        ''' <remarks>�������ʽ������ԭ�еĸ�ʽ���á�</remarks>
        Private Shared Sub ClearContents(ByRef a As Area)
            a.ClearContents()
        End Sub

        Enum AlignType As Integer
            CENTER = XlHAlign.xlHAlignCenter
            RIGHT = XlHAlign.xlHAlignRight
        End Enum

        ''' <summary>
        '''��Excel�������е�ָ��������ж��롣
        ''' </summary>
        ''' <param name="a">ָ����Excel����</param>
        ''' <param name="type">��Ԫ�����ķ�ʽ��CENTER���м���룻RIGHT���ұ߶��롣</param>
        ''' <remarks></remarks>
        Public Shared Sub Align(ByRef a As Area, Optional ByVal type As AlignType = AlignType.CENTER)
            a.Align(type)
        End Sub

        Public Shared Sub DeleteColumn(ByRef a As Area)
            a.DeleteColumn()
        End Sub

        Friend Shared Function IsEnoughData(ByRef SrcData() As Area, _
        ByVal ExpectedNumber As Integer, Optional ByVal Fraction As Double = 0.75) As Boolean
            Dim count As Integer = 0
            For i As Integer = 0 To SrcData.Length - 1
                count += SrcData(i).ValidCellCount
            Next
            Return count > Fraction * ExpectedNumber
        End Function

        ''' <summary>
        ''' Excel����������ķ���
        ''' </summary>
        ''' <remarks></remarks>
        Enum DirectionType
            LEFT
            RIGHT
            UP
            DOWN
        End Enum

        ''' <summary>
        ''' ���Excel��������ָ����Ԫ����������ı߽�λ�á�
        ''' </summary>
        ''' <param name="a">Excel�������е�ָ����Ԫ��</param>
        ''' <param name="direction">��ָ����Ԫ�������̽Ѱ����߽�ķ���</param>
        ''' <returns>��̽Ѱ����Ϊ���£��򷵻�����߽����ڵ�Ԫ���������
        ''' ��̽Ѱ����Ϊ���ң��򷵻�����߽����ڵ�Ԫ���������</returns>
        ''' <remarks></remarks>
        Public Shared Function GetBorderCellIndex(ByRef a As Area, direction As DirectionType) As Integer
            Return a.GetBorderCellIndex(direction)
        End Function

        Enum FormatType
            NUMBER
            NUMBER_NO_COMMA
            CURRENCY
            PERCENT
            PERCENT2
            DEC1
            DEC2
            DEC3
            DEC4
            TEXT
            SHORTDATE
        End Enum

        ''' <summary>
        ''' ����Excel������ָ����Ԫ�����ʾ��ʽ��
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="format">��Ԫ���ʽ���ͣ�NUMBER-ǧ��λ�ָ���������NUMBER_NO_COMMA-��ǧ��λ�ָ�������������
        ''' CURRENCY-ǧ��λ�ָ�������2λС������ֵ��PERCENT-����С�����1λ�İٷ�����
        ''' PERCENT2-����С�����2λ�İٷ�����DEC1��DEC4-����С������Ӧλ����С����TEXT-�ı���
        ''' SHORTDATE-�����ڸ�ʽ��
        ''' </param>
        ''' <remarks></remarks>
        Public Shared Sub setFormat(ByRef a As Area, ByVal format As FormatType)
            Select Case format
                Case FormatType.NUMBER
                    a.setFormat("#,##0")
                Case FormatType.NUMBER_NO_COMMA
                    a.setFormat("0")
                Case FormatType.CURRENCY
                    a.setFormat("#,##0.00")
                Case FormatType.PERCENT
                    a.setFormat("0.0%")
                Case FormatType.PERCENT2
                    a.setFormat("0.00%")
                Case FormatType.DEC1
                    a.setFormat("0.0")
                Case FormatType.DEC2
                    a.setFormat("0.00")
                Case FormatType.DEC3
                    a.setFormat("0.000")
                Case FormatType.DEC4
                    a.setFormat("0.0000")
                Case FormatType.TEXT
                    a.setFormat("@")
                Case FormatType.SHORTDATE
                    a.setFormat("yyyy-mm-dd")
            End Select
        End Sub

        Public Shared Sub setFontItalic(ByRef a As Area, Optional ByVal isItalic As Boolean = True)
            a.setFontItalic(isItalic)
        End Sub

        ''' <summary>
        ''' ����Excel������ָ����Ԫ��ļ��㹫ʽ��
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="formula">Excel��ʽ���ַ�������������ͷ�ĵȺš�=��</param>
        ''' <remarks>��ʽ�ַ����г��ֵĶ��������������ʹ��A1��ʽ</remarks>
        Public Shared Sub setFormula(ByRef a As Area, ByVal formula As String)
            a.setFormula(formula)
        End Sub

        ''' <summary>
        '''��R1C1���õķ�ʽ����Excel������ָ����Ԫ��ļ��㹫ʽ
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="formula">Excel��ʽ���ַ�������������ͷ�ĵȺš�=��</param>
        ''' <remarks>��ʽ�ַ����г��ֵĶ��������������ʹ��R1C1��ʽ</remarks>
        Public Shared Sub setFormulaR1C1(ByRef a As Area, ByVal formula As String)
            a.setFormulaR1C1(formula)
        End Sub

        Public Shared Sub DivideByRow(ByRef a As Area, ByVal numeratorColumn As Integer, ByVal denominatorColumn As Integer)
            setFormula(a, Divide(Address(a.Row, numeratorColumn), Address(a.Row, denominatorColumn)))
        End Sub

        Public Shared Function PrintDataTableToHTML(ByRef dt As Data.DataTable) As String
            Dim m_HtmlStr As String = Nothing
            Dim rowCount As Integer = dt.Rows.Count
            If rowCount > 0 Then
                Dim columnCount As Integer = dt.Columns.Count
                For i As Integer = 0 To rowCount - 1
                    For j As Integer = 0 To columnCount - 1
                        m_HtmlStr = m_HtmlStr + "<td align=""center"">" + CStr(dt.Rows(i).ItemArray(j)) + "</td>"
                    Next
                    m_HtmlStr = m_HtmlStr + "</tr>"
                Next
                columnCount = Nothing
            End If
            rowCount = Nothing
            Return m_HtmlStr
        End Function

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�����������һ�����ݾ����е�����Ԫ�ء������Excel�ĵ�Ԫ��
        ''' ����������Ԫ����һһ��Ӧ�ģ����������Ԫ�������������������ֱ���ھ����������
        ''' �������ҵ�Ԫ���������Ͻǡ����½ǻ�����һ��λ�õĵ�Ԫ������������Ͻǡ����½ǻ�
        ''' ��Ӧλ�õ�Ԫ��һһ��Ӧ��
        ''' </summary>
        ''' <param name="a">ָ��Excel��Ԫ����������Ͻǵ�Ԫ��</param>
        ''' <param name="mt">��������ݵ����ݾ���</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintMatrix(ByRef a As Area, ByRef mt As Matrix(Of Object))
            PrintMatrix(a, mt, 1, 1, mt.RowsCount, mt.ColumnsCount)
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�����������SQL��ѯ���Ľ�������ݡ������Excel�ĵ�Ԫ��
        ''' ������������������һһ��Ӧ�ģ����������Ԫ�������������������ֱ���ڽ������������
        ''' ������
        ''' </summary>
        ''' <param name="a">ָ��Excel��Ԫ����������Ͻǵ�Ԫ��</param>
        ''' <param name="SQL">��ִ�е�SQL��ѯ��䡣</param>
        ''' <returns>��ѯ�����������������</returns>
        Private Shared Function PrintSQLQueryByMatrix(ByRef a As Area, SQL As String) As Integer
            Dim mt As Matrix(Of Object) = New Matrix(Of Object)(SQL)
            Dim row As Integer = mt.RowsCount
            PrintMatrix(a, mt)
            mt.Clear()
            mt = Nothing
            Return row
        End Function

        Private Shared m_PrintInGroupName As String = Nothing
        Private Shared m_GroupNameTable As Dictionary(Of String, List(Of Data.DataTable)) = Nothing
        Public Shared Property PrintInGroupName As String
            Get
                Return m_PrintInGroupName
            End Get
            Set(value As String)
                ClearGroupNameTable()
                m_PrintInGroupName = value
                m_GroupNameTable = New Dictionary(Of String, List(Of Data.DataTable))
            End Set
        End Property

        Private Shared Sub ClearGroupNameTable()
            If m_GroupNameTable IsNot Nothing Then
                If m_GroupNameTable.Count > 0 Then
                    Dim cnt As Integer = Nothing
                    For Each lst As List(Of Data.DataTable) In m_GroupNameTable.Values
                        If lst IsNot Nothing Then
                            cnt = lst.Count
                            If cnt > 0 Then
                                For i As Integer = 0 To cnt - 1
                                    lst.Item(i).Clear()
                                    lst.Item(i) = Nothing
                                Next
                                lst.Clear()
                            End If
                            lst = Nothing
                        End If
                    Next
                    m_GroupNameTable.Clear()
                    cnt = Nothing
                End If
                m_GroupNameTable = Nothing
            End If
        End Sub

        Public Shared Function GroupValues() As List(Of String)
            If m_GroupNameTable Is Nothing Then
                Return Nothing
            End If
            If m_GroupNameTable.Count = 0 Then
                Return Nothing
            End If
            Dim lst As List(Of String) = New List(Of String)
            For Each s As String In m_GroupNameTable.Keys
                lst.Add(s)
            Next
            Return lst
        End Function

        Private Shared Sub FillGroupNameTable(ByRef dt As Data.DataTable)
            Dim index As Integer = IndexOfColumnName(dt, m_PrintInGroupName)
            If index <> -1 Then
                Dim subDts As Dictionary(Of String, Data.DataTable) = Split(dt, index)
                Dim lst As List(Of Data.DataTable) = Nothing
                If m_GroupNameTable.Count = 0 Then
                    For Each s As String In subDts.Keys
                        lst = New List(Of Data.DataTable)
                        lst.Add(subDts(s).Copy)
                        m_GroupNameTable.Add(s, lst)
                    Next
                    lst = Nothing
                Else
                    Dim length As Integer = -1
                    For Each s As String In m_GroupNameTable.Keys
                        If length = -1 Then
                            length = m_GroupNameTable(s).Count
                        End If
                        If subDts.ContainsKey(s) Then
                            m_GroupNameTable(s).Add(subDts(s).Copy)
                        Else
                            m_GroupNameTable(s).Add(New Data.DataTable())
                        End If
                    Next
                    For Each s As String In subDts.Keys
                        If Not m_GroupNameTable.ContainsKey(s) Then
                            lst = New List(Of Data.DataTable)
                            For i As Integer = 1 To length
                                lst.Add(New Data.DataTable)
                            Next
                            lst.Add(subDts(s).Copy)
                            m_GroupNameTable.Add(s, lst)
                        End If
                    Next
                    lst = Nothing
                    length = Nothing
                End If
                For Each dt1 As Data.DataTable In subDts.Values
                    dt1.Clear()
                    dt1 = Nothing
                Next
                subDts.Clear()
                subDts = Nothing
            End If
            index = Nothing
        End Sub

        Public Shared Function PrintGroupValue(ByRef a As Area, GroupValue As String, Index As Integer) As Integer
            If m_GroupNameTable Is Nothing Then
                Return 0
            End If
            If Not m_GroupNameTable.ContainsKey(GroupValue) Then
                Return 0
            End If
            Dim lst As List(Of Data.DataTable) = m_GroupNameTable(GroupValue)
            If lst Is Nothing Then
                Return 0
            End If
            If lst.Count < Index Or Index < 1 Then
                Return 0
            End If
            Dim dt As Data.DataTable = lst(Index - 1)
            If dt Is Nothing Then
                Return 0
            End If
            PrintDataTableToExcel(a, dt)
            Return dt.Rows.Count
        End Function
        Public Shared Function PrintGroupValue(GroupValue As String) As String
            Dim lst As List(Of Data.DataTable) = m_GroupNameTable(GroupValue)
            Dim dt As Data.DataTable = lst(0)
            Return PrintDataTableToHTML(dt)
        End Function

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�����������SQL��ѯ���Ľ�������ݡ������Excel�ĵ�Ԫ��
        ''' ������������������һһ��Ӧ�ģ����������Ԫ�������������������ֱ���ڽ������������
        ''' ������
        ''' </summary>
        ''' <param name="a">ָ��Excel��Ԫ����������Ͻǵ�Ԫ��</param>
        ''' <param name="SQL">��ִ�е�SQL��ѯ��䡣</param>
        ''' <returns>��ѯ�����������������</returns>
        Public Shared Function PrintSQLQuery(ByRef a As Area, SQL As String) As Integer
            Dim dt As Data.DataTable = GetDataTable(SQL)
            Dim rowCount As Integer = dt.Rows.Count
            If a IsNot Nothing Then
                PrintDataTableToExcel(a, dt)
            End If
            If m_PrintInGroupName IsNot Nothing Then
                FillGroupNameTable(dt)
            End If
            dt.Clear()
            dt = Nothing
            Return rowCount
        End Function

        Private Shared Sub PrintDataTableToExcel(ByRef a As Area, ByRef dt As Data.DataTable)
            Dim rowCount As Integer = dt.Rows.Count
            If rowCount > 0 Then
                Dim columnCount As Integer = dt.Columns.Count
                Dim rawData(rowCount - 1, columnCount - 1) As Object
                For i As Integer = 0 To rowCount - 1
                    For j As Integer = 0 To columnCount - 1
                        rawData(i, j) = dt.Rows(i).ItemArray(j)
                    Next
                Next
                If rawData IsNot Nothing Then
                    Cell(Base(a.Row, a.Column), rowCount - 1, columnCount - 1).setValueArray(rawData)
                    Array.Clear(rawData, 0, rawData.Length)
                    rawData = Nothing
                End If
                columnCount = Nothing
            End If
            rowCount = Nothing
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�������а�Ԥ�����ʽ���һ�����ݾ����е�����Ԫ�ء������Excel�ĵ�Ԫ��
        ''' ����������Ԫ����һһ��Ӧ�ģ����������Ԫ�������������������ֱ���ھ����������
        ''' �������ҵ�Ԫ���������Ͻǡ����½ǻ�����һ��λ�õĵ�Ԫ������������Ͻǡ����½ǻ�
        ''' ��Ӧλ�õ�Ԫ��һһ��Ӧ��
        ''' </summary>
        ''' <param name="a">ָ��Excel��Ԫ����������Ͻǵ�Ԫ��</param>
        ''' <param name="mt">��������ݵ����ݾ���</param>
        ''' <param name="cf">Ԥ����ľ����е�Excel��Ԫ���ʽ��</param>
        ''' <remarks></remarks>
        Private Shared Sub PrintMatrix(ByRef a As Area, ByRef mt As Matrix(Of Object), ByRef cf As ColumnFormatter)
            Dim maxNo As Integer = cf.Count - 1
            Dim baseRow As Integer = a.Row, baseColumn As Integer = a.Column
            Dim rowCnt As Integer = mt.RowsCount, columnCnt As Integer = mt.ColumnsCount
            For i As Integer = 0 To maxNo
                setFormat(Cell(Base(baseRow, baseColumn + cf.ColumnNo(i) - 1), rowCnt - 1, 0), cf.ColumnFormat(i))
            Next
            PrintMatrix(Cell(baseRow, baseColumn), mt, 1, 1, rowCnt, columnCnt)
            maxNo = Nothing
            baseRow = Nothing
            baseColumn = Nothing
            rowCnt = Nothing
            columnCnt = Nothing
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�����������һ�����ݾ�����ָ�����Ӿ��������Ԫ�ء�
        ''' �����Excel�ĵ�Ԫ���������Ӿ����Ԫ����һһ��Ӧ�ģ����������Ԫ�������
        ''' �������������ֱ�����Ӿ�����������������ҵ�Ԫ���������Ͻǡ����½ǻ�����
        ''' һ��λ�õĵ�Ԫ������Ӿ�������Ͻǡ����½ǻ��Ӧλ�õ�Ԫ��һһ��Ӧ��
        ''' </summary>
        ''' <param name="a">ָ��Excel��Ԫ����������Ͻǵ�Ԫ��</param>
        ''' <param name="mt">��������ݵ����ݾ���</param>
        ''' <param name="topRow">�Ӿ�������Ͻǵ�Ԫ���ھ����е��е�λ�ã�������һ�е�λ��Ϊ1��</param>
        ''' <param name="leftColumn">�Ӿ�������Ͻǵ�Ԫ���ھ����е��е�λ�ã������һ�е�λ��Ϊ1��</param>
        ''' <param name="rowCounts">�Ӿ��������</param>
        ''' <param name="columnCounts">�Ӿ��������</param>
        ''' <remarks>���Ӿ�������ϽǺ�����������ķ�Χ���ܳ���ԭʼ���ݾ���ķ�Χ��������������ݡ�</remarks>
        Private Shared Sub PrintMatrix(ByRef a As Area, ByRef mt As Matrix(Of Object),
                                      ByVal topRow As Integer, ByVal leftColumn As Integer,
                                      ByVal rowCounts As Integer, ByVal columnCounts As Integer)
            If topRow + rowCounts - 1 > mt.RowsCount _
                Or leftColumn + columnCounts - 1 > mt.ColumnsCount Then
                Return
            End If

            Dim row As Integer = a.Row, column As Integer = a.Column
            For i As Integer = 1 To rowCounts
                For j As Integer = 1 To columnCounts
                    Cell(row + i - 1, column + j - 1).setValue(mt.Cell(i + topRow - 1, j + leftColumn - 1))
                Next
            Next
            row = Nothing
            column = Nothing
        End Sub

        Public Shared Sub PrintTriangle(ByRef a As Area, ByRef t As Triangle(Of Object))
            Dim row As Integer = a.Row, column As Integer = a.Column
            Dim n As Integer = t.EdgeLength
            For i As Integer = 1 To n
                For j As Integer = 1 To n + 1 - i
                    Cell(row + i - 1, column + j - 1).setValue(t.Cell(i, j))
                Next
            Next
            row = Nothing
            column = Nothing
        End Sub

        Public Shared Sub PrintDerivativeTriangle(ByRef a As Area, ByVal row As Integer, ByVal column As Integer, ByVal length As Integer)
            Dim srcRow As Integer = a.Row, srcColumn As Integer = a.Column
            For i As Integer = 1 To length - 1
                For j As Integer = 1 To length - i
                    Cell(row + i - 1, column + j - 1).setFormula( _
                    Divide(Address(srcRow + i - 1, srcColumn + j), Address(srcRow + i - 1, srcColumn + j - 1)))
                Next
            Next
            srcRow = Nothing
            srcColumn = Nothing
        End Sub

        Enum BorderType As Integer
            ALL
            OUTSIDE
            VERTICAL
            HORIZONTAL
            BOTTOM
        End Enum

        Enum BorderWeight As Integer
            THIN = XlBorderWeight.xlThin
            MEDIUM = XlBorderWeight.xlMedium
        End Enum

        Public Shared Sub setBorders(ByRef a As Area, Optional ByVal type As BorderType = BorderType.ALL, _
        Optional ByVal weight As BorderWeight = BorderWeight.THIN)
            a.setBorders(type, weight)
        End Sub

        Enum Color As Integer
            BLUE = 5
            RED = 3
            YELLOW = 6
            SKYBLUE = 23
        End Enum

        ''' <summary>
        ''' ����Excel������ָ����Ԫ����ڲ������ɫ��
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="clr">ָ�����ڲ������ɫ��Ŀǰ֧�֣�BLUE��ɫ��RED��ɫ��YELLOW��ɫ��SKYBLUE����ɫ��</param>
        ''' <remarks></remarks>
        Public Shared Sub setInteriorColor(ByRef a As Area, ByVal clr As Color)
            a.setInteriorColor(clr)
        End Sub

        ''' <summary>
        ''' ����Excel������ָ����Ԫ���������ʾ��ɫ��
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="clr">ָ����������ʾ��ɫ��Ŀǰ֧�֣�BLUE��ɫ��RED��ɫ��YELLOW��ɫ��SKYBLUE����ɫ��</param>
        ''' <remarks></remarks>
        Public Shared Sub setColor(ByRef a As Area, ByVal clr As Color)
            a.setColor(clr)
        End Sub

        ''' <summary>
        ''' ��Excel����������ɸ�ָ����Ԫ���а�˳�����һ���ַ��������е�ÿ���ַ�����
        ''' ÿ���ַ������Ե���ռһ��Excel��Ԫ��Ҳ���Ժϲ���ռ�ݶ����Ԫ��
        ''' </summary>
        ''' <param name="a">ָ����Ԫ����λ�������Ͻǵĵ�Ԫ��</param>
        ''' <param name="s">��������ַ������顣</param>
        ''' <param name="orient">�ַ�����Excel�������е��������
        ''' HORIZONTALΪ����ˮƽ�����VERTICALΪ����ֱ�����</param>
        ''' <param name="columnCount">ÿһ���ַ�����ռ��Excel�������е�������</param>
        ''' <param name="rowCount">ÿһ���ַ�����ռ��Excel�������е�������</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintText(ByRef a As Area, ByRef s As String(), _
        Optional ByVal orient As Orientation = Orientation.HORIZONTAL, _
        Optional ByVal columnCount As Integer = 1, Optional ByVal rowCount As Integer = 1)
            a.setValue(s, orient, columnCount, rowCount)
        End Sub

        Public Shared Sub PrintText(ByRef a As Area, ByVal s As String, _
        ByVal Frequency As Integer, ByVal rowInterval As Integer, ByVal columnInterval As Integer)
            a.setValue(s, Frequency, rowInterval, columnInterval)
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�������һ���ַ��������ַ������Զ�ռһ��Excel��Ԫ��
        ''' Ҳ���Ժϲ���ռ�ݶ����Ԫ��
        ''' </summary>
        ''' <param name="a">ָ����Ԫ��</param>
        ''' <param name="s">��������ַ�����</param>
        ''' <param name="rowDelta">����ַ�����ռExcel������Ԫ�������������һ����ָ����Ԫ�������еĲ</param>
        ''' <param name="columnDelta">����ַ�����ռExcel������Ԫ�������������һ����ָ����Ԫ�������еĲ</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintText(ByRef a As Area, ByVal s As String, _
        Optional ByVal rowDelta As Integer = 0, Optional ByVal columnDelta As Integer = 0)
            a.setValue(s, rowDelta, columnDelta)
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ���е��ض��ַ����滻Ϊ��һ��ָ�����ַ�����
        ''' </summary>
        ''' <param name="a">ָ����Ԫ��</param>
        ''' <param name="oldValue">Ҫ���滻���ַ�����</param>
        ''' <param name="newValue">Ҫ�滻���ֵ������ַ�����</param>
        Public Shared Sub ReplaceText(ByRef a As Area, oldValue As String, newValue As String)
            Dim s As String = a.ValueAsString
            a.setValue(s.Replace(oldValue, newValue), 0, 0)
            s = Nothing
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�������һ�����ڡ�
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="dte">�����������ֵ��</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintDate(ByRef a As Area, ByVal dte As Date)
            a.setValue(dte)
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�������һ��������
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="val">�����������ֵ��</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintInteger(ByRef a As Area, ByVal val As Integer)
            a.setValue(val)
        End Sub

        Public Shared Sub PrintDouble(ByRef a As Area, ByVal val As Double)
            a.setValue(val)
        End Sub

        ''' <summary>
        ''' ��Excel�������ָ����Ԫ�������һ����ȷС����
        ''' </summary>
        ''' <param name="a">ָ���ĵ�Ԫ��</param>
        ''' <param name="val">������ľ�ȷС��ֵ��</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintDecimal(ByRef a As Area, ByVal val As Decimal)
            a.setValue(val)
        End Sub

    End Class
End Namespace
