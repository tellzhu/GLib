Imports Microsoft.Office.Interop.Excel
Imports dotNet.db.Admin
Imports dotNet.db.DBExpert
Imports dotNet.math
Imports dotNet.office.Excelist
Imports dotNet.office.FormulaBuilder

Namespace office
	Public Class AreaExpert

        ''' <summary>
        ''' 设置Excel工作表指定单元格的注释内容。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="comment">注释的文本内容。</param>
        ''' <remarks></remarks>
		Public Shared Sub setComment(ByRef a As Area, ByVal comment As String)
			a.setComment(comment)
		End Sub

		Public Shared Sub setFontUnderline(ByRef a As Area, Optional ByVal isUnderline As Boolean = True)
			a.setFontUnderline(isUnderline)
		End Sub

        ''' <summary>
        ''' 设置Excel工作表指定单元格的字体加粗显示方式。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="isBold">是否将字体加粗。</param>
        ''' <remarks></remarks>
		Public Shared Sub setFontBold(ByRef a As Area, Optional ByVal isBold As Boolean = True)
			a.setFontBold(isBold)
        End Sub

        Private Shared Sub PrintRecordSet(ByRef a As Area)
            a.PrintRecordSet()
        End Sub

        ''' <summary>
        ''' 在Excel工作表的指定单元格区域中输出一个SQL查询语句结果集中的所有元素。输出到Excel的单元格
        ''' 区域与结果集矩阵的元素是一一对应的，即：输出单元格的区域的行数、列数分别等于矩阵的行数、
        ''' 列数，且单元格区域左上角、右下角或任意一个位置的单元格均与矩阵的左上角、右下角或
        ''' 对应位置的元素一一对应。
        ''' </summary>
        ''' <param name="a">指定Excel单元格区域的左上角单元格。</param>
        ''' <param name="Command">待输出内容的SQL查询语句。</param>
        ''' <returns>结果集的行数。</returns>
        ''' <remarks></remarks>
        Public Shared Function PrintRecordSet(ByRef a As Area, ByVal Command As String) As Integer
            Dim rowCnt As Integer = LoadCommand(Command)
            a.PrintRecordSet()
            Return rowCnt
        End Function

        ''' <summary>
        ''' 在Excel工作表的指定单元格区域中按预定义格式输出一个SQL查询语句结果集中的所有元素。输出到Excel的单元格
        ''' 区域与结果集矩阵的元素是一一对应的，即：输出单元格的区域的行数、列数分别等于矩阵的行数、
        ''' 列数，且单元格区域左上角、右下角或任意一个位置的单元格均与矩阵的左上角、右下角或
        ''' 对应位置的元素一一对应。
        ''' </summary>
        ''' <param name="a">指定Excel单元格区域的左上角单元格。</param>
        ''' <param name="Command">待输出内容的SQL查询语句。</param>
        ''' <param name="cf">预定义的矩阵列到Excel单元格格式。</param>
        ''' <returns>结果集的行数。</returns>
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
        '''将Excel工作表的指定区域的单元格宽度按照实际数据的宽度自动进行调整。
        ''' </summary>
        ''' <param name="a">指定的Excel区域。</param>
        ''' <remarks></remarks>
        Public Shared Sub FitColumns(ByRef a As Area)
            a.FitColumns()
        End Sub

        ''' <summary>
        ''' 清除Excel工作表指定区域中的公式。
        ''' </summary>
        ''' <param name="a">指定的Excel区域。</param>
        ''' <remarks>仅清除公式，保留原有的格式设置。</remarks>
        Private Shared Sub ClearContents(ByRef a As Area)
            a.ClearContents()
        End Sub

        Enum AlignType As Integer
            CENTER = XlHAlign.xlHAlignCenter
            RIGHT = XlHAlign.xlHAlignRight
        End Enum

        ''' <summary>
        '''将Excel工作表中的指定区域进行对齐。
        ''' </summary>
        ''' <param name="a">指定的Excel区域。</param>
        ''' <param name="type">单元格对齐的方式：CENTER是中间对齐；RIGHT是右边对齐。</param>
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
        ''' Excel工作表区域的方向。
        ''' </summary>
        ''' <remarks></remarks>
        Enum DirectionType
            LEFT
            RIGHT
            UP
            DOWN
        End Enum

        ''' <summary>
        ''' 获得Excel工作表中指定单元格所在区域的边界位置。
        ''' </summary>
        ''' <param name="a">Excel工作表中的指定单元格。</param>
        ''' <param name="direction">从指定单元格出发，探寻区域边界的方向。</param>
        ''' <returns>若探寻方向为上下，则返回区域边界所在单元格的行数；
        ''' 若探寻方向为左右，则返回区域边界所在单元格的列数。</returns>
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
        ''' 设置Excel工作表指定单元格的显示格式。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="format">单元格格式类型：NUMBER-千分位分隔的整数；NUMBER_NO_COMMA-无千分位分隔的连续整数；
        ''' CURRENCY-千分位分隔、保留2位小数的数值；PERCENT-保留小数点后1位的百分数；
        ''' PERCENT2-保留小数点后2位的百分数；DEC1至DEC4-保留小数点后对应位数的小数；TEXT-文本；
        ''' SHORTDATE-短日期格式。
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
        ''' 设置Excel工作表指定单元格的计算公式。
        ''' </summary>
        ''' <param name="a">指定的单元格</param>
        ''' <param name="formula">Excel公式的字符串，不包括开头的等号“=”</param>
        ''' <remarks>公式字符串中出现的对其他区域的引用使用A1样式</remarks>
        Public Shared Sub setFormula(ByRef a As Area, ByVal formula As String)
            a.setFormula(formula)
        End Sub

        ''' <summary>
        '''以R1C1引用的方式设置Excel工作表指定单元格的计算公式
        ''' </summary>
        ''' <param name="a">指定的单元格</param>
        ''' <param name="formula">Excel公式的字符串，不包括开头的等号“=”</param>
        ''' <remarks>公式字符串中出现的对其他区域的引用使用R1C1样式</remarks>
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
        ''' 在Excel工作表的指定单元格区域中输出一个数据矩阵中的所有元素。输出到Excel的单元格
        ''' 区域与矩阵的元素是一一对应的，即：输出单元格的区域的行数、列数分别等于矩阵的行数、
        ''' 列数，且单元格区域左上角、右下角或任意一个位置的单元格均与矩阵的左上角、右下角或
        ''' 对应位置的元素一一对应。
        ''' </summary>
        ''' <param name="a">指定Excel单元格区域的左上角单元格。</param>
        ''' <param name="mt">待输出内容的数据矩阵。</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintMatrix(ByRef a As Area, ByRef mt As Matrix(Of Object))
            PrintMatrix(a, mt, 1, 1, mt.RowsCount, mt.ColumnsCount)
        End Sub

        ''' <summary>
        ''' 在Excel工作表的指定单元格区域中输出SQL查询语句的结果集数据。输出到Excel的单元格
        ''' 区域与结果集的数据是一一对应的，即：输出单元格的区域的行数、列数分别等于结果集的行数、
        ''' 列数。
        ''' </summary>
        ''' <param name="a">指定Excel单元格区域的左上角单元格。</param>
        ''' <param name="SQL">待执行的SQL查询语句。</param>
        ''' <returns>查询结果集的数据行数。</returns>
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
        ''' 在Excel工作表的指定单元格区域中输出SQL查询语句的结果集数据。输出到Excel的单元格
        ''' 区域与结果集的数据是一一对应的，即：输出单元格的区域的行数、列数分别等于结果集的行数、
        ''' 列数。
        ''' </summary>
        ''' <param name="a">指定Excel单元格区域的左上角单元格。</param>
        ''' <param name="SQL">待执行的SQL查询语句。</param>
        ''' <returns>查询结果集的数据行数。</returns>
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
        ''' 在Excel工作表的指定单元格区域中按预定义格式输出一个数据矩阵中的所有元素。输出到Excel的单元格
        ''' 区域与矩阵的元素是一一对应的，即：输出单元格的区域的行数、列数分别等于矩阵的行数、
        ''' 列数，且单元格区域左上角、右下角或任意一个位置的单元格均与矩阵的左上角、右下角或
        ''' 对应位置的元素一一对应。
        ''' </summary>
        ''' <param name="a">指定Excel单元格区域的左上角单元格。</param>
        ''' <param name="mt">待输出内容的数据矩阵。</param>
        ''' <param name="cf">预定义的矩阵列到Excel单元格格式。</param>
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
        ''' 在Excel工作表的指定单元格区域中输出一个数据矩阵中指定的子矩阵的所有元素。
        ''' 输出到Excel的单元格区域与子矩阵的元素是一一对应的，即：输出单元格的区域
        ''' 的行数、列数分别等于子矩阵的行数、列数，且单元格区域左上角、右下角或任意
        ''' 一个位置的单元格均与子矩阵的左上角、右下角或对应位置的元素一一对应。
        ''' </summary>
        ''' <param name="a">指定Excel单元格区域的左上角单元格</param>
        ''' <param name="mt">待输出内容的数据矩阵</param>
        ''' <param name="topRow">子矩阵的左上角单元格在矩阵中的行的位置，最上面一行的位置为1。</param>
        ''' <param name="leftColumn">子矩阵的左上角单元格在矩阵中的列的位置，最左边一列的位置为1。</param>
        ''' <param name="rowCounts">子矩阵的行数</param>
        ''' <param name="columnCounts">子矩阵的列数</param>
        ''' <remarks>由子矩阵的左上角和行列数定义的范围不能超出原始数据矩阵的范围，否则不能输出内容。</remarks>
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
        ''' 设置Excel工作表指定单元格的内部填充颜色。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="clr">指定的内部填充颜色。目前支持：BLUE蓝色，RED红色、YELLOW黄色和SKYBLUE天蓝色。</param>
        ''' <remarks></remarks>
        Public Shared Sub setInteriorColor(ByRef a As Area, ByVal clr As Color)
            a.setInteriorColor(clr)
        End Sub

        ''' <summary>
        ''' 设置Excel工作表指定单元格的字体显示颜色。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="clr">指定的字体显示颜色。目前支持：BLUE蓝色，RED红色、YELLOW黄色和SKYBLUE天蓝色。</param>
        ''' <remarks></remarks>
        Public Shared Sub setColor(ByRef a As Area, ByVal clr As Color)
            a.setColor(clr)
        End Sub

        ''' <summary>
        ''' 在Excel工作表的若干个指定单元格中按顺序输出一个字符串数组中的每个字符串，
        ''' 每个字符串可以单独占一个Excel单元格，也可以合并后占据多个单元格。
        ''' </summary>
        ''' <param name="a">指定单元格中位于最左上角的单元格。</param>
        ''' <param name="s">待输出的字符串数组。</param>
        ''' <param name="orient">字符串在Excel工作表中的输出方向：
        ''' HORIZONTAL为横向水平输出；VERTICAL为竖向垂直输出。</param>
        ''' <param name="columnCount">每一个字符串所占的Excel工作表列的数量。</param>
        ''' <param name="rowCount">每一个字符串所占的Excel工作表行的数量。</param>
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
        ''' 在Excel工作表的指定单元格中输出一个字符串，该字符串可以独占一个Excel单元格，
        ''' 也可以合并后占据多个单元格。
        ''' </summary>
        ''' <param name="a">指定单元格。</param>
        ''' <param name="s">待输出的字符串。</param>
        ''' <param name="rowDelta">输出字符串所占Excel工作表单元格区域的最下面一行与指定单元格所在行的差。</param>
        ''' <param name="columnDelta">输出字符串所占Excel工作表单元格区域的最右面一列与指定单元格所在列的差。</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintText(ByRef a As Area, ByVal s As String, _
        Optional ByVal rowDelta As Integer = 0, Optional ByVal columnDelta As Integer = 0)
            a.setValue(s, rowDelta, columnDelta)
        End Sub

        ''' <summary>
        ''' 将Excel工作表的指定单元格中的特定字符串替换为另一个指定的字符串。
        ''' </summary>
        ''' <param name="a">指定单元格。</param>
        ''' <param name="oldValue">要被替换的字符串。</param>
        ''' <param name="newValue">要替换出现的所有字符串。</param>
        Public Shared Sub ReplaceText(ByRef a As Area, oldValue As String, newValue As String)
            Dim s As String = a.ValueAsString
            a.setValue(s.Replace(oldValue, newValue), 0, 0)
            s = Nothing
        End Sub

        ''' <summary>
        ''' 在Excel工作表的指定单元格中输出一个日期。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="dte">待输出的日期值。</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintDate(ByRef a As Area, ByVal dte As Date)
            a.setValue(dte)
        End Sub

        ''' <summary>
        ''' 在Excel工作表的指定单元格中输出一个整数。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="val">待输出的整数值。</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintInteger(ByRef a As Area, ByVal val As Integer)
            a.setValue(val)
        End Sub

        Public Shared Sub PrintDouble(ByRef a As Area, ByVal val As Double)
            a.setValue(val)
        End Sub

        ''' <summary>
        ''' 在Excel工作表的指定单元格中输出一个精确小数。
        ''' </summary>
        ''' <param name="a">指定的单元格。</param>
        ''' <param name="val">待输出的精确小数值。</param>
        ''' <remarks></remarks>
        Public Shared Sub PrintDecimal(ByRef a As Area, ByVal val As Decimal)
            a.setValue(val)
        End Sub

    End Class
End Namespace
