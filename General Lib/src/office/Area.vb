Imports Microsoft.Office.Interop.Excel
Imports dotNet.db.Admin
Imports dotNet.office.Excelist
Imports dotNet.office.AreaExpert

Namespace office
    ''' <summary>
    ''' Excel工作表中的一个区域范围，绝大多数情况下是单个单元格（如“A7”）或一个矩形区域（如“A7：B15”）
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Area

        Private r As Range
        Private sheet As Worksheet

        Friend Function PrintRecordSet() As Integer
            Return CopyFromRecordset(r)
        End Function

        Friend Sub setWorksheet(ByRef sheet As Worksheet)
            Me.sheet = sheet
        End Sub

        Friend Sub FitRows()
            r.Rows.AutoFit()
        End Sub

        Friend Sub FitColumns()
            r.Columns.AutoFit()
        End Sub

        Friend Sub setFontSize(ByVal size As Integer)
            r.Font.Size = size
        End Sub

        Public Sub Union(ByRef a As Area)
            Me.r = Me.r.Application.Union(Me.r, a.r)
        End Sub

        ''' <summary>
        ''' 获得Excel单元格对应的小数值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ValueAsDecimal() As Decimal
            Get
                Return CDec(r.Value)
            End Get
        End Property

        ''' <summary>
        ''' 获得Excel单元格对应的整数值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ValueAsInteger() As Integer
            Get
                Return CInt(r.Value)
            End Get
        End Property

        ''' <summary>
        ''' 获得Excel单元格对应的日期值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ValueAsDate As Date
            Get
                Return CDate(r.Value).Date
            End Get
        End Property

        ''' <summary>
        ''' 获得Excel单元格对应的字符串类型值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ValueAsString() As String
            Get
                Return CStr(r.Value)
            End Get
        End Property

        ''' <summary>
        ''' 获得Excel单元格对应的布尔类型值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ValueAsBoolean As Boolean
            Get
                Return CBool(r.Value)
            End Get
        End Property

        Friend ReadOnly Property Range() As Range
            Get
                Return r
            End Get
        End Property

        Friend Function GetBorderCellIndex(direction As DirectionType) As Integer
            Select Case direction
                Case DirectionType.LEFT
                    Return r.End(XlDirection.xlToLeft).Column
                Case DirectionType.RIGHT
                    Return r.End(XlDirection.xlToRight).Column
                Case DirectionType.UP
                    Return r.End(XlDirection.xlUp).Row
                Case DirectionType.DOWN
                    Return r.End(XlDirection.xlDown).Row
            End Select
            Return Nothing
        End Function

        Friend Sub setComment(ByVal comment As String)
            With r
                .AddComment()
                .Comment.Visible = False
                .Comment.Text(comment)
            End With
        End Sub

        Friend Sub setFontUnderline(Optional ByVal isUnderline As Boolean = True)
            r.Font.Underline = isUnderline
        End Sub

        Friend Sub Align(ByVal type As AlignType)
            r.HorizontalAlignment = type
        End Sub

        Friend Sub setFormat(ByVal format As String)
            r.NumberFormat = format
        End Sub

        Friend Sub setFormula(ByVal formula As String)
            r.Formula = "=" + formula
        End Sub

        Friend Sub setFormulaR1C1(ByVal formula As String)
            r.FormulaR1C1 = "=" + formula
        End Sub

        Friend Sub setPoints(ByVal topRow As Integer, ByVal topColumn As Integer, _
        ByVal bottomRow As Integer, ByVal bottomColumn As Integer)
            r = Cell(topRow, topColumn, bottomRow, bottomColumn)
        End Sub

        Friend Sub setRange(ByVal address As String)
            r = sheet.Range(address)
        End Sub

        ''' <summary>
        ''' 获得Excel单元格区域左上角的列号。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Row() As Integer
            Get
                Return r.Row
            End Get
        End Property

        Friend ReadOnly Property ValidCellCount() As Integer
            Get
                Dim count As Integer = 0
                For n As Integer = 1 To r.Areas.Count
                    For i As Integer = 1 To r.Areas(n).Rows.Count
                        For j As Integer = 1 To r.Areas(n).Columns.Count
                            If Cell(r.Areas(n).Row + i - 1, r.Areas(n).Column + j - 1).Value IsNot Nothing Then
                                count += 1
                            End If
                        Next
                    Next
                Next
                Return count
            End Get
        End Property

        Friend Sub setFontBold(ByVal isBold As Boolean)
            r.Font.Bold = isBold
        End Sub

        Friend Sub setFontItalic(ByVal isItalic As Boolean)
            r.Font.Italic = isItalic
        End Sub

        Friend Sub setColor(ByVal clr As Color)
            r.Font.ColorIndex = clr
        End Sub

        Public Sub Alert(ByVal clr As Color)
            setColor(clr)
            setFontBold(True)
        End Sub

        Friend Sub setInteriorColor(ByVal clr As Color)
            r.Interior.ColorIndex = clr
        End Sub

        ''' <summary>
        ''' 判断Excel单元格是否包含有效值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsValid() As Boolean
            Get
                Return r.Value IsNot Nothing
            End Get
        End Property

        ''' <summary>
        ''' 获得Excel单元格区域左上角的行号。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Column() As Integer
            Get
                Return r.Column
            End Get
        End Property

        Friend Sub SelectColumn()
            r.EntireColumn.Select()
        End Sub

        Friend Sub DeleteColumn()
            r.EntireColumn.Delete()
        End Sub

        Friend Sub setValue(ByRef obj As Object)
            r.Value = obj
        End Sub

        Friend Sub setValueArray(obj As Object)
            r.Value = obj
        End Sub

        Friend Sub setValue(ByRef v As Decimal)
            r.Value = v
        End Sub

        Friend Sub setValue(ByRef v As Double)
            r.Value = v
        End Sub

        Friend Sub setValue(ByRef v As Integer)
            r.Value = v
        End Sub

        Friend Sub ClearContents()
            r.ClearContents()
        End Sub

        Friend Sub setValue(ByRef v As Date)
            r.Value = v
        End Sub

        Friend Sub setValue(ByRef s As String, ByVal rowDelta As Integer, ByVal columnDelta As Integer)
            setText(r, s, rowDelta, columnDelta)
        End Sub

        Private Sub setText(ByRef r As Range, ByRef s As String, ByVal rowDelta As Integer, ByVal columnDelta As Integer)
            r.Value = s
            If rowDelta > 0 Or columnDelta > 0 Then
                Cell(r, rowDelta, columnDelta).Merge()
            End If
        End Sub

        Friend Sub setValue(ByVal s As String, ByVal Frequency As Integer, _
        ByVal rowInterval As Integer, ByVal columnInterval As Integer)
            For i As Integer = 1 To Frequency
                Cell(r.Row + (i - 1) * rowInterval, r.Column + (i - 1) * columnInterval).Value = s
            Next
        End Sub

        Friend Sub setValue(ByRef s As String(), ByVal orient As Orientation, ByVal columnCount As Integer, ByVal rowCount As Integer)
            Dim rowDelta As Integer = 0, columnDelta As Integer = 0
            If orient = Orientation.HORIZONTAL Then
                columnDelta = columnCount
            Else
                rowDelta = rowCount
            End If
            For i As Integer = 0 To s.Length - 1
                setText(Cell(r.Row + i * rowDelta, r.Column + i * columnDelta), s(i), rowCount - 1, columnCount - 1)
            Next
        End Sub

        Friend Sub setBorders(ByVal type As BorderType, ByVal weight As BorderWeight)
            With r
                If .Columns.Count > 1 And (type = BorderType.ALL Or type = BorderType.VERTICAL) Then
                    .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlInsideVertical).Weight = XlBorderWeight.xlThin
                End If
                If .Rows.Count > 1 And (type = BorderType.ALL Or type = BorderType.HORIZONTAL) Then
                    .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlInsideHorizontal).Weight = XlBorderWeight.xlThin
                End If
                If type = BorderType.BOTTOM Then
                    .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlEdgeBottom).Weight = weight
                Else
                    .BorderAround(LineStyle:=XlLineStyle.xlContinuous, Weight:=CType(weight, XlBorderWeight))
                End If
            End With
        End Sub

        Private Function Cell(ByVal row As Integer, ByVal column As Integer) As Range
            Return Cell(row, column, row, column)
        End Function

        Private Function Cell(ByRef r As Range, ByVal rowDelta As Integer, ByVal columnDelta As Integer) As Range
            Return Cell(r.Row, r.Column, r.Row + rowDelta, r.Column + columnDelta)
        End Function

        Private Function Cell(ByVal topRow As Integer, ByVal topColumn As Integer, ByVal bottomRow As Integer, ByVal bottomColumn As Integer) As Range
            Return sheet.Range(sheet.Cells(topRow, topColumn), sheet.Cells(bottomRow, bottomColumn))
        End Function

        Friend Sub SelectArea()
            r.Select()
        End Sub

        Protected Overrides Sub Finalize()
            r = Nothing
            sheet = Nothing
            MyBase.Finalize()
        End Sub

        Public Sub New()
        End Sub

        Public Sub New(ByRef p As Point)
            Me.New(p, 0, 0)
        End Sub

        Public Sub New(ByRef p As Point, ByVal rowDelta As Integer, ByVal columnDelta As Integer)
            setWorksheet(CurrentSheet)
            setPoints(p.Row, p.Column, p.Row + rowDelta, p.Column + columnDelta)
        End Sub
    End Class
End Namespace
