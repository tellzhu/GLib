Imports Microsoft.Office.Interop.Excel

Imports dotNet.math
Imports dotNet.office.AreaExpert
Imports dotNet.office.Excelist
Imports dotNet.text.StringHandler

Namespace office
	Public Class Chartist

		Private Shared Function Union(ByRef r() As Area) As Range
			Dim a As Area = New Area(Base(r(0).Row, r(0).Column), r(0).Range.Rows.Count - 1, r(0).Range.Columns.Count - 1)
			If r.Length > 1 Then
				For i As Integer = 1 To r.Length - 1
					a.Union(r(i))
				Next
			End If
			Return a.Range
		End Function

		Private Shared Sub setSourceData(ByRef SrcData() As Area)
			With CurrentChart
				.SetSourceData(Union(SrcData), PlotBy:=XlRowCol.xlRows)
				Dim count As Integer = CType(.SeriesCollection, SeriesCollection).Count
				If count < SrcData.Length Then
					For i As Integer = 1 To SrcData.Length - count
						CType(.SeriesCollection, SeriesCollection).NewSeries()
					Next
				ElseIf count > SrcData.Length Then
					For i As Integer = 1 To count - SrcData.Length
						CType(.SeriesCollection(1), Series).Delete()
					Next
				End If
			End With
		End Sub

		Private Shared Sub setData(ByRef xValues As Area, ByRef names() As String, ByRef values() As Area)
			setSourceData(values)
			For i As Integer = 0 To names.Length - 1
				setSeries(i + 1, xValues, names(i), values(i))
			Next
		End Sub

		Private Shared Sub setSeries(ByVal index As Integer, ByRef xValues As Area, ByVal name As String, ByRef values As Area)
			With CType(CurrentChart.SeriesCollection(index), Series)
				.Values = "=(" + RangeAddress(values.Range) + ")"
				.XValues = xValues.Range
				.Name = name
			End With
		End Sub

		Private Shared Function RangeAddress(ByRef r As Range) As String
			Dim s() As String = Split(r.Address(ReferenceStyle:=XlReferenceStyle.xlR1C1), ",")
			Dim str As String = "'" + CurrentSheet.Name + "'!" + s(0)
			For i As Integer = 1 To s.Length - 1 Step 1
				str = str + ",'" + CurrentSheet.Name + "'!" + s(i)
			Next
			s = Nothing
			Return str
		End Function

        Private Shared Sub ScaleAxis()
            With CurrentChart
                Dim maxValue As Long = Long.MinValue, minValue As Long = Long.MaxValue

                Dim count As Integer = CType(.SeriesCollection, SeriesCollection).Count
                Dim val As Integer
                For k As Integer = 1 To count
                    With CType(.SeriesCollection(k), Series)
                        Dim arr As Array = CType(.Values, Array)
                        val = CInt(System.Math.Ceiling(Max(arr)))
                        If val > maxValue Then
                            maxValue = val
                        End If
                        val = CInt(System.Math.Floor(Min(arr)))
                        If val < minValue Then
                            minValue = val
                        End If
                        arr = Nothing
                    End With
                Next

                Dim scale As Integer = 1000
                Dim maxVal As Long = maxValue * scale, minVal As Long = minValue * scale
                Dim unit As Long = (maxVal - minVal) \ 5
                If unit < 1 Then
                    unit = 1
                End If
                Dim i As Long = 1
                Do While unit \ 10 > 0
                    unit = unit \ 10
                    minVal = minVal \ 10
                    i = i * 10
                Loop
                unit = unit * i
                minVal = minVal * i
                Do While minVal > minValue * scale
                    minVal = minVal - unit
                Loop
                i = minVal
                Do While i < maxValue * scale
                    i = i + unit
                Loop

                With CType(CurrentChart.Axes(XlAxisType.xlValue), Axis)
                    .MinimumScale = minVal / scale
                    .MaximumScale = i / scale
                    .MinorUnit = unit / scale
                    .MajorUnit = 2 * unit / scale
                    .Crosses = XlAxisCrosses.xlAxisCrossesCustom
                    .CrossesAt = minVal / scale
                    If unit / scale <= 5 Then
                        .TickLabels.NumberFormatLocal = "0.00_ "
                    Else
                        .TickLabels.NumberFormatLocal = "0_ "
                    End If
                End With
            End With
        End Sub

        Private Shared Sub ScaleLineChartAxis()
            With m_currentChart
                Dim maxValue As Long = Long.MinValue, minValue As Long = Long.MaxValue

                Dim count As Integer = CType(.SeriesCollection, SeriesCollection).Count
                Dim val As Integer
                For k As Integer = 1 To count
                    With CType(.SeriesCollection(k), Series)
                        Dim arr As Array = CType(.Values, Array)
                        val = CInt(System.Math.Ceiling(Max(arr)))
                        If val > maxValue Then
                            maxValue = val
                        End If
                        val = CInt(System.Math.Floor(Min(arr)))
                        If val < minValue Then
                            minValue = val
                        End If
                        arr = Nothing
                    End With
                Next

                Dim scale As Integer = 1000
                Dim maxVal As Long = maxValue * scale, minVal As Long = minValue * scale
                Dim unit As Long = (maxVal - minVal) \ 5
                If unit < 1 Then
                    unit = 1
                End If
                Dim i As Long = 1
                Do While unit \ 10 > 0
                    unit = unit \ 10
                    minVal = minVal \ 10
                    i = i * 10
                Loop
                unit = unit * i
                minVal = minVal * i
                Do While minVal > minValue * scale
                    minVal = minVal - unit
                Loop
                i = minVal
                Do While i < maxValue * scale
                    i = i + unit
                Loop

                With CType(m_currentChart.Axes(XlAxisType.xlValue), Axis)
                    .MinimumScale = minVal / scale
                    .MaximumScale = i / scale
                    .MinorUnit = unit / scale
                    .MajorUnit = 2 * unit / scale
                    .Crosses = XlAxisCrosses.xlAxisCrossesCustom
                    .CrossesAt = minVal / scale
                    If unit / scale <= 5 And System.Math.Abs(.MinimumScale) <= 5 And System.Math.Abs(.MaximumScale) <= 5 Then
                        .TickLabels.NumberFormatLocal = "0.00_ "
                    Else
                        .TickLabels.NumberFormatLocal = "0_ "
                    End If
                End With
            End With
        End Sub

		Private Shared ReadOnly Property CurrentChartObject() As ChartObject
			Get
				Return CType(CurrentSheet.ChartObjects(CType(CurrentSheet.ChartObjects, ChartObjects).Count), ChartObject)
			End Get
		End Property

		Private Shared ReadOnly Property CurrentChart() As Chart
			Get
				Return CurrentChartObject.Chart
			End Get
		End Property

		Private Shared Sub setLocation(ByRef r As Range)
			With CurrentChartObject
				.Left = CDec(r.Left)
				.Top = CDec(r.Top)
				.Width = CDec(r.Width)
				.Height = CDec(r.Height)
			End With
		End Sub

		Public Shared Function AddChart(ByRef Location As Area, ByVal ChartTitle As String, _
		  ByVal AxisX_Name As String, ByRef AxisX_Data As Area, _
		  ByVal AxisY_Name As String, ByRef DataNames() As String, _
		  ByRef AxisY_Data() As Area, ByVal TotalDataPoints As Integer) As Boolean

			If IsEnoughData(AxisY_Data, TotalDataPoints) Then
				AddChart(Location)
				setData(AxisX_Data, DataNames, AxisY_Data)
				setTitles(ChartTitle, AxisX_Name, AxisY_Name)
				Return True
			Else
				Return False
			End If
        End Function

        Private Shared m_currentChart As Chart = Nothing
        Private Shared m_currentSeriesCollection As SeriesCollection = Nothing
        Private Shared m_InitialSeriesCount As Integer = Nothing

        ''' <summary>
        ''' 打开当前Excel文件的指定图表。
        ''' </summary>
        ''' <param name="name">图表名称，如“Chart1”。</param>
        ''' <remarks></remarks>
        Private Shared Sub OpenChart(ByVal name As String)
            CloseCurrentChart()
            m_currentChart = GetChartByName(name)
            m_currentSeriesCollection = CType(m_currentChart.SeriesCollection, SeriesCollection)
            m_InitialSeriesCount = m_currentSeriesCollection.Count
        End Sub

        Private Shared Sub OpenChart(ByVal SheetName As String, ByVal ChartName As String)
            CloseCurrentChart()
            OpenSheet(SheetName)
            m_currentChart = CType(CurrentSheet.ChartObjects(ChartName), ChartObject).Chart
            m_currentSeriesCollection = CType(m_currentChart.SeriesCollection, SeriesCollection)
            m_InitialSeriesCount = m_currentSeriesCollection.Count
        End Sub

        Private Shared Sub DeleteExistedOldSeries()
            For i As Integer = 1 To m_InitialSeriesCount
                m_currentSeriesCollection.Item(1).Delete()
            Next
            m_InitialSeriesCount = Nothing
        End Sub

        Private Shared Sub CloseCurrentChart()
            m_currentSeriesCollection = Nothing
            m_currentChart = Nothing
        End Sub

        Private Shared Sub setTitle(ByVal ChartTitle As String)
            m_currentChart.ChartTitle.Text = ChartTitle
        End Sub

        ''' <summary>
        ''' 图表类型，目前包括折线、柱状图、饼图、百分比堆积图和普通堆积图。
        ''' </summary>
        ''' <remarks></remarks>
        Enum ChartType
            LieChart
            Histogram
            PieChart
            StowageDiagram
            AccumulatedStowageDiagram
        End Enum

        ''' <summary>
        ''' 将输入的数据系列按规定图表格式输出。
        ''' </summary>
        ''' <param name="SheetName" >图表所在的工作表名称。</param>
        ''' <param name="ChartName" >图表所在的工作表图对象中的名称。 </param>
        ''' <param name="ChrtType">图表类型，目前包括LieChart（折线图）、Histogram（柱状图）、
        ''' PieChart（饼图）、StowageDiagram（百分比堆积图）、AccumulatedStowageDiagram（堆积图）。</param>
        ''' <param name="ChartTitle">图表标题。</param>
        ''' <param name="SeriesNames">数据系列名称。</param>
        ''' <param name="mt">已存储待输出数据的矩阵。</param>
        ''' <remarks>矩阵的第1列将作为图表的X轴，矩阵的第2列对应数据系列名称的第1个值，...，
        ''' 矩阵的第N列对应数据系列名称的第N-1个值。</remarks>
        Public Shared Sub PrintChart(ByVal SheetName As String, ByVal ChartName As String, ByVal chrtType As ChartType, ByVal ChartTitle As String, _
                             ByRef SeriesNames() As String, ByRef mt As Matrix(Of String))
            OpenChart(SheetName, ChartName)
            PrintChart(chrtType, ChartTitle, SeriesNames, mt)
            CloseCurrentChart()
        End Sub

        ''' <summary>
        ''' 将指定的图表拷贝到剪贴板。
        ''' </summary>
        ''' <param name="ChartSheetName">图表所在的工作表名称。</param>
        ''' <remarks></remarks>
        Public Shared Sub CopyChart(ByVal ChartSheetName As String)
            m_currentChart = GetChartByName(ChartSheetName)
            m_currentChart.ChartArea.Copy()
            m_currentChart = Nothing
        End Sub

        ''' <summary>
        ''' 将输入的数据系列按规定图表格式输出。
        ''' </summary>
        ''' <param name="ChartSheetName" >图表所在的工作表名称。</param>
        ''' <param name="ChrtType">图表类型，目前包括LieChart（折线图）、Histogram（柱状图）、
        ''' PieChart（饼图）、StowageDiagram（百分比堆积图）、AccumulatedStowageDiagram（堆积图）。</param>
        ''' <param name="ChartTitle">图表标题。</param>
        ''' <param name="SeriesNames">数据系列名称。</param>
        ''' <param name="mt">已存储待输出数据的矩阵。</param>
        ''' <remarks>矩阵的第1列将作为图表的X轴，矩阵的第2列对应数据系列名称的第1个值，...，
        ''' 矩阵的第N列对应数据系列名称的第N-1个值。</remarks>
        Public Shared Sub PrintChart(ByVal ChartSheetName As String, ByVal chrtType As ChartType, ByVal ChartTitle As String, _
                                     ByRef SeriesNames() As String, ByRef mt As Matrix(Of String))
            OpenChart(ChartSheetName)
            PrintChart(chrtType, ChartTitle, SeriesNames, mt)
            CloseCurrentChart()
        End Sub

        Private Shared Sub PrintPieChart(ByRef mt As Matrix(Of String))
            If mt.ColumnsCount <> 2 Then
                Return
            End If
            Dim row As Integer = mt.RowsCount
            Dim name As String = "", value As String = ""
            For i As Integer = 1 To row - 1
                name = name + """" + mt.Cell(i, 1) + ""","
                value = value + mt.Cell(i, 2) + ","
            Next
            name += """" + mt.Cell(row, 1) + """"
            value += mt.Cell(row, 2)
            m_currentSeriesCollection.Item(1).Values = "={" + value + "}"
            m_currentSeriesCollection.Item(1).XValues = "={" + name + "}"
            row = Nothing
            name = Nothing
            value = Nothing
        End Sub

        Private Shared Sub PrintChart(ByVal chrtType As ChartType, ByVal ChartTitle As String, _
                             ByRef SeriesNames() As String, ByRef mt As Matrix(Of String))
            If chrtType = ChartType.PieChart Then
                PrintPieChart(mt)
                Return
            End If

            setTitle(ChartTitle)
            Dim row As Integer = mt.RowsCount, column As Integer = mt.ColumnsCount
            Dim s(column - 1) As String
            For i As Integer = 0 To column - 1
                s(i) = ""
            Next
            For i As Integer = 1 To row - 1
                For j As Integer = 1 To column
                    s(j - 1) = s(j - 1) + NumberStringTrim(CStr(mt.Cell(i, j))) + ","
                Next
            Next
            For j As Integer = 1 To column
                s(j - 1) = s(j - 1) + NumberStringTrim(CStr(mt.Cell(row, j)))
            Next

            Dim cnt As Integer = Nothing
            For i As Integer = 1 To column - 1
                m_currentSeriesCollection.NewSeries()
                cnt = m_currentSeriesCollection.Count
                m_currentSeriesCollection.Item(cnt).Formula = "=SERIES(""" + SeriesNames(i - 1) + """,,{" + s(i) + "}," + CStr(cnt) + ")"
            Next
            m_currentSeriesCollection.Item(cnt).XValues = "={" + s(0) + "}"
            cnt = Nothing
            row = Nothing
            column = Nothing
            Array.Clear(s, 0, s.Length)
            s = Nothing

            DeleteExistedOldSeries()
            If chrtType = ChartType.LieChart Then
                ScaleLineChartAxis()
            End If
        End Sub

        Private Shared Sub AddChart(ByRef Location As Area)
            CurrentBook.Save()
            Dim currentWorkchart As Chart = CType(CurrentBook.Charts.Add, Chart)
            With currentWorkchart
                .HasTitle = False
                .ChartType = XlChartType.xlLineMarkers

                .HasAxis(XlAxisType.xlCategory, XlAxisGroup.xlPrimary) = True
                .HasAxis(XlAxisType.xlValue, XlAxisGroup.xlPrimary) = True

                CType(.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary), Axis).CategoryType = XlCategoryType.xlAutomaticScale
                With CType(.Axes(XlAxisType.xlCategory), Axis)
                    .HasMajorGridlines = False
                    .HasMinorGridlines = False
                End With
                With CType(.Axes(XlAxisType.xlValue), Axis)
                    .HasMajorGridlines = True
                    .HasMinorGridlines = False
                    .ReversePlotOrder = False
                    .ScaleType = XlScaleType.xlScaleLinear
                End With

                .HasLegend = True
                With .Legend
                    .Position = XlLegendPosition.xlLegendPositionBottom
                    .AutoScaleFont = False
                    .Font.Size = 9
                End With

                .HasDataTable = False
                .Location(XlChartLocation.xlLocationAsObject, CurrentSheet.Name)
            End With
            currentWorkchart = Nothing
            setLocation(Location.Range)
        End Sub

        Private Shared Sub setTitles(ByVal ChartTitle As String, ByVal X_Name As String, ByVal Y_Name As String)
            With CurrentChart
                .HasTitle = True

                With .ChartTitle
                    .Text = ChartTitle
                    .AutoScaleFont = False
                    .Font.Size = 9
                End With

                Dim axis As Axis = CType(.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary), Axis)
                With axis
                    .HasTitle = True
                    With .AxisTitle
                        .Characters.Text = X_Name
                        .AutoScaleFont = False
                        .Font.Size = 9
                    End With
                    With .TickLabels
                        .AutoScaleFont = False
                        .Font.Size = 9
                    End With
                End With

                axis = CType(.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary), Axis)
                With axis
                    .HasTitle = True
                    With .AxisTitle
                        .Characters.Text = Y_Name
                        .AutoScaleFont = False
                        .Font.Size = 9
                    End With
                    With .TickLabels
                        .AutoScaleFont = False
                        .Font.Size = 9
                    End With
                End With
                axis = Nothing
            End With
            ScaleAxis()
        End Sub
    End Class
End Namespace
