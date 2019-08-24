Imports Microsoft.Office.Interop.Excel
Imports dotNet.i18n.Languager

Namespace office
    ''' <summary>
    ''' 通过COM互操作方式使用Excel API的抽象专家类。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Excelist

        Private Shared originFilePath As String

        Private Shared currentApp As Application = Nothing
        Private Shared m_currentAppHandle As Integer = -1
        Private Shared currentWorkbook As Workbook = Nothing
        Private Shared currentWorksheet As Worksheet = Nothing

        Private Shared currentArea As Area = Nothing
        Private Shared pnt As Point = Nothing

        Private Shared m_FileIsCreatedByCode As Boolean = Nothing
        Private Shared m_FilePath_NPOI As String = Nothing

        ''' <summary>
        ''' 设置Excel应用程序对应的工作目录。
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Shared WriteOnly Property FilePath() As String
            Set(ByVal path As String)
                If currentApp IsNot Nothing Then
                    currentApp.DefaultFilePath = path
                End If
            End Set
        End Property

        ''' <summary>
        ''' 获得当前打开的Excel文件所在的Excel应用程序进程句柄。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub EnterCurrentBook()
            SetCurrentApplication(CType(GetObject(, "Excel.Application"), Application))
        End Sub

        ''' <summary>
        ''' 退出当前打开的Excel文件所在的Excel应用程序进程句柄。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub ExitCurrentBook()
            SetCurrentApplication(Nothing)
        End Sub

        Friend Shared Sub SetCurrentApplication(ByRef currentApplication As Application)
            If currentApplication Is Nothing Then
                If currentApp IsNot Nothing Then
                    currentWorksheet = Nothing
                    currentArea.setWorksheet(Nothing)
                    currentArea = Nothing
                    pnt = Nothing
                    If Not m_IsReadOnly Then
                        currentWorkbook.Save()
                    End If
                    currentWorkbook = Nothing
                    With currentApp
                        .DisplayAlerts = True
                        .ScreenUpdating = True
                    End With
                    currentApp = Nothing
                End If
            Else
                currentApp = currentApplication
                With currentApp
                    .DisplayAlerts = False
                    .ScreenUpdating = False
                End With
                currentArea = New Area
                pnt = New Point
                m_FileIsCreatedByCode = False
                currentWorkbook = currentApp.ActiveWorkbook
                currentWorksheet = CType(currentWorkbook.ActiveSheet, Worksheet)
                currentArea.setWorksheet(currentWorksheet)
            End If
        End Sub

        ''' <summary>
        ''' 在后台启动一个Excel应用程序进程。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub StartApplication()
            currentApp = New Application
            m_currentAppHandle = currentApp.Hwnd
            setScreenDisplay(False)
            currentArea = New Area()
            pnt = New Point
            With currentApp
                originFilePath = .DefaultFilePath
                .StandardFontSize = 10
            End With
        End Sub

        ''' <summary>
        ''' 获取当前的Excel应用程序进程。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared ReadOnly Property CurrentApplication As Application
            Get
                Return currentApp
            End Get
        End Property

        ''' <summary>
        ''' 退出后台启动的Excel应用程序进程。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub ExitApplication()
            setScreenDisplay(True)
            currentArea = Nothing
            pnt = Nothing
            With currentApp
                .DefaultFilePath = originFilePath
                .Quit()
            End With
            currentApp = Nothing
            Dim pId As Integer
            GetWindowThreadProcessId(CType(m_currentAppHandle, IntPtr), pId)
            m_currentAppHandle = Nothing
            If pId <> 0 Then
                sys.Process.Kill(pId)
            End If
        End Sub

        <Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
        End Function

        Private Shared Sub setScreenDisplay(ByVal value As Boolean)
            If currentApp IsNot Nothing Then
                If value Then
                    CloseBook()
                End If
                With currentApp
                    .DisplayAlerts = value
                    .ScreenUpdating = value
                End With
            End If
        End Sub

        Public Shared Sub CreateBook(ByVal fileName As String, Optional ByVal extension As String = "xls")
            CloseBook()
            currentWorkbook = currentApp.Workbooks.Add
            currentWorkbook.SaveAs(
            Filename:=currentApp.DefaultFilePath + "\" + fileName + "." + extension)
            m_IsReadOnly = False
            m_FileIsCreatedByCode = True
        End Sub

        Private Shared m_IsReadOnly As Boolean = False

        ''' <summary>
        ''' 打开Excel当前工作目录下的指定Excel文件。
        ''' </summary>
        ''' <param name="fileName">Excel文件名称，不包含文件目录路径或后缀名。</param>
        ''' <param name="extension">Excel文件后缀名。</param>
        ''' <param name="IsReadOnly" >是否以只读方式打开Excel文件。</param>
        ''' <remarks></remarks>
        Public Shared Sub OpenBook(ByVal fileName As String, Optional ByVal extension As String = "xls",
                                   Optional ByVal IsReadOnly As Boolean = False)
            CloseBook()
            currentWorkbook = currentApp.Workbooks.Open(
            Filename:=currentApp.DefaultFilePath + "\" + fileName + "." + extension,
            ReadOnly:=IsReadOnly)
            m_IsReadOnly = IsReadOnly
            m_FileIsCreatedByCode = False
        End Sub

        ''' <summary>
        ''' 按照全路径名称打开文件。
        ''' </summary>
        ''' <param name="fileName">全路径文件名。</param>
        ''' <remarks></remarks>
        Public Shared Sub OpenBookByFullName(ByVal fileName As String)
            CloseBook()
            currentWorkbook = currentApp.Workbooks.Open(
            Filename:=fileName, ReadOnly:=True)
            m_IsReadOnly = True
            m_FileIsCreatedByCode = False
        End Sub

        ''' <summary>
        ''' 按照全路径名称保存文件。
        ''' </summary>
        ''' <param name="fileName">全路径文件名。</param>
        ''' <remarks></remarks>
        Public Shared Sub SaveBookByFullName(fileName As String)
            currentWorkbook.SaveAs(Filename:=fileName, AccessMode:=XlSaveAsAccessMode.xlNoChange,
                ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges)
        End Sub

        Public Shared Sub FreezePanes(ByRef r As Area)
            r.SelectColumn()
            currentApp.ActiveWindow.FreezePanes = True
            Cell(1, 1).SelectArea()
        End Sub

        Friend Shared ReadOnly Property CurrentBook() As Workbook
            Get
                Return currentWorkbook
            End Get
        End Property

        ''' <summary>
        ''' 关闭当前工作簿。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseBook()
            If currentWorkbook IsNot Nothing Then
                CloseCurrentSheet()
                If m_FileIsCreatedByCode Then
                    For i As Integer = 1 To currentApp.SheetsInNewWorkbook
                        CType(currentWorkbook.Worksheets(1), Worksheet).Delete()
                    Next
                    CType(currentWorkbook.Worksheets(1), Worksheet).Activate()
                End If
                If Not m_IsReadOnly Then
                    currentWorkbook.Save()
                End If
                currentWorkbook.Close()
                currentWorkbook = Nothing
            End If
        End Sub

        Friend Shared Function Min(ByRef a As Array) As Double
            Return ExcelFunction.Min(a)
        End Function

        Friend Shared Function Max(ByRef a As Array) As Double
            Return ExcelFunction.Max(a)
        End Function

        Public Shared Function Avg(ByRef a As Area) As Double
            Return ExcelFunction.Average(a.Range)
        End Function

        ''' <summary>
        ''' 返回作为概率和自由度函数的T分布的t值。
        ''' </summary>
        ''' <param name="probability">对应于双尾T分布的概率。</param>
        ''' <param name="degrees_freedom">分布的自由度。</param>
        ''' <returns></returns>
        Public Shared Function TInv(probability As Double, degrees_freedom As Integer) As Double
            Return ExcelFunction.TInv(probability, degrees_freedom)
        End Function

        Private Shared ReadOnly Property ExcelFunction() As WorksheetFunction
            Get
                Return currentApp.WorksheetFunction
            End Get
        End Property

        Private Shared Sub FitToPage()
            With currentWorksheet.PageSetup
                .CenterFooter = "&9&P/&N"
                .CenterHorizontally = True
                .CenterVertically = True
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .LeftMargin = 30
                .RightMargin = 30
                .TopMargin = 40
                .BottomMargin = 40
                .HeaderMargin = 25
                .FooterMargin = 25
            End With
        End Sub

        Enum Orientation As Integer
            VERTICAL = XlPageOrientation.xlPortrait
            HORIZONTAL = XlPageOrientation.xlLandscape
        End Enum

        Public Shared Sub setPageOrientation(ByVal orient As Orientation)
            currentWorksheet.PageSetup.Orientation = CType(orient, XlPageOrientation)
        End Sub

        Private Shared Sub CloseCurrentSheet()
            If currentWorksheet IsNot Nothing Then
                If m_FileIsCreatedByCode Then
                    currentWorksheet.UsedRange.Font.Name = FontName()
                End If
                If Not m_IsReadOnly Then
                    currentWorkbook.Save()
                End If
                currentWorksheet = Nothing
                currentArea.setWorksheet(Nothing)
            End If
        End Sub

        ''' <summary>
        ''' 从当前Excel文件中删除一个指定名称的工作表。
        ''' </summary>
        ''' <param name="name">工作表名称，如“Sheet1”。</param>
        ''' <remarks></remarks>
        Public Shared Sub RemoveSheet(ByVal name As String)
            CType(currentWorkbook.Worksheets(name), Worksheet).Delete()
        End Sub

        ''' <summary>
        ''' 选中当前Excel文件中一个指定名称的工作表。
        ''' </summary>
        ''' <param name="name">工作表名称，如“Sheet1”。</param>
        ''' <remarks></remarks>
        Public Shared Sub ActivateSheet(ByVal name As String)
            CType(currentWorkbook.Worksheets(name), Worksheet).Select()
            currentWorkbook.Save()
        End Sub

        ''' <summary>
        ''' 重命名当前Excel文件中的一个工作表名称。
        ''' </summary>
        ''' <param name="oldName">工作表原名称。</param>
        ''' <param name="newName">工作表新名称。</param>
        ''' <remarks></remarks>
        Public Shared Sub RenameSheet(ByVal oldName As String, ByVal newName As String)
            CType(currentWorkbook.Worksheets(oldName), Worksheet).Name = newName
        End Sub

        ''' <summary>
        ''' 打开当前Excel文件的指定工作表。
        ''' </summary>
        ''' <param name="name">工作表名称，如“Sheet1”。</param>
        ''' <remarks>系统会自动保存Excel文件。</remarks>
        Public Shared Sub OpenSheet(ByVal name As String)
            CloseCurrentSheet()
            currentWorksheet = CType(currentWorkbook.Worksheets(name), Worksheet)
            currentArea.setWorksheet(currentWorksheet)
        End Sub

        ''' <summary>
        ''' 将指定Excel文件中的工作表复制到当前Excel文件中并命名为新的工作表，并将其置于当前Excel文件的末尾。
        ''' </summary>
        ''' <param name="FileName">指定的Excel文件。</param>
        ''' <param name="OldSheetName">需要被复制的工作表的名称。</param>
        ''' <param name="NewSheetName">新工作表的名称。</param>
        Public Shared Sub CopySheet(FileName As String, OldSheetName As String, NewSheetName As String)
            Dim tempBook As Workbook = currentApp.Workbooks.Open(
            Filename:=FileName, ReadOnly:=True)
            Dim tempSheet As Worksheet = CType(tempBook.Sheets(OldSheetName), Worksheet)
            tempSheet.Select()
            tempSheet.Copy(After:=currentWorkbook.Sheets(SheetsCount))
            tempSheet = CType(currentWorkbook.Sheets(SheetsCount), Worksheet)
            tempSheet.Name = NewSheetName
            tempBook.Close(SaveChanges:=False)
            tempBook = Nothing
            tempSheet = Nothing
        End Sub

        ''' <summary>
        ''' 将当前Excel文件的指定工作表复制为新的工作表，并将其置于原有工作表之后。
        ''' </summary>
        ''' <param name="OldName">现有工作表的名称。</param>
        ''' <param name="NewName">新工作表的名称。</param>
        ''' <remarks></remarks>
        Public Shared Sub CopySheet(OldName As String, NewName As String)
            Dim csheet As Worksheet = CType(currentWorkbook.Sheets(OldName), Worksheet)
            csheet.Select()
            csheet.Copy(After:=currentWorkbook.Sheets(OldName))
            Dim index As Integer = csheet.Index + 1
            csheet = CType(currentWorkbook.Sheets(index), Worksheet)
            csheet.Name = NewName
            index = Nothing
            csheet = Nothing
        End Sub

        Friend Shared Function GetChartByName(ByVal name As String) As Chart
            Return CType(currentWorkbook.Charts(name), Chart)
        End Function

        ''' <summary>
        ''' 获得当前Excel文件中的工作表总数。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property SheetsCount() As Integer
            Get
                Return currentWorkbook.Sheets.Count - currentWorkbook.Charts.Count
            End Get
        End Property

        ''' <summary>
        ''' 判断当前Excel文件中是否存在指定名称的工作表。
        ''' </summary>
        ''' <param name="name">指定的工作表名称。</param>
        ''' <returns>若存在指定名称的工作表，则返回true；否则返回false。</returns>
        ''' <remarks></remarks>
        Public Shared Function SheetIsContained(ByVal name As String) As Boolean
            For Each a As Worksheet In currentWorkbook.Worksheets
                If a.Name = name Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' 检测当前Excel文件中是否存在包含指定关键词的工作表。
        ''' </summary>
        ''' <param name="keyword">指定的关键词。</param>
        ''' <returns>若工作表名称包含指定的关键词，则打开该工作表并返回true；否则返回false。</returns>
        ''' <remarks></remarks>
        Public Shared Function SheetNameContainsKeyword(ByVal keyword As String) As Boolean
            For Each a As Worksheet In currentWorkbook.Worksheets
                If a.Name.Contains(keyword) Then
                    OpenSheet(a.Name)
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' 打开当前Excel文件的指定工作表。
        ''' </summary>
        ''' <param name="index">工作表的索引号,索引号以1开始。</param>
        ''' <remarks>系统会自动保存Excel文件。</remarks>
        Public Shared Sub OpenSheet(ByVal index As Integer)
            CloseCurrentSheet()
            currentWorksheet = CType(currentWorkbook.Worksheets(index), Worksheet)
            currentArea.setWorksheet(currentWorksheet)
        End Sub

        ''' <summary>
        ''' 获得当前Excel文件中的活动工作表名称。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property SheetName() As String
            Get
                Return currentWorksheet.Name
            End Get
        End Property

        ''' <summary>
        ''' 将图表移至指定的工作表之后。
        ''' </summary>
        ''' <param name="ChartName">将移动的图表名称。</param>
        ''' <param name="Index">指定的工作表序号。</param>
        ''' <remarks></remarks>
        Public Shared Sub MoveChartAfterSheet(ByVal ChartName As String, ByVal Index As Integer)
            CType(currentWorkbook.Sheets(ChartName), Chart).Move(After:=CType(currentWorkbook.Sheets(Index), Worksheet))
        End Sub

        ''' <summary>
        ''' 在当前Excel文件中增加一个指定名称的工作表，并打开该工作表。
        ''' </summary>
        ''' <param name="name">工作表名称，如“Sheet1”。</param>
        ''' <remarks></remarks>
        Private Shared Sub AddNewSheet(ByVal name As String)
            CloseCurrentSheet()
            currentWorksheet = CType(currentWorkbook.Worksheets.Add(
                    After:=currentWorkbook.Worksheets(currentWorkbook.Worksheets.Count)), Worksheet)
            currentWorksheet.Name = name
            currentArea.setWorksheet(currentWorksheet)
        End Sub

        ''' <summary>
        ''' 在当前Excel文件中增加一个指定名称的工作表，并打开该工作表。
        ''' </summary>
        ''' <param name="name">工作表名称，如“Sheet1”。</param>
        ''' <param name="IsAutoIncrement">当指定名称的工作表已存在时，是否自动递增工作表名称后的序号。</param>
        ''' <remarks>若IsAutoIncrement设置为True，当指定名称“MySheet”已经存在时，会自动按照“MySheet1"、
        ''' “MySheet2”...的顺序检测并自动增加新的工作表。</remarks>
        Public Shared Sub AddSheet(ByVal name As String, Optional ByVal IsAutoIncrement As Boolean = True)
            If SheetIsContained(name) Then
                If IsAutoIncrement Then
                    Dim i As Integer = 1
                    While SheetIsContained(name + CStr(i))
                        i = i + 1
                    End While
                    AddNewSheet(name + CStr(i))
                    i = Nothing
                Else
                    RemoveSheet(name)
                    AddNewSheet(name)
                End If
            Else
                AddNewSheet(name)
            End If
        End Sub

        Friend Shared ReadOnly Property CurrentSheet() As Worksheet
            Get
                Return currentWorksheet
            End Get
        End Property

        Public Shared Sub setRowHeight(ByVal row As Integer, ByVal pixel As Integer)
            CType(currentWorksheet.Rows(row), Range).RowHeight = pixel
        End Sub

        ''' <summary>
        ''' 获得指定的一个Excel单元格。
        ''' </summary>
        ''' <param name="row">单元格行号。</param>
        ''' <param name="column">单元格列号。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Cell(ByVal row As Integer, ByVal column As Integer) As Area
            Return Cell(row, column, row, column)
        End Function

        ''' <summary>
        ''' 获得指定的一个Excel单元格。
        ''' </summary>
        ''' <param name="column">单元格列号。</param>
        ''' <returns></returns>
        ''' <remarks>单元格行号由AddressRow属性指定。</remarks>
        Public Shared Function CellColumn(ByVal column As Integer) As Area
            Return Cell(m_AddressRow, column, m_AddressRow, column)
        End Function

        ''' <summary>
        ''' 获得指定的一个Excel单元格。
        ''' </summary>
        ''' <param name="row">单元格行号。</param>
        ''' <returns></returns>
        ''' <remarks>单元格列号由AddressColumn属性指定。</remarks>
        Public Shared Function CellRow(ByVal row As Integer) As Area
            Return Cell(row, m_AddressColumn, row, m_AddressColumn)
        End Function

        ''' <summary>
        ''' 获得指定的一个Excel单元格区域。
        ''' </summary>
        ''' <param name="address">以“字母列号+数字行号”的A1样式表示的单元格地址。
        ''' 可以为单个单元格（如“A7”），也可以为一个矩形区域（如“A7:B15”）
        ''' </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Cell(ByVal address As String) As Area
            currentArea.setRange(address)
            Return currentArea
        End Function

        ''' <summary>
        ''' 将Excel工作表中指定单元格区域的内容复制并粘贴到同一工作表的其他单元格区域中。
        ''' </summary>
        ''' <param name="sourceAreaAddress">以A1样式表示的待复制的单元格区域地址。</param>
        ''' <param name="targetAreaAddress">以A1样式表示的目标单元格地址。</param>
        ''' <remarks></remarks>
        Private Shared Sub CopyAndPaste(ByVal sourceAreaAddress As String, ByVal targetAreaAddress As String)
            currentWorksheet.Range(sourceAreaAddress).Select()
            CType(currentApp.Selection, Range).Copy()
            currentWorksheet.Range(targetAreaAddress).Select()
            currentWorksheet.Paste()
        End Sub

        ''' <summary>
        ''' 将Excel工作表中指定单元格区域的内容复制并粘贴到同一工作表的其他单元格区域中。
        ''' </summary>
        ''' <param name="sourceAddress">以A1样式表示的源数据单元格区域地址。</param>
        ''' <param name="numberOfCopies">以源数据单元格大小为基准计算的需要复制的单元格倍数。</param>
        ''' <remarks>若源数据为1行，则向下复制；若源数据为1列，则向右复制；若源数据为其他情况，则不复制。</remarks>
        Public Shared Sub CopyMultipleCells(ByVal sourceAddress As String, ByVal numberOfCopies As Integer)
            Dim r As Range = currentWorksheet.Range(sourceAddress)
            Dim targetAddress As String = Nothing
            If r.Rows.Count = 1 And r.Columns.Count >= 1 Then
                targetAddress = Address(Base(r.Row + 1, r.Column), numberOfCopies - 1, r.Columns.Count - 1)
            ElseIf r.Columns.Count = 1 And r.Rows.Count > 1 Then
                targetAddress = Address(Base(r.Row, r.Column + 1), r.Rows.Count - 1, numberOfCopies - 1)
            End If
            r = Nothing
            If targetAddress = Nothing Then
                Return
            End If
            CopyAndPaste(sourceAddress, targetAddress)
        End Sub

        ''' <summary>
        ''' 将Excel工作表中指定单元格区域的内容自动填充到同一工作表的其他单元格区域中。
        ''' </summary>
        ''' <param name="sourceAreaAddress">以A1样式表示的待复制的单元格区域地址</param>
        ''' <param name="targetAreaAddress">以A1样式表示的目标单元格地址</param>
        ''' <remarks></remarks>
        Public Shared Sub AutoFill(ByVal sourceAreaAddress As String, ByVal targetAreaAddress As String)
            currentWorksheet.Range(sourceAreaAddress).Select()
            CType(currentApp.Selection, Range).AutoFill(currentWorksheet.Range(targetAreaAddress))
        End Sub

        ''' <summary>
        ''' 根据坐标得到一个Excel单元格点。
        ''' </summary>
        ''' <param name="row">单元格行号。</param>
        ''' <param name="column">单元格列号。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Base(ByVal row As Integer, ByVal column As Integer) As Point
            pnt.Row = row
            pnt.Column = column
            Return pnt
        End Function

        ''' <summary>
        ''' 获得指定的一个Excel单元格区域。
        ''' </summary>
        ''' <param name="p">单元格区域的左上角顶点。</param>
        ''' <param name="rowDelta">单元格区域的最下行与左上角顶点所在行之差。</param>
        ''' <param name="columnDelta">单元格区域的最右列与左上角顶点所在列之差。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Cell(ByRef p As Point, ByVal rowDelta As Integer, ByVal columnDelta As Integer) As Area
            Return Cell(p.Row, p.Column, p.Row + rowDelta, p.Column + columnDelta)
        End Function

        Private Shared Function Cell(ByVal topRow As Integer, ByVal topColumn As Integer, ByVal bottomRow As Integer, ByVal bottomColumn As Integer) As Area
            currentArea.setPoints(topRow, topColumn, bottomRow, bottomColumn)
            Return currentArea
        End Function

        ''' <summary>
        ''' 获得指定的一个Excel单元格地址。
        ''' </summary>
        ''' <param name="row">单元格行号。</param>
        ''' <param name="column">单元格列号。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Address(ByVal row As Integer, ByVal column As Integer) As String
            Return Address(row, column, row, column)
        End Function

        ''' <summary>
        ''' 获得指定的一个Excel单元格地址。
        ''' </summary>
        ''' <param name="column">单元格列号。</param>
        ''' <returns></returns>
        ''' <remarks>单元格行号由AddressRow属性指定。</remarks>
        Public Shared Function AddrColumn(ByVal column As Integer) As String
            Return Address(m_AddressRow, column, m_AddressRow, column)
        End Function

        ''' <summary>
        ''' 获得指定的一个Excel单元格地址。
        ''' </summary>
        ''' <param name="row">单元格行号。</param>
        ''' <returns></returns>
        ''' <remarks>单元格列号由AddressColumn属性指定。</remarks>
        Public Shared Function AddrRow(ByVal row As Integer) As String
            Return Address(row, m_AddressColumn, row, m_AddressColumn)
        End Function

        Public Shared Function Address(ByRef p As Point, ByVal rowDelta As Integer, ByVal columnDelta As Integer) As String
            Return Address(p.Row, p.Column, p.Row + rowDelta, p.Column + columnDelta)
        End Function

        Private Shared Function Address(ByVal topRow As Integer, ByVal topColumn As Integer, ByVal bottomRow As Integer, ByVal bottomColumn As Integer) As String
            Return currentWorksheet.Range(currentWorksheet.Cells(topRow, topColumn), currentWorksheet.Cells(bottomRow, bottomColumn)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End Function

        Private Shared m_AddressRow As Integer = Nothing
        Private Shared m_AddressColumn As Integer = Nothing

        ''' <summary>
        ''' 当前Excel单元格所在的行。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Property AddressRow() As Integer
            Get
                Return m_AddressRow
            End Get
            Set(ByVal value As Integer)
                m_AddressRow = value
            End Set
        End Property

        ''' <summary>
        ''' 当前Excel单元格所在的列。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Property AddressColumn As Integer
            Get
                Return m_AddressColumn
            End Get
            Set(ByVal value As Integer)
                m_AddressColumn = value
            End Set
        End Property

    End Class
End Namespace
