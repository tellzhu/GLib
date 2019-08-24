Imports System.Globalization
Imports dotNet.i18n.Languager

Namespace time
    Public Class DateExpert
        ''' <summary>
        ''' 将YYYYMMDD格式的字符串转换为日期类型。
        ''' </summary>
        ''' <param name="value">以YYYYMMDD格式表达的日期字符串。</param>
        ''' <returns></returns>
        Public Shared Function DateValue(ByVal value As String) As Date
            Return CDate(value.Substring(0, 4) + "-" + value.Substring(4, 2) + "-" + value.Substring(6, 2))
        End Function

        Public Shared Function MonthValue(ByVal value As Date) As Integer
            Return Year(value) * 12 + Month(value)
        End Function

        Public Shared Function MonthValue(ByVal value As String) As Integer
            Return MonthValue(CDate(value))
        End Function

        Public Shared Function FirstDayOfMonth(ByVal monthValue As Integer) As Date
            Return CDate(YearPart(monthValue) & "-" & MonthPart(monthValue) & "-1")
        End Function

        Public Shared Function FirstDayOfMonth(ByVal value As String) As Date
            Return FirstDayOfMonth(MonthValue(value))
        End Function

        Public Shared Function FirstDayOfMonth(ByVal value As Date) As Date
            Return FirstDayOfMonth(MonthValue(value))
        End Function

        Public Shared Function LastDayOfMonth(ByVal monthValue As Integer) As Date
            Return Yesterday(FirstDayOfMonth(monthValue + 1))
        End Function

        Public Shared Function YearPart(ByVal monthValue As Integer) As Integer
            Return (monthValue - 1) \ 12
        End Function

        Public Shared Function MonthPart(ByVal monthValue As Integer) As Integer
            Return (monthValue - 1) Mod 12 + 1
        End Function

        Public Shared Function DateSeries(ByVal seed As Date, ByVal delta As Integer) As String()
            Dim s(System.Math.Abs(delta)) As String
            Dim isMonthEnd As Boolean = (Microsoft.VisualBasic.Day(seed) <> 1)
            For i As Integer = 0 To s.Length - 1
                If isMonthEnd Then
                    s(i) = DateString(Yesterday(DateAdd(DateInterval.Month, i * System.Math.Sign(delta), Tomorrow(seed))))
                Else
                    s(i) = DateString(DateAdd(DateInterval.Month, i * System.Math.Sign(delta), seed))
                End If
            Next
            Return s
        End Function

        ''' <summary>
        ''' 获取指定日期的上一日。
        ''' </summary>
        ''' <param name="value">指定的日期。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Yesterday(ByVal value As Date) As Date
            Return DateAdd(DateInterval.Day, -1, value)
        End Function

        Public Shared Function Yesterday(ByVal value As String) As Date
            Return Yesterday(CDate(value))
        End Function

        ''' <summary>
        ''' 获取指定日期的下一日。
        ''' </summary>
        ''' <param name="value">指定的日期。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Tomorrow(ByVal value As Date) As Date
            Return DateAdd(DateInterval.Day, 1, value)
        End Function

        Public Shared Function Tomorrow(ByVal value As String) As Date
            Return Tomorrow(CDate(value))
        End Function

        Public Shared Function Format(ByVal dateValue As String, ByVal formatString As String) As String
            Return Format(CDate(dateValue), formatString)
        End Function

        Public Shared Function Format(ByVal monthValue As Integer, ByVal formatString As String) As String
            Return Format(FirstDayOfMonth(monthValue), formatString)
        End Function

        ''' <summary>
        ''' 将一个日期类型的数据格式化为符合预定义格式的字符串
        ''' </summary>
        ''' <param name="dateValue">待转换的日期类型数据</param>
        ''' <param name="formatString">预定义字符串格式代码。
        ''' 目前支持的格式有YYYYMMDD、YYYYMM、YYMM、YYYY-MM-DD、YYYYMMDDHHMMSS</param>
        ''' <returns></returns>
        ''' <remarks>若输入的格式代码为目前不支持的格式，则按系统默认方式转换为字符串返回</remarks>
        Public Shared Function Format(ByVal dateValue As Date, ByVal formatString As String) As String
            Select Case formatString.ToUpper()
                Case "YYYYMMDD"
                    Return Microsoft.VisualBasic.Format(dateValue, "yyyyMMdd")
                Case "YYYYMM"
                    Return Microsoft.VisualBasic.Format(dateValue, "yyyyMM")
                Case "YYMM"
                    Return Microsoft.VisualBasic.Format(dateValue, "yyMM")
                Case "YYYY-MM-DD"
                    Return Microsoft.VisualBasic.Format(dateValue, "yyyy-MM-dd")
                Case "YYYYMMDDHHMMSS"
                    Return Microsoft.VisualBasic.Format(dateValue, "yyyyMMddHHmmss")
                Case Else
                    Return CStr(dateValue)
            End Select
        End Function

        Private Shared Function getMonthName(ByVal month As Integer, Optional ByVal isAbbr As Boolean = True) As String
            Dim info As DateTimeFormatInfo = New DateTimeFormatInfo
            If isAbbr Then
                Return info.AbbreviatedMonthNames(month - 1)
            Else
                Return info.MonthNames(month - 1)
            End If
        End Function

        Private Shared Function DateString(ByVal value As Date) As String
            If Language = LanguageCategory.CHINESE Then
                Return Year(value) & "年" & Month(value) & "月" & Day(value) & "日"
            Else
                Return getMonthName(Month(value), False) + " " & Day(value) & ", " & Year(value)
            End If
        End Function

        Public Shared Function DateString(ByVal word As String) As String
            If Language = LanguageCategory.ENGLISH Then
                If word.Chars(word.Length - 1) = "月" And IsDate(word) Then
                    Dim index As Integer = word.IndexOf("年")
                    If index = -1 Then
                        Return getMonthName(CInt(word.Substring(0, word.Length - 1)))
                    Else
                        Return getMonthName(CInt(word.Substring(index + 1, word.Length - index - 2)), False) + " " + word.Substring(0, index)
                    End If
                Else
                    Return DateString(CDate(word))
                End If
            End If
            Return word
        End Function
    End Class
End Namespace
