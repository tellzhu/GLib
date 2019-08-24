Imports System.Globalization
Imports dotNet.i18n.Languager

Namespace time
    Public Class DateExpert
        ''' <summary>
        ''' ��YYYYMMDD��ʽ���ַ���ת��Ϊ�������͡�
        ''' </summary>
        ''' <param name="value">��YYYYMMDD��ʽ���������ַ�����</param>
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
        ''' ��ȡָ�����ڵ���һ�ա�
        ''' </summary>
        ''' <param name="value">ָ�������ڡ�</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Yesterday(ByVal value As Date) As Date
            Return DateAdd(DateInterval.Day, -1, value)
        End Function

        Public Shared Function Yesterday(ByVal value As String) As Date
            Return Yesterday(CDate(value))
        End Function

        ''' <summary>
        ''' ��ȡָ�����ڵ���һ�ա�
        ''' </summary>
        ''' <param name="value">ָ�������ڡ�</param>
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
        ''' ��һ���������͵����ݸ�ʽ��Ϊ����Ԥ�����ʽ���ַ���
        ''' </summary>
        ''' <param name="dateValue">��ת����������������</param>
        ''' <param name="formatString">Ԥ�����ַ�����ʽ���롣
        ''' Ŀǰ֧�ֵĸ�ʽ��YYYYMMDD��YYYYMM��YYMM��YYYY-MM-DD��YYYYMMDDHHMMSS</param>
        ''' <returns></returns>
        ''' <remarks>������ĸ�ʽ����ΪĿǰ��֧�ֵĸ�ʽ����ϵͳĬ�Ϸ�ʽת��Ϊ�ַ�������</remarks>
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
                Return Year(value) & "��" & Month(value) & "��" & Day(value) & "��"
            Else
                Return getMonthName(Month(value), False) + " " & Day(value) & ", " & Year(value)
            End If
        End Function

        Public Shared Function DateString(ByVal word As String) As String
            If Language = LanguageCategory.ENGLISH Then
                If word.Chars(word.Length - 1) = "��" And IsDate(word) Then
                    Dim index As Integer = word.IndexOf("��")
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
