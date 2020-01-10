Imports System.Math

Namespace finance
    Public Class DateProcessor
        Private Shared Function MonthIndex(ByVal DateVal As Date) As Integer
            Return DateVal.Year * 12 + DateVal.Month
        End Function

        Private Shared Function IndexOfPeriod(ByVal DateVal As Date, ByVal PeriodLength As Integer) As Integer
            Return (MonthIndex(DateVal) - 1) Mod PeriodLength
        End Function

        Private Shared Function FirstMonthIndexOfPeriod(ByVal DateVal As Date, ByVal PeriodLength As Integer) As Integer
            Return MonthIndex(DateVal) - IndexOfPeriod(DateVal, PeriodLength)
        End Function

        Private Shared Function DaysInMonth(ByVal MonthIndex As Integer) As Integer
            Return Date.DaysInMonth((MonthIndex - 1) \ 12, (MonthIndex - 1) Mod 12 + 1)
        End Function

        Private Shared Function DateInMonthIndex(ByVal MonthIndex As Integer, ByVal Day As Integer) As Date
            Return New Date((MonthIndex - 1) \ 12, (MonthIndex - 1) Mod 12 + 1, Day)
        End Function

        Friend Shared Function CouponPreviousPaymentDate(ByVal PaymentDate As Date, ByVal MaturityDate As Date, _
                                                           ByVal Frequency As Integer) As Date
            Dim mIndex As Integer = MonthIndex(PaymentDate) - 12 \ Frequency
            Return DateInMonthIndex(mIndex, Min(DaysInMonth(mIndex), MaturityDate.Day))
        End Function

        Friend Shared Function CouponNumber(ByVal NextPaymentDate As Date, ByVal MaturityDate As Date, _
                                                    ByVal Frequency As Integer) As Integer
            Return (MonthIndex(MaturityDate) - MonthIndex(NextPaymentDate)) \ (12 \ Frequency) + 1
        End Function

        Friend Enum InterestBasis As Integer
            US30_360 = 0
            ACTUAL = 1
            A360 = 2
            A365 = 3
            EU30_360 = 4
        End Enum

        Private Shared Function DaysInUS30_360(ByVal Date1 As Date, ByVal Date2 As Date) As Integer
            Dim D1 As Integer = Date1.Day, D2 As Integer = Date2.Day
            If (Date1.Month = 2 And D1 = Date.DaysInMonth(Date1.Year, 2) _
                Or (D1 = 30 Or D1 = 31)) _
                And Date2.Month = 2 And D2 = Date.DaysInMonth(Date2.Year, 2) Then
                D2 = 30
            End If
            If Date1.Month = 2 And D1 = Date.DaysInMonth(Date1.Year, 2) Then
                D1 = 30
            End If
            If (D1 = 30 Or D1 = 31) And D2 = 31 Then
                D2 = 30
            End If
            If D1 = 31 Then
                D1 = 30
                D2 -= 1
            End If
            Return 360 * (Date2.Year - Date1.Year) + 30 * (Date2.Month - Date1.Month) + (D2 - D1)
        End Function

        Private Shared Function DaysInEU30_360(ByVal Date1 As Date, ByVal Date2 As Date) As Integer
            Dim D1 As Integer = Date1.Day, D2 As Integer = Date2.Day
            If D2 = 31 Then
                D2 = 30
            End If
            If D1 = 31 Then
                D1 = 30
            End If
            Return 360 * (Date2.Year - Date1.Year) + 30 * (Date2.Month - Date1.Month) + (D2 - D1)
        End Function

        Friend Shared Function DaysCount(ByVal PrevDate As Date, ByVal NextDate As Date, _
                                                          ByVal Basis As InterestBasis) As Integer
            Select Case Basis
                Case InterestBasis.ACTUAL, InterestBasis.A360, InterestBasis.A365
                    Return CInt(DateDiff(DateInterval.Day, PrevDate, NextDate))
                Case InterestBasis.US30_360
                    Return DaysInUS30_360(PrevDate, NextDate)
                Case InterestBasis.EU30_360
                    Return DaysInEU30_360(PrevDate, NextDate)
                Case Else
                    Return 0
            End Select
        End Function

        Friend Shared Function CouponDays(ByVal PrevPaymentDate As Date, ByVal NextPaymentDate As Date, _
                                            ByVal Frequency As Integer, ByVal Basis As InterestBasis) As Decimal
            Select Case Basis
                Case InterestBasis.ACTUAL
                    Return DateDiff(DateInterval.Day, PrevPaymentDate, NextPaymentDate)
                Case InterestBasis.US30_360, InterestBasis.A360, InterestBasis.EU30_360
                    Return 360 \ Frequency
                Case InterestBasis.A365
                    Return CDec(365 / Frequency)
                Case Else
                    Return 0
            End Select
        End Function

        Private Shared Function YearFracInACTUAL(ByVal Date1 As Date, ByVal Date2 As Date) As Double
            If Date1.Date >= Date2.Date Then
                Return 0
            End If
            If Date1.Year = Date2.Year Or (Date1.Year + 1 = Date2.Year _
                                           And (Date1.Month > Date2.Month Or Date1.Month = Date2.Month And Date1.Day >= Date2.Day)) Then
                If Date1.Year = Date2.Year And Date.IsLeapYear(Date1.Year) Or _
                Date.IsLeapYear(Date1.Year) And Date1 < New Date(Date1.Year, 3, 1) Or _
                Date.IsLeapYear(Date2.Year) And Date2 > New Date(Date2.Year, 2, 28) Then
                    Return DateDiff(DateInterval.Day, Date1, Date2) / 366.0
                Else
                    Return DateDiff(DateInterval.Day, Date1, Date2) / 365.0
                End If
            Else
                Dim frac As Double = 0
                Dim yr As Integer = Date1.Year
                While yr <= Date2.Year
                    If Date.IsLeapYear(yr) Then
                        frac += 366
                    Else
                        frac += 365
                    End If
                    yr += 1
                End While
                frac /= (yr - Date1.Year)
                Return DateDiff(DateInterval.Day, Date1, Date2) / frac
            End If
        End Function

        Public Shared Function YearFrac(ByVal StartDate As Date, ByVal EndDate As Date, ByVal Basis As Integer) As Double
            Select Case Basis
                Case InterestBasis.A360
                    Return DateDiff(DateInterval.Day, StartDate, EndDate) / 360.0
                Case InterestBasis.A365
                    Return DateDiff(DateInterval.Day, StartDate, EndDate) / 365.0
                Case InterestBasis.ACTUAL
                    Return YearFracInACTUAL(StartDate, EndDate)
                Case InterestBasis.US30_360
                    Return DaysInUS30_360(StartDate, EndDate) / 360.0
                Case InterestBasis.EU30_360
                    Return DaysInEU30_360(StartDate, EndDate) / 360.0
            End Select
            Return -1
        End Function

        Friend Shared Function CouponNextPaymentDate(ByVal SettlementDate As Date, ByVal MaturityDate As Date, _
                                                        ByVal Frequency As Integer) As Date
            Dim periodLength As Integer = 12 \ Frequency
            Dim monthIndex As Integer = FirstMonthIndexOfPeriod(SettlementDate, periodLength) _
                                 + IndexOfPeriod(MaturityDate, periodLength)
            Dim date1 As Date = DateInMonthIndex(monthIndex, Min(DaysInMonth(monthIndex), MaturityDate.Day))
            If SettlementDate < date1 Then
                Return date1
            Else
                monthIndex += periodLength
                Return DateInMonthIndex(monthIndex, Min(DaysInMonth(monthIndex), MaturityDate.Day))
            End If
        End Function
    End Class

End Namespace
