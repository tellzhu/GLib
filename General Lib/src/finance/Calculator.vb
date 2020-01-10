Imports System.Math
Imports dotNet.finance.DateProcessor

Namespace finance
    Public Class Calculator

        Private Const MinError As Double = 0.000000000001

        Public Shared Function PriceDisc(ByVal SettlementDate As Date, ByVal Maturity As Date, _
                                     ByVal DiscountRate As Decimal, ByVal Redemption As Decimal, ByVal InterestBasis As Integer) As Double
            Return Redemption * (1 - DiscountRate * YearFrac(SettlementDate, Maturity, InterestBasis))
        End Function

        Public Shared Function PriceMat(ByVal SettlementDate As Date, ByVal Maturity As Date, ByVal IssueDate As Date, _
                                    ByVal Rate As Decimal, ByVal Yield As Decimal, ByVal InterestBasis As Integer) As Double
            Dim YIM As Double = YearFrac(IssueDate, Maturity, InterestBasis)
            Dim YIS As Double = YearFrac(IssueDate, SettlementDate, InterestBasis)
            Dim YSM As Double = YearFrac(SettlementDate, Maturity, InterestBasis)
            Return 100 * (1 + YIM * Rate) / (1 + YSM * Yield) - YIS * Rate * 100
        End Function

        Public Shared Function AccrInt(ByVal SettlementDate As Date, ByVal Maturity As Date, ByVal Rate As Decimal, _
                                   ByVal Par As Integer, ByVal Frequency As Integer, ByVal InterestBasis As Integer) As Double
            If Rate < 0 Then
                Return -1
            End If
            If Frequency <> 1 And Frequency <> 2 And Frequency <> 4 And Frequency <> 12 Then
                Return -1
            End If
            If InterestBasis < 0 Or InterestBasis > 4 Then
                Return -1
            End If
            If SettlementDate >= Maturity Then
                Return -1
            End If

            Dim nextPaymentDate As Date = CouponNextPaymentDate(SettlementDate, Maturity, Frequency)
            Dim prevPaymentDate As Date = CouponPreviousPaymentDate(nextPaymentDate, Maturity, Frequency)
            Dim E As Decimal = CouponDays(prevPaymentDate, nextPaymentDate, Frequency, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim A As Integer = DaysCount(prevPaymentDate, SettlementDate, CType(InterestBasis, DateProcessor.InterestBasis))

            Return Par * (Rate / Frequency) * (A / E)
        End Function

        Public Shared Function Price(ByVal SettlementDate As Date, ByVal Maturity As Date, ByVal Rate As Decimal, _
                                  ByVal Yield As Decimal, ByVal Redemption As Decimal, ByVal Frequency As Integer, _
                                  ByVal InterestBasis As Integer) As Double
            If Rate < 0 Or Redemption <= 0 Then
                Return -1
            End If
            If Frequency <> 1 And Frequency <> 2 And Frequency <> 4 And Frequency <> 12 Then
                Return -1
            End If
            If InterestBasis < 0 Or InterestBasis > 4 Then
                Return -1
            End If
            If SettlementDate >= Maturity Then
                Return -1
            End If

            Dim nextPaymentDate As Date = CouponNextPaymentDate(SettlementDate, Maturity, Frequency)
            Dim prevPaymentDate As Date = CouponPreviousPaymentDate(nextPaymentDate, Maturity, Frequency)
            Dim E As Decimal = CouponDays(prevPaymentDate, nextPaymentDate, Frequency, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim DSC As Integer = DaysCount(SettlementDate, nextPaymentDate, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim N As Integer = CouponNumber(nextPaymentDate, Maturity, Frequency)
            Dim A As Integer = DaysCount(prevPaymentDate, SettlementDate, CType(InterestBasis, DateProcessor.InterestBasis))

            Dim Sn As Double
            If Yield = 0 Then
                Sn = Redemption + (N - A / E) * 100 * (Rate / Frequency)
            Else
                Dim SingleProduct As Double = 1 + Yield / Frequency
                Sn = (1 - Pow(SingleProduct, -N)) / (SingleProduct - 1)
                Sn *= (100 * Rate / Frequency) * Pow(SingleProduct, 1 - DSC / E)
                Sn = Sn + Redemption * Pow(SingleProduct, -(N - 1 + DSC / E)) - 100 * (Rate / Frequency) * (A / E)
            End If
            Return Sn
        End Function

        Public Shared Function MDuration(ByVal SettlementDate As Date, ByVal Maturity As Date, ByVal Rate As Decimal, _
ByVal Yield As Double, ByVal Frequency As Integer, ByVal InterestBasis As Integer) As Double
            If Rate < 0 Or Yield < 0 Then
                Return -1
            End If
            If Frequency <> 1 And Frequency <> 2 And Frequency <> 4 And Frequency <> 12 Then
                Return -1
            End If
            If InterestBasis < 0 Or InterestBasis > 4 Then
                Return -1
            End If
            If SettlementDate >= Maturity Then
                Return -1
            End If
            Return Duration(SettlementDate, Maturity, Rate, Yield, Frequency, InterestBasis) / (1 + Yield / Frequency)
        End Function

        Private Shared Function Duration(ByVal SettlementDate As Date, ByVal Maturity As Date, ByVal Rate As Decimal, _
    ByVal Yield As Double, ByVal Frequency As Integer, ByVal InterestBasis As Integer) As Double
            If Rate < 0 Or Yield < 0 Then
                Return -1
            End If
            If Frequency <> 1 And Frequency <> 2 And Frequency <> 4 And Frequency <> 12 Then
                Return -1
            End If
            If InterestBasis < 0 Or InterestBasis > 4 Then
                Return -1
            End If
            If SettlementDate >= Maturity Then
                Return -1
            End If

            If Abs(Yield) < MinError Then
                Yield = 0
            End If

            Dim nextPaymentDate As Date = CouponNextPaymentDate(SettlementDate, Maturity, Frequency)
            Dim prevPaymentDate As Date = CouponPreviousPaymentDate(nextPaymentDate, Maturity, Frequency)
            Dim E As Decimal = CouponDays(prevPaymentDate, nextPaymentDate, Frequency, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim DSC As Integer = DaysCount(SettlementDate, nextPaymentDate, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim N As Integer = CouponNumber(nextPaymentDate, Maturity, Frequency)

            Dim Sn As Double
            Dim d As Double = -1 + DSC / E
            If Yield = 0 Then
                Sn = (N + d + Rate * (N * d + 0.5 * N * (N + 1)) / Frequency) / (Frequency + N * Rate)
            Else
                Dim c As Double = 1 + Yield / Frequency
                Dim b As Double = Rate / Frequency
                Dim a As Double = Pow(c, -(N + d))
                Dim f As Double = (1 - Pow(c, -N)) * Frequency / Yield
                Sn = (b * Pow(c, -d) * f + a) * Frequency
                Sn = ((N + d) * a + b * (d * Pow(c, -d) * f + (1 + Frequency / Yield) * (f - N * Pow(c, -N - 1)))) / Sn
            End If
            Return Sn
        End Function

        Public Shared Function Yield(ByVal SettlementDate As Date, ByVal Maturity As Date, ByVal Rate As Decimal, _
                                    ByVal Price As Decimal, ByVal Redemption As Decimal, ByVal Frequency As Integer, _
                                    ByVal InterestBasis As Integer, ByVal YieldSeed As Double) As Double
            If Rate < 0 Or Price <= 0 Or Redemption <= 0 Then
                Return -1
            End If
            If Frequency <> 1 And Frequency <> 2 And Frequency <> 4 And Frequency <> 12 Then
                Return -1
            End If
            If InterestBasis < 0 Or InterestBasis > 4 Then
                Return -1
            End If
            If SettlementDate >= Maturity Then
                Return -1
            End If

            Dim nextPaymentDate As Date = CouponNextPaymentDate(SettlementDate, Maturity, Frequency)
            Dim prevPaymentDate As Date = CouponPreviousPaymentDate(nextPaymentDate, Maturity, Frequency)
            Dim N As Integer = CouponNumber(nextPaymentDate, Maturity, Frequency)
            Dim E As Decimal = CouponDays(prevPaymentDate, nextPaymentDate, Frequency, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim A As Integer = DaysCount(prevPaymentDate, SettlementDate, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim DSR As Integer = DaysCount(SettlementDate, nextPaymentDate, CType(InterestBasis, DateProcessor.InterestBasis))
            Dim Sn As Double = 0
            Dim SingleRate As Double

            If N <= 1 Then
                SingleRate = Rate / Frequency
                Sn = ((0.01 * Redemption + SingleRate) / (0.01 * Price + (SingleRate * A / E)) _
                                   - 1) * Frequency * E / DSR
            Else
                Dim count As Integer = 0
                Dim lowLimit As Double = Nothing, highLimit As Double = Nothing
                If YieldSeed <= -1 Then
                    YieldSeed = -1 + 0.00000001
                ElseIf YieldSeed > 1 Then
                    YieldSeed = 1
                End If
                Dim IsFixedDouble As Boolean = False
                Dim borderCount As Integer = 0
                While count <= 100
                    If YieldSeed = 0 Then
                        Sn = Redemption + (N - A / E) * 100 * (Rate / Frequency)
                    Else
                        SingleRate = 1 + YieldSeed / Frequency
                        Sn = (1 - Pow(SingleRate, -N)) / (SingleRate - 1)
                        Sn *= (100 * Rate / Frequency) * Pow(SingleRate, 1 - DSR / E)
                        Sn = Sn + Redemption * Pow(SingleRate, -(N - 1 + DSR / E)) - 100 * (Rate / Frequency) * (A / E)
                    End If
                    If Abs(Sn - Price) > MinError Then
                        If Sn > Price Then
                            If borderCount = 0 Then
                                borderCount = 1
                            ElseIf borderCount = 2 Then
                                borderCount = 3
                            End If
                            lowLimit = YieldSeed
                            YieldSeed += 0.1
                            highLimit = YieldSeed
                        Else
                            If borderCount = 0 Then
                                borderCount = 2
                            ElseIf borderCount = 1 Then
                                borderCount = 3
                            End If
                            highLimit = YieldSeed
                            YieldSeed -= 0.1
                            lowLimit = YieldSeed
                        End If
                    Else
                        Sn = YieldSeed
                        IsFixedDouble = True
                        Exit While
                    End If
                    If borderCount = 3 Then
                        Exit While
                    End If
                    count += 1
                End While
                If Not IsFixedDouble Then
                    While count <= 100
                        YieldSeed = 0.5 * (lowLimit + highLimit)
                        If YieldSeed = 0 Then
                            Sn = Redemption + (N - A / E) * 100 * (Rate / Frequency)
                        Else
                            SingleRate = 1 + YieldSeed / Frequency
                            Sn = (1 - Pow(SingleRate, -N)) / (SingleRate - 1)
                            Sn *= (100 * Rate / Frequency) * Pow(SingleRate, 1 - DSR / E)
                            Sn = Sn + Redemption * Pow(SingleRate, -(N - 1 + DSR / E)) - 100 * (Rate / Frequency) * (A / E)
                        End If
                        If Abs(Sn - Price) > MinError Then
                            If Sn > Price Then
                                lowLimit = YieldSeed
                            Else
                                highLimit = YieldSeed
                            End If
                        Else
                            Exit While
                        End If
                        count += 1
                    End While
                    Sn = YieldSeed
                End If
            End If

            Return NormalizeYield(Sn)
        End Function

        Private Shared Function NormalizeYield(ByVal value As Double) As Double
            If value < -0.2 Then
                Return -0.2
            ElseIf value > 10 Then
                Return 10
            Else
                Return value
            End If
        End Function

        Public Shared Function YieldDisc(ByVal SettlementDate As Date, ByVal Maturity As Date, _
                                     ByVal Price As Decimal, ByVal Redemption As Decimal, ByVal InterestBasis As Integer) As Double
            Dim denominator As Double = Price * YearFrac(SettlementDate, Maturity, InterestBasis)
            If denominator = 0 Then
                Return 0
            End If
            Return NormalizeYield((Redemption - Price) / denominator)
        End Function

        Public Shared Function YieldMat(ByVal SettlementDate As Date, ByVal Maturity As Date, ByVal IssueDate As Date, _
                                    ByVal Rate As Decimal, ByVal Price As Decimal, ByVal InterestBasis As Integer) As Double
            Dim YIM As Double = YearFrac(IssueDate, Maturity, InterestBasis)
            Dim YIS As Double = YearFrac(IssueDate, SettlementDate, InterestBasis)
            Dim YSM As Double = YearFrac(SettlementDate, Maturity, InterestBasis)
            If YSM = 0 Then
                Return 0
            End If
            Return NormalizeYield((100 * (1 + YIM * Rate) / (Price + YIS * Rate * 100) - 1) / YSM)
        End Function
    End Class
End Namespace
