Imports System.Math

Namespace math
    Public Class NormalDistribution

        Public Function NormSInv(Probability As Double) As Double
            If Probability > 1 Or Probability < 0 Then
                Return 0
            End If
            If Probability = 0.5 Then
                Return 0
            ElseIf Probability < 0.5 Then
                Return -NormSInv(1 - Probability)
            Else
                Dim count As Integer = 1
                Dim seed As Double = 0.675
                Dim cdf As Double
                Dim lowLimit As Double = 0
                Dim highLimit As Double = 6
                While count <= 100
                    cdf = NormSDist(seed)
                    If Abs(cdf - Probability) <= NormSMinError Then
                        Exit While
                    Else
                        If cdf > Probability Then
                            highLimit = seed
                        Else
                            lowLimit = seed
                        End If
                        seed = 0.5 * (highLimit + lowLimit)
                    End If
                    count += 1
                End While
                Return seed
            End If
        End Function

        Private Const NormSMinError As Double = 0.000001
        Private Const SqrtInv2Pi As Double = (2 * PI) ^ (-0.5)
        Private Const gamma As Double = 0.2316419
        Private Const a1 As Double = 0.31938153
        Private Const a2 As Double = -0.356563782
        Private Const a3 As Double = 1.781477937
        Private Const a4 As Double = -1.821255978
        Private Const a5 As Double = 1.330274429

        Private Function NormSDist(z As Double) As Double
            If z = 0 Then
                Return 0.5
            ElseIf z < 0 Then
                Return 1 - NormSDist(-z)
            Else
                Dim k As Double = 1 / (1 + gamma * z)
                Return 1 - SqrtInv2Pi * Exp(-0.5 * (z ^ 2)) * (a1 * k + a2 * (k ^ 2) + a3 * (k ^ 3) + a4 * (k ^ 4) + a5 * (k ^ 5))
            End If
        End Function
    End Class
End Namespace
