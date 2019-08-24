Imports System.Math
Imports dotNet.db
Imports dotNet.math
Imports dotNet.time.DateExpert

Namespace finance
    Public Class BrownMotion
        Private m_randomSeed As Random = Nothing
        Private m_cdf As NormalDistribution = Nothing
        Private m_random() As Double = Nothing

        Private Sub UnInit()
            m_randomSeed = Nothing
            m_cdf = Nothing
            If m_random IsNot Nothing Then
                Array.Clear(m_random, 0, m_random.Length)
                m_random = Nothing
            End If
            m_RoundDigits = Nothing
            m_DBTable = Nothing
            m_TargetId = Nothing
            m_MaxPathIndex = Nothing
        End Sub

        Private Sub InitRandomSeed(TotalNumber As Integer)
            ReDim m_random(TotalNumber - 1)
            m_cdf = New NormalDistribution
            m_randomSeed = New Random
        End Sub

        Private Sub FillRandomSeed()
            Dim maxIndex As Integer = m_random.Length - 1
            For i As Integer = 0 To maxIndex
                m_random(i) = m_cdf.NormSInv(m_randomSeed.NextDouble)
            Next
            maxIndex = Nothing
        End Sub

        ''' <summary>
        ''' 根据给定的布朗运动参数，应用蒙特卡洛方法计算资产价格的预测数据。
        ''' </summary>
        ''' <param name="BMParameter">应用到蒙特卡洛模拟中的布朗运动参数。</param>
        ''' <remarks></remarks>
        Public Sub Calculate(ByRef BMParameter As BrownMotionParameter)
            Dim m_BrownMotionParameter As BrownMotionParameter = BMParameter
            Dim firstDate As Date = m_BrownMotionParameter.BeginDate
            Dim s As String = m_TargetId & "," + CStr(firstDate) + ","
            Dim spot As Single = m_BrownMotionParameter.InitPrice
            Dim spotUp As Single = spot + m_BrownMotionParameter.Shift
            Dim spotDown As Single = spot - m_BrownMotionParameter.Shift

            Dim priceOdd(m_MaxPathIndex) As Double
            Dim priceUpOdd(m_MaxPathIndex) As Double
            Dim priceDownOdd(m_MaxPathIndex) As Double
            Dim priceEven(m_MaxPathIndex) As Double
            Dim priceUpEven(m_MaxPathIndex) As Double
            Dim priceDownEven(m_MaxPathIndex) As Double

            Dim m_Mu As Double() = m_BrownMotionParameter.Mu
            Dim m_Sigma As Double() = m_BrownMotionParameter.Sigma
            Dim m_SigmaUp As Double() = m_BrownMotionParameter.SigmaUp
            Dim m_SigmaDown As Double() = m_BrownMotionParameter.SigmaDown

            Dim const1 As Double = 1.0 / 730.0
            Dim const2 As Double = 1.0 / Sqrt(365.0)
            Dim coef1 As Double = m_Mu(0) - const1 * Pow(m_Sigma(0), 2)
            Dim coef2 As Double = const2 * m_Sigma(0)
            Dim coef1Up As Double = Nothing
            Dim coef2Up As Double = Nothing
            Dim coef1Down As Double = Nothing
            Dim coef2Down As Double = Nothing
            If m_BrownMotionParameter.IncludeShift Then
                coef1Up = m_Mu(0) - const1 * Pow(m_SigmaUp(0), 2)
                coef2Up = const2 * m_SigmaUp(0)
                coef1Down = m_Mu(0) - const1 * Pow(m_SigmaDown(0), 2)
                coef2Down = const2 * m_SigmaDown(0)
            Else
                coef1Up = m_Mu(0) - const1 * Pow(m_Sigma(0), 2)
                coef2Up = const2 * m_Sigma(0)
                coef1Down = m_Mu(0) - const1 * Pow(m_Sigma(0), 2)
                coef2Down = const2 * m_Sigma(0)
            End If

            Dim sLoader As StringLoader = New StringLoader
            sLoader.DataTableName = m_DBTable

            FillRandomSeed()
            For k As Integer = 0 To m_MaxPathIndex
                priceOdd(k) = spot * Exp(coef1 + coef2 * m_random(3 * k))
                priceUpOdd(k) = spotUp * Exp(coef1Up + coef2Up * m_random(3 * k + 1))
                priceDownOdd(k) = spotDown * Exp(coef1Down + coef2Down * m_random(3 * k + 2))
                sLoader.Append(s & k + 1 & "," & Round(priceOdd(k), m_RoundDigits) & _
                               "," & Round(priceUpOdd(k), m_RoundDigits) _
                               & "," & Round(priceDownOdd(k), m_RoundDigits))
            Next

            Dim m_PredictedLength As Integer = m_BrownMotionParameter.PredictedLength
            For i As Integer = 2 To m_PredictedLength
                s = m_TargetId & "," + CStr(DateAdd(DateInterval.Day, i - 1, firstDate)) + ","
                FillRandomSeed()

                coef1 = m_Mu(i - 1) - const1 * Pow(m_Sigma(i - 1), 2)
                coef2 = const2 * m_Sigma(i - 1)
                If m_BrownMotionParameter.IncludeShift Then
                    coef1Up = m_Mu(i - 1) - const1 * Pow(m_SigmaUp(i - 1), 2)
                    coef2Up = const2 * m_SigmaUp(i - 1)
                    coef1Down = m_Mu(i - 1) - const1 * Pow(m_SigmaDown(i - 1), 2)
                    coef2Down = const2 * m_SigmaDown(i - 1)
                Else
                    coef1Up = m_Mu(i - 1) - const1 * Pow(m_Sigma(i - 1), 2)
                    coef2Up = const2 * m_Sigma(i - 1)
                    coef1Down = m_Mu(i - 1) - const1 * Pow(m_Sigma(i - 1), 2)
                    coef2Down = const2 * m_Sigma(i - 1)
                End If

                If i Mod 2 = 0 Then
                    For k As Integer = 0 To m_MaxPathIndex
                        priceEven(k) = priceOdd(k) * Exp(coef1 + coef2 * m_random(3 * k))
                        priceUpEven(k) = priceUpOdd(k) * Exp(coef1Up + coef2Up * m_random(3 * k + 1))
                        priceDownEven(k) = priceDownOdd(k) * Exp(coef1Down + coef2Down * m_random(3 * k + 2))
                        sLoader.Append(s & k + 1 & "," & Round(priceEven(k), m_RoundDigits) & _
                                       "," & Round(priceUpEven(k), m_RoundDigits) _
                                       & "," & Round(priceDownEven(k), m_RoundDigits))
                    Next
                Else
                    For k As Integer = 0 To m_MaxPathIndex
                        priceOdd(k) = priceEven(k) * Exp(coef1 + coef2 * m_random(3 * k))
                        priceUpOdd(k) = priceUpEven(k) * Exp(coef1Up + coef2Up * m_random(3 * k + 1))
                        priceDownOdd(k) = priceDownEven(k) * Exp(coef1Down + coef2Down * m_random(3 * k + 2))
                        sLoader.Append(s & k + 1 & "," & Round(priceOdd(k), m_RoundDigits) & _
                                       "," & Round(priceUpOdd(k), m_RoundDigits) _
                                       & "," & Round(priceDownOdd(k), m_RoundDigits))
                    Next
                End If
            Next

            firstDate = Nothing
            s = Nothing
            spot = Nothing
            spotUp = Nothing
            spotDown = Nothing
            m_PredictedLength = Nothing

            Array.Clear(priceOdd, 0, priceOdd.Length)
            priceOdd = Nothing
            Array.Clear(priceUpOdd, 0, priceUpOdd.Length)
            priceUpOdd = Nothing
            Array.Clear(priceDownOdd, 0, priceDownOdd.Length)
            priceDownOdd = Nothing
            Array.Clear(priceEven, 0, priceEven.Length)
            priceEven = Nothing
            Array.Clear(priceUpEven, 0, priceUpEven.Length)
            priceUpEven = Nothing
            Array.Clear(priceDownEven, 0, priceDownEven.Length)
            priceDownEven = Nothing

            Array.Clear(m_Mu, 0, m_Mu.Length)
            m_Mu = Nothing
            Array.Clear(m_Sigma, 0, m_Sigma.Length)
            m_Sigma = Nothing
            Array.Clear(m_SigmaUp, 0, m_SigmaUp.Length)
            m_SigmaUp = Nothing
            Array.Clear(m_SigmaDown, 0, m_SigmaDown.Length)
            m_SigmaDown = Nothing

            const1 = Nothing
            const2 = Nothing
            coef1 = Nothing
            coef2 = Nothing
            coef1Up = Nothing
            coef2Up = Nothing
            coef1Down = Nothing
            coef2Down = Nothing
            m_BrownMotionParameter = Nothing

            sLoader.Load(m_DBTable)
            sLoader = Nothing
        End Sub

        Private m_TargetId As Integer = Nothing
        Private m_DBTable As String = Nothing
        Private m_RoundDigits As Integer = Nothing
        ''' <summary>
        ''' 设置蒙特卡洛模拟预测数据的存储参数。
        ''' </summary>
        ''' <param name="TableName">存储预测数据的数据库表名称。</param>
        ''' <param name="TargetId">预测标的资产的编码ID。</param>
        ''' <param name="RoundDigits">标的资产价格的小数位数。</param>
        ''' <remarks>数据库表的定义定为：资产标的编码、预测日期、路径、预测价格、风险因子向上波动后的预测价格和风险因子向下波动后的预测价格。</remarks>
        Public Sub SetDBParameter(TableName As String, TargetId As Integer, _
                                  Optional RoundDigits As Integer = 2)
            m_DBTable = TableName
            m_TargetId = TargetId
            m_RoundDigits = RoundDigits
        End Sub

        Protected Overrides Sub Finalize()
            UnInit()
            MyBase.Finalize()
        End Sub

        Private m_MaxPathIndex As Integer = Nothing
        ''' <summary>
        ''' 初始化一个布朗运动实例，并设置蒙特卡洛模拟的路径数量。
        ''' </summary>
        ''' <param name="MaxPath">路径数量。</param>
        ''' <remarks></remarks>
        Public Sub New(MaxPath As Integer)
            m_MaxPathIndex = MaxPath - 1
            InitRandomSeed(3 * MaxPath)
        End Sub

    End Class
End Namespace

