Namespace finance
    Public Class BrownMotionParameter
        Private m_InitPrice As Single = Nothing
        ''' <summary>
        ''' 资产价格的初始值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property InitPrice As Single
            Get
                Return m_InitPrice
            End Get
            Set(value As Single)
                m_InitPrice = value
            End Set
        End Property

        Private m_Shift As Single = Nothing
        ''' <summary>
        ''' 风险因子变动1个单位的绝对值。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>按微分方法计算Delta、Gamma等风险参数所需的自变量的变动量，即微分的分母。</remarks>
        Public Property Shift As Single
            Get
                Return m_Shift
            End Get
            Set(value As Single)
                m_Shift = value
            End Set
        End Property

        Private m_IncludeShift As Boolean = Nothing
        ''' <summary>
        ''' 布朗运动参数是否包括风险因子偏移1个单位的波动率参数。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IncludeShift As Boolean
            Get
                Return m_IncludeShift
            End Get
            Set(value As Boolean)
                m_IncludeShift = value
            End Set
        End Property

        Private m_Mu() As Double = Nothing
        ''' <summary>
        ''' 布朗运动参数中对数正态分布的均值数组。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend ReadOnly Property Mu As Double()
            Get
                Return m_Mu
            End Get
        End Property

        Private m_Sigma() As Double = Nothing
        Friend ReadOnly Property Sigma As Double()
            Get
                Return m_Sigma
            End Get
        End Property

        ''' <summary>
        ''' 设置资产价格变化的对数正态分布的统计量参数。
        ''' </summary>
        ''' <param name="Mu">对数正态分布的均值数组。</param>
        ''' <param name="Sigma">对数正态分布的方差数组。</param>
        ''' <remarks></remarks>
        Public Sub SetStatisticalMeasures(ByRef Mu() As Double, ByRef Sigma() As Double)
            CopyArrayFully(Mu, m_Mu)
            CopyArrayFully(Sigma, m_Sigma)
        End Sub

        Private m_SigmaUp() As Double = Nothing
        Friend ReadOnly Property SigmaUp As Double()
            Get
                Return m_SigmaUp
            End Get
        End Property

        Private m_SigmaDown() As Double = Nothing
        Friend ReadOnly Property SigmaDown As Double()
            Get
                Return m_SigmaDown
            End Get
        End Property

        ''' <summary>
        ''' 设置风险因子变动1个单位后的波动率参数。
        ''' </summary>
        ''' <param name="SigmaUp">风险因子向上变动1个单位后的波动率参数。</param>
        ''' <param name="SigmaDown">风险因子向下变动1个单位后的波动率参数。</param>
        ''' <remarks></remarks>
        Public Sub SetShiftMeasures(ByRef SigmaUp() As Double, ByRef SigmaDown() As Double)
            CopyArrayFully(SigmaUp, m_SigmaUp)
            CopyArrayFully(SigmaDown, m_SigmaDown)
        End Sub

        Private Sub CopyArrayFully(ByRef source() As Double, ByRef dest() As Double)
            If source Is Nothing Then
                If dest IsNot Nothing Then
                    Array.Clear(dest, 0, dest.Length)
                    dest = Nothing
                End If
                Return
            Else
                If dest Is Nothing Then
                    ReDim dest(source.Length - 1)
                ElseIf dest.Length <> source.Length Then
                    Array.Clear(dest, 0, dest.Length)
                    ReDim dest(source.Length - 1)
                End If
                Array.Copy(source, dest, source.Length)
                Return
            End If
        End Sub

        Private m_BeginDate As Date = Nothing
        ''' <summary>
        ''' 布朗运动预测资产价格的第一个日期。
        ''' </summary>
        ''' <value>设置布朗运动预测资产价格的第一个日期。</value>
        ''' <returns>返回布朗运动预测资产价格的第一个日期。</returns>
        ''' <remarks></remarks>
        Public Property BeginDate As Date
            Get
                Return m_BeginDate
            End Get
            Set(value As Date)
                m_BeginDate = value
            End Set
        End Property

        ''' <summary>
        ''' 布朗运动的预测时间总长度。
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>时间长度的单位是天。</remarks>
        Friend ReadOnly Property PredictedLength As Integer
            Get
                If m_Sigma Is Nothing Then
                    Return 0
                Else
                    Return m_Sigma.Length
                End If
            End Get
        End Property

        Private Sub Clear()
            m_InitPrice = Nothing
            m_Shift = Nothing
            m_IncludeShift = Nothing
            m_BeginDate = Nothing
            If m_Mu IsNot Nothing Then
                Array.Clear(m_Mu, 0, m_Mu.Length)
                m_Mu = Nothing
            End If
            If m_Sigma IsNot Nothing Then
                Array.Clear(m_Sigma, 0, m_Sigma.Length)
                m_Sigma = Nothing
            End If
            If m_SigmaUp IsNot Nothing Then
                Array.Clear(m_SigmaUp, 0, m_SigmaUp.Length)
                m_SigmaUp = Nothing
            End If
            If m_SigmaDown IsNot Nothing Then
                Array.Clear(m_SigmaDown, 0, m_SigmaDown.Length)
                m_SigmaDown = Nothing
            End If
        End Sub

        Protected Overrides Sub Finalize()
            Clear()
            MyBase.Finalize()
        End Sub
    End Class
End Namespace
