Namespace time
    ''' <summary>
    ''' 用于统计系统运行时间的计时器。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Chronometer
        Private Shared m_FirstTimestamp As Date = Nothing
        Private Shared m_LastTimestamp As Date = Nothing

        ''' <summary>
        ''' 启动计时器。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Start()
            m_FirstTimestamp = Now
            m_LastTimestamp = m_FirstTimestamp
        End Sub

        ''' <summary>
        ''' 获取当前时间与计时器启动时间之间的时间间隔值。
        ''' </summary>
        ''' <param name="Unit">度量时间间隔的单位。</param>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property Interval(Optional Unit As DateInterval = DateInterval.Second) As Integer
            Get
                Dim m_CurrentTimestamp As Date = Now
                Dim val As Integer = CInt(DateDiff(Unit, m_FirstTimestamp, m_CurrentTimestamp))
                m_LastTimestamp = m_CurrentTimestamp
                Return val
            End Get
        End Property

        ''' <summary>
        ''' 获取当前时间与计时器上次计数时间之间的时间间隔值。
        ''' </summary>
        ''' <param name="Unit">度量时间间隔的单位。</param>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property LastInterval(Optional Unit As DateInterval = DateInterval.Second) As Integer
            Get
                Dim m_CurrentTimestamp As Date = Now
                Dim val As Integer = CInt(DateDiff(Unit, m_LastTimestamp, m_CurrentTimestamp))
                m_LastTimestamp = m_CurrentTimestamp
                Return val
            End Get
        End Property

    End Class
End Namespace
