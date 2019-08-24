Namespace office
    ''' <summary>
    ''' 自动构建Excel计算公式的工具箱
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FormulaBuilder

        ''' <summary>
        ''' IF判断函数。
        ''' </summary>
        ''' <param name="condition">逻辑判断条件。</param>
        ''' <param name="trueValue">若条件为真时的取值。</param>
        ''' <param name="falseValue">若条件为假时的取值。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IfCondition(ByVal condition As String, ByVal trueValue As String, ByVal falseValue As String) As String
            Return "IF(" + condition + "," + trueValue + "," + falseValue + ")"
        End Function

        ''' <summary>
        ''' 数值相除函数。
        ''' </summary>
        ''' <param name="numerator">以字符串表示的被除数。</param>
        ''' <param name="denominator">以字符串表示的除数。</param>
        ''' <param name="IsTestZeroDenominator" >是否对除数为零的情况进行检测，并在除数为零时将结果设置为空值。</param>
        ''' <returns>若除数为零，则返回空字符串（""）；否则返回相除结果。</returns>
        ''' <remarks></remarks>
        Public Shared Function Divide(ByVal numerator As String, ByVal denominator As String, _
                                      Optional ByVal IsTestZeroDenominator As Boolean = True) As String
            If IsTestZeroDenominator Then
                Return IfCondition(denominator + "=0", """""", "(" + numerator + ")/(" + denominator + ")")
            Else
                Return numerator + "/" + denominator
            End If
        End Function

        Public Shared Function Subtract(ByVal minuend As String, ByVal subtractor As String) As String
            Return "(" + minuend + "-" + subtractor + ")"
        End Function

        ''' <summary>
        ''' 日期取值函数。
        ''' </summary>
        ''' <param name="value">以日期类型表示的待转换数值。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DateVal(ByVal value As Date) As String
            Return "DATE(" + CStr(value.Year) + "," + CStr(value.Month) + "," + CStr(value.Day) + ")"
        End Function

        ''' <summary>
        ''' 数值相加函数。
        ''' </summary>
        ''' <param name="value1">以字符串表示的加数。</param>
        ''' <param name="value2">以字符串表示的被加数。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Add(ByVal value1 As String, ByVal value2 As String) As String
            Return "(" + value1 + "+" + value2 + ")"
        End Function

        Public Shared Function Multiple(ByVal value1 As String, ByVal value2 As String) As String
            Return "(" + value1 + ")*(" + value2 + ")"
        End Function

        Public Shared Function SumProduct(ByVal array1 As String, ByVal array2 As String) As String
            Return BinaryFunction("SUMPRODUCT", array1, array2)
        End Function

        Public Shared Function LogicalAnd(ByVal value1 As String, ByVal value2 As String) As String
            Return BinaryFunction("AND", value1, value2)
        End Function

        Public Shared Function IsNumber(ByVal value As String) As String
            Return UniFunction("ISNUMBER", value)
        End Function

        Public Shared Function Sum(ByVal value As String) As String
            Return UniFunction("SUM", value)
        End Function

        Public Shared Function Average(ByVal value As String) As String
            Return UniFunction("AVERAGE", value)
        End Function

        Public Shared Function Count(ByVal value As String) As String
            Return UniFunction("COUNT", value)
        End Function

        Public Shared Function Maximum(ByVal value As String) As String
            Return UniFunction("MAX", value)
        End Function

        Private Shared Function UniFunction(ByVal funcName As String, ByVal para As String) As String
            Return funcName + "(" + para + ")"
        End Function

        Private Shared Function BinaryFunction(ByVal funcName As String, ByVal para1 As String, ByVal para2 As String) As String
            Return funcName + "(" + para1 + "," + para2 + ")"
        End Function

        Public Shared Function Minimum(ByVal value As String) As String
            Return UniFunction("MIN", value)
        End Function
    End Class
End Namespace
