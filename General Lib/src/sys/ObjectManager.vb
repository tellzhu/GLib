Imports System.Reflection

Namespace sys
    ''' <summary>
    ''' 对面向对象的程序提供基础服务支持。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ObjectManager
        ''' <summary>
        ''' 创建一个指定类的新实例。
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="AssemblyFileName">包含指定类的程序集名称。</param>
        ''' <param name="FullClassName">指定的类名称。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreateInstance(Of T)(AssemblyFileName As String, FullClassName As String) As T
            Dim asm As Assembly = Assembly.LoadFrom(AssemblyFileName)
            Dim t1 As Type = asm.GetType(FullClassName, True)
            Dim o As Object = asm.CreateInstance(t1.FullName)
            Return CType(o, T)
        End Function

        ''' <summary>
        ''' 从当前程序集中创建一个指定类的新实例。
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="FullClassName">指定的类名称。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreateInstance(Of T)(FullClassName As String) As T
            Dim asm As Assembly = Assembly.GetCallingAssembly()
            Dim t1 As Type = asm.GetType(FullClassName, True)
            Dim o As Object = asm.CreateInstance(t1.FullName)
            Return CType(o, T)
        End Function
    End Class
End Namespace
