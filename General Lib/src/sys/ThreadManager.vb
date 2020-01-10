Imports System.Threading.Tasks

Namespace sys
    Public Class ThreadManager

        ''' <summary>
        ''' 以异步线程的方式在后台执行一个方法。
        ''' </summary>
        ''' <param name="action">待执行的方法。</param>
        ''' <remarks></remarks>
        Public Sub ExecuteAction(ByRef action As System.Action)
            Dim mTask As Task = Task.Factory.StartNew(action)
            mTask.Wait()
        End Sub

        ''' <summary>
        ''' 以异步多线程的方式在后台执行一组方法。
        ''' </summary>
        ''' <param name="action">待执行的一组方法。</param>
        ''' <remarks>只有当该组方法对应的多个线程全部执行完毕或者产生异常时，本调用方才返回。</remarks>
        Public Sub ExecuteAction(ByRef action() As System.Action)
            Dim maxIndex As Integer = action.Length - 1
            Dim mTask(maxIndex) As Task
            For i As Integer = 0 To maxIndex
                mTask(i) = Task.Factory.StartNew(action(i))
            Next
            Task.WaitAll(mTask)
            Array.Clear(mTask, 0, mTask.Length)
        End Sub
    End Class
End Namespace
