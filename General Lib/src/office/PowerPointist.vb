Imports Microsoft.Office.Interop.PowerPoint

Namespace office
    ''' <summary>
    ''' 通过COM互操作方式使用PowerPoint API的抽象专家类。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PowerPointist
        Private Shared currentPresentation As Presentation = Nothing
        Private Shared currentApp As Application = Nothing

        ''' <summary>
        ''' 打开指定的PowerPoint文件。
        ''' </summary>
        ''' <param name="fileName">PowerPoint文件名称，不包含文件后缀名。</param>
        ''' <param name="extension">PowerPoint文件后缀名。</param>
        ''' <remarks></remarks>
        Public Shared Sub OpenPresentation(ByVal fileName As String, Optional ByVal extension As String = "ppt")
            CloseLastPresentation()
            '     currentPresentation = currentApp.Presentations.Open( _
            '    FileName:=fileName + "." + extension, WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse)
        End Sub

        ''' <summary>
        ''' 在当前PowerPoint文件的指定位置之后增加一张新的幻灯片。
        ''' </summary>
        ''' <param name="Index">指定的幻灯片位置。</param>
        ''' <remarks></remarks>
        Public Shared Sub AddSlide(ByVal Index As Integer)
            Dim pptLayout As CustomLayout = currentPresentation.Slides(Index).CustomLayout
            currentPresentation.Slides.AddSlide(Index, pptLayout)
            pptLayout = Nothing
        End Sub

        ''' <summary>
        ''' 将剪贴板中的图表粘贴在当前PowerPoint文件的指定幻灯片中。
        ''' </summary>
        ''' <param name="Index">指定的幻灯片位置。</param>
        ''' <remarks></remarks>
        Public Shared Sub PasteChart(ByVal Index As Integer)
            currentPresentation.Slides(Index).Shapes.PasteSpecial(DataType:=PpPasteDataType.ppPasteBitmap)
        End Sub

        Private Shared Sub CloseLastPresentation()
            If currentPresentation IsNot Nothing Then
                currentPresentation.Save()
                currentPresentation.Close()
                currentPresentation = Nothing
            End If
        End Sub

        ''' <summary>
        ''' 保存当前的PowerPoint文件。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Save()
            If currentPresentation IsNot Nothing Then
                currentPresentation.Save()
            End If
        End Sub

        ''' <summary>
        ''' 在后台启动一个PowerPoint应用程序进程。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub StartApplication()
            currentApp = New Application()
        End Sub

        ''' <summary>
        ''' 退出后台启动的PowerPoint应用程序进程。
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub ExitApplication()
            CloseLastPresentation()
            With currentApp
                .Quit()
            End With
            currentApp = Nothing
        End Sub

    End Class
End Namespace

