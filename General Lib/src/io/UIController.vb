Imports System.Windows.Forms

Namespace io
    Public Class UIController
        Private Shared ctrlNameSet As HashSet(Of String) = Nothing

        Private Shared Sub ForbidToolStripMenuItem(ByRef menu As ToolStripMenuItem, Optional ByVal IsEnabled As Boolean = True)
            Dim cnt As Integer = menu.DropDownItems.Count
            Dim subMenu As ToolStripMenuItem = Nothing
            Dim ctrl As ToolStripItem = Nothing
            Dim t As Type = Nothing
            For i As Integer = 1 To cnt
                ctrl = menu.DropDownItems(i - 1)
                t = ctrl.GetType
                If t.Name <> "ToolStripSeparator" Then
                    subMenu = CType(ctrl, ToolStripMenuItem)
                    subMenu.Enabled = IsEnabled And Not ctrlNameSet.Contains(subMenu.Name)
                    If subMenu.HasDropDownItems Then
                        ForbidToolStripMenuItem(subMenu, subMenu.Enabled)
                    End If
                End If
            Next
            ctrl = Nothing
            t = Nothing
            subMenu = Nothing
            cnt = Nothing
        End Sub

        Private Shared Sub ForbidMenuStrip(ByRef menu As MenuStrip, Optional ByVal IsEnabled As Boolean = True)
            Dim cnt As Integer = menu.Items.Count
            Dim subMenu As ToolStripMenuItem = Nothing
            For i As Integer = 1 To cnt
                subMenu = CType(menu.Items(i - 1), ToolStripMenuItem)
                subMenu.Enabled = IsEnabled And Not ctrlNameSet.Contains(subMenu.Name)
                If subMenu.HasDropDownItems Then
                    ForbidToolStripMenuItem(subMenu, subMenu.Enabled)
                End If
            Next
            cnt = Nothing
            subMenu = Nothing
        End Sub

        Private Shared Sub ForbidTabControl(ByRef tabCtrl As TabControl, Optional ByVal IsEnabled As Boolean = True)
            Dim cnt As Integer = tabCtrl.TabPages.Count
            Dim tabPg As TabPage = Nothing
            For i As Integer = 1 To cnt
                tabPg = tabCtrl.TabPages.Item(i - 1)
                tabPg.Enabled = IsEnabled And Not ctrlNameSet.Contains(tabPg.Name)
                If tabPg.HasChildren Then
                    ForbidControl(CType(tabPg, Control), tabPg.Enabled)
                End If
            Next
            cnt = Nothing
            tabPg = Nothing
        End Sub

        Private Shared Sub ForbidControl(ByRef ctrl As Control, Optional ByVal IsEnabled As Boolean = True)
            Dim cnt As Integer = ctrl.Controls.Count
            Dim childControl As Control = Nothing
            For i As Integer = 1 To cnt
                childControl = ctrl.Controls.Item(i - 1)
                childControl.Enabled = IsEnabled And Not ctrlNameSet.Contains(childControl.Name)
            Next
            cnt = Nothing
            childControl = Nothing
        End Sub

        ''' <summary>
        ''' 设置一个窗体所包含的部分子控件的可用性属性。
        ''' </summary>
        ''' <param name="ParentForm ">包含子控件的窗体。</param>
        ''' <param name="SubControlNames">子控件名称的集合。</param>
        ''' <param name="IsEnabled" >待设置的可用性属性值。</param>
        ''' <remarks></remarks>
        Public Shared Sub SetControlsEnabled(ByRef ParentForm As Form, ByRef SubControlNames As HashSet(Of String), _
                                             Optional IsEnabled As Boolean = False)
            If SubControlNames Is Nothing Then
                Return
            End If
            If SubControlNames.Count = 0 Then
                Return
            End If
            ctrlNameSet = SubControlNames
            Dim cnt As Integer = ParentForm.Controls.Count
            Dim m_ctrl As Control = Nothing
            Dim t As Type = Nothing
            For i As Integer = 1 To cnt
                m_ctrl = ParentForm.Controls.Item(i - 1)
                If ctrlNameSet.Contains(m_ctrl.Name) Then
                    m_ctrl.Enabled = IsEnabled
                End If
                t = m_ctrl.GetType
                Select Case t.Name
                    Case "MenuStrip"
                        ForbidMenuStrip(CType(m_ctrl, MenuStrip), m_ctrl.Enabled)
                    Case "TabControl"
                        ForbidTabControl(CType(m_ctrl, TabControl), m_ctrl.Enabled)
                    Case "GroupBox"
                        ForbidControl(CType(m_ctrl, Control), m_ctrl.Enabled)
                End Select
            Next
            cnt = Nothing
            m_ctrl = Nothing
            t = Nothing
            ctrlNameSet.Clear()
            ctrlNameSet = Nothing
        End Sub

        ''' <summary>
        ''' 在垂直方向上移动ListBox中的当前选中项目。
        ''' </summary>
        ''' <param name="List">需要移动项目的ListBox控件。</param>
        ''' <param name="UpStep">当前选中项目向上移动的步长。若为向下移动，则取负值。</param>
        ''' <remarks></remarks>
        Public Shared Sub MoveItemVerital(ByRef List As ListBox, UpStep As Integer)
            Dim index As Integer = List.SelectedIndex
            If index <> -1 Then
                Dim cnt As Integer = List.Items.Count
                If cnt >= 2 And index - UpStep >= 0 And index - UpStep <= cnt - 1 Then
                    Dim s As String = CStr(List.Items(index))
                    List.Items.RemoveAt(index)
                    List.Items.Insert(index - UpStep, s)
                    List.SelectedIndex = index - UpStep
                    s = Nothing
                End If
                cnt = Nothing
            End If
            index = Nothing
        End Sub

        ''' <summary>
        ''' 将一个ListBox中的当前选中项目移动到另一个ListBox的末尾，并在原ListBox中删除该项目。
        ''' </summary>
        ''' <param name="SourceListBox">待移出项目的ListBox。</param>
        ''' <param name="DestListBox">待移入项目的ListBox。</param>
        ''' <remarks></remarks>
        Public Shared Sub MoveItemBetweenListBoxes(ByRef SourceListBox As ListBox, ByRef DestListBox As ListBox)
            Dim index As Integer = SourceListBox.SelectedIndex
            If index <> -1 Then
                DestListBox.Items.Add(SourceListBox.SelectedItem)
                SourceListBox.Items.RemoveAt(index)
                Dim count As Integer = SourceListBox.Items.Count
                If count >= index + 1 Then
                    SourceListBox.SelectedIndex = index
                ElseIf count > 0 Then
                    SourceListBox.SelectedIndex = count - 1
                End If
                count = Nothing
            End If
            index = Nothing
        End Sub

    End Class
End Namespace
