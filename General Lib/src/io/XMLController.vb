Imports System.Xml
Imports dotNet.util

Namespace io

    Public Class XMLController

        ''' <summary>
        ''' 根据名称获得指定的XML文档子节点在父节点的子节点列表中的位置。
        ''' </summary>
        ''' <param name="ChildrenList">父节点的子节点列表。</param>
        ''' <param name="ChildName">子节点名称。</param>
        ''' <param name="NumberOfChild">子节点列表中名称为指定名称值的子节点数量。</param>
        ''' <returns>若找到指定名称的子节点，则返回该子节点在列表中的位置序号；否则返回-1。</returns>
        ''' <remarks>若存在同名的子节点，则返回最后一个子节点在列表中的位置序号。</remarks>
        Public Shared Function IndexOfChild(ByRef ChildrenList As XmlNodeList, ByVal ChildName As String, ByVal NumberOfChild As Integer) As Integer
            Dim node As XmlNode = Nothing
            Dim count As Integer = 0, position As Integer = -1
            Dim totalNumberOfChild As Integer = ChildrenList.Count
            For i As Integer = 1 To totalNumberOfChild
                node = ChildrenList(i - 1)
                If node.Name = ChildName Then
                    count = count + 1
                    position = i - 1
                    If count > NumberOfChild Then
                        totalNumberOfChild = Nothing
                        count = Nothing
                        position = Nothing
                        node = Nothing
                        Return -1
                    End If
                End If
            Next
            node = Nothing
            totalNumberOfChild = Nothing
            If count <> NumberOfChild Then
                count = Nothing
                position = Nothing
                Return -1
            Else
                count = Nothing
                Return position
            End If
        End Function

        ''' <summary>
        ''' 根据子节点和孙子节点名称获得指定的XML文档子节点在父节点的子节点列表中的位置。
        ''' </summary>
        ''' <param name="ChildrenList">父节点的子节点列表。</param>
        ''' <param name="ChildName">子节点名称。</param>
        ''' <param name="GrandsonName">孙子节点名称。</param>
        ''' <returns>若找到唯一一个指定名称的子节点和孙子节点，则返回该子节点在列表中的位置序号；否则返回-1。</returns>
        ''' <remarks></remarks>
        Public Shared Function IndexOfChild(ByRef ChildrenList As XmlNodeList, ByVal ChildName As String, ByVal GrandsonName As String) As Integer
            Dim node1 As XmlNode = Nothing, node2 As XmlNode = Nothing
            Dim count As Integer = 0, position As Integer = -1
            Dim totalNumber1 As Integer = ChildrenList.Count, totalNumber2 As Integer = Nothing
            For i As Integer = 1 To totalNumber1
                node1 = ChildrenList(i - 1)
                totalNumber2 = node1.ChildNodes.Count
                For j As Integer = 1 To totalNumber2
                    node2 = node1.ChildNodes(j - 1)
                    If node2.Name = ChildName Then
                        If node2.ChildNodes.Count = 1 And node2.ChildNodes(0).Value = GrandsonName Then
                            count = count + 1
                            position = i - 1
                            If count > 1 Then
                                node1 = Nothing
                                node2 = Nothing
                                count = Nothing
                                position = Nothing
                                totalNumber1 = Nothing
                                totalNumber2 = Nothing
                                Return -1
                            End If
                        End If
                    End If
                Next
            Next
            node1 = Nothing
            node2 = Nothing
            totalNumber1 = Nothing
            totalNumber2 = Nothing
            If count <> 1 Then
                count = Nothing
                position = Nothing
                Return -1
            Else
                count = Nothing
                Return position
            End If
        End Function

        Friend Shared Sub Save(ByRef t As Tree, ByVal fileName As String)
            If IsNothing(t) Then
                Return
            End If

            Dim doc As New XmlDocument
            Dim node As XmlNode = doc.CreateElement(t.Name)
            SaveNode(t, node)

            doc.AppendChild(node)
            doc.Save(fileName)
            node = Nothing
            doc = Nothing
        End Sub

        Private Shared Sub SaveNode(ByRef parentTree As Tree, ByRef parentNode As XmlNode)
            Dim sonNode As XmlNode
            For i As Integer = 0 To parentTree.ChildCount - 1
                With parentNode
                    If parentTree.Child(i).ChildCount = 0 Then
                        sonNode = .OwnerDocument.CreateTextNode(parentTree.Child(i).Name)
                        .AppendChild(sonNode)
                    Else
                        sonNode = .OwnerDocument.CreateElement(parentTree.Child(i).Name)
                        .AppendChild(.OwnerDocument.CreateTextNode(vbCrLf))
                        .AppendChild(.OwnerDocument.CreateTextNode(Space$(4 * parentTree.Level)))
                        SaveNode(parentTree.Child(i), sonNode)
                        .AppendChild(sonNode)
                        If i = parentTree.ChildCount - 1 Then
                            .AppendChild(.OwnerDocument.CreateTextNode(vbCrLf))
                            .AppendChild(.OwnerDocument.CreateTextNode(Space$(4 * (parentTree.Level - 1))))
                        End If
                    End If
                End With
            Next
        End Sub

        Friend Shared Function Load(ByVal fileName As String) As Tree
            Dim doc As New XmlDocument
            doc.Load(fileName)

            If doc.ChildNodes.Count <> 1 Then
                doc = Nothing
                Return Nothing
            End If

            Dim node As XmlNode = doc.ChildNodes(0)
            Dim t As Tree = New Tree(node.Name)
            LoadNode(t, node)
            node = Nothing
            doc = Nothing
            Return t
        End Function

        Private Shared Sub LoadNode(ByRef parentTree As Tree, ByRef parentNode As XmlNode)
            If parentNode.ChildNodes.Count = 0 Then
                parentTree.Add("")
            Else
                For Each sonNode As XmlNode In parentNode.ChildNodes
                    Select Case sonNode.NodeType
                        Case XmlNodeType.Element
                            parentTree.Add(sonNode.Name)
                            LoadNode(parentTree.Child(parentTree.ChildCount() - 1), sonNode)
                        Case XmlNodeType.Text
                            parentTree.Add(sonNode.Value)
                    End Select
                Next
            End If
        End Sub

    End Class
End Namespace