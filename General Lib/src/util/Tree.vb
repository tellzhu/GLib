Namespace util

    Public Class Tree

        Private parent As Tree
        Private data As Object
        Private children As ArrayList
        Private separatorChar As String = "."

        Public Sub New(ByVal rootNode As String)
            Me.parent = Nothing
            data = rootNode
            children = Nothing
        End Sub

        Public Function Add(ByVal child As String) As Tree
            Return Add(New Tree(child))
        End Function

		Private Function Add(ByRef child As Tree) As Tree
			If IsNothing(children) Then
				children = New ArrayList()
			End If
			If IsNothing(child) Then
				Return Nothing
			End If
			child.parent = Me
			children.Add(child)
			Return child
		End Function

        Private Function IndexOf(ByVal name As String) As Integer
            If children Is Nothing Then
                Return -1
            Else
                Dim t As Tree
                For i As Integer = 0 To children.Count - 1
                    t = CType(children(i), Tree)
                    If t.Name = name Then
						Return i
					End If
                Next
				Return -1
			End If
        End Function

        Public Sub Rename(ByVal name As String)
            Me.data = name
        End Sub

        Public Sub Remove(ByVal child As String)
            If children IsNot Nothing Then
                children.Remove(Me.Child(child))
            End If
        End Sub

		Friend ReadOnly Property Name() As String
			Get
				Return data.ToString()
			End Get
		End Property

		Friend ReadOnly Property ChildCount() As Integer
			Get
				If IsNothing(children) Then
					Return 0
				Else
					Return children.Count
				End If
			End Get
		End Property

		Public Property Keys(ByVal key As String) As String
			Get
				Dim index As Integer = key.IndexOf(separatorChar)
				If index = -1 Then
					If IsNothing(children) Then
						Return Nothing
					Else
						Dim t As Tree = Child(key)
						If IsNothing(t) Then
							Return Nothing
						Else
							If t.ChildCount = 1 Then
								If t.Child(0).ChildCount = 0 Then
									Return t.Child(0).Name
								Else
									Return Nothing
								End If
							Else
								Return Nothing
							End If
						End If
					End If
				Else
					Return Child(key.Substring(0, index)).Keys(key.Substring(index + 1))
				End If
			End Get
			Set(ByVal value As String)
				Dim index As Integer = key.IndexOf(separatorChar)
				If index = -1 Then
					If IsNothing(children) Then
						Add(key)
						Child(0).Add(value)
					Else
						Dim childIndex As Integer = IndexOf(key)
						If childIndex <> -1 Then
							Dim t As Tree = Child(childIndex)
							If t.ChildCount = 1 Then
								t.Child(0).data = value
							End If
						Else
							Add(key)
							Child(ChildCount - 1).Add(value)
						End If
					End If
				Else
					If IsNothing(Child(key.Substring(0, index))) Then
						Add(key.Substring(0, index)).Keys(key.Substring(index + 1)) = value
					Else
						Child(key.Substring(0, index)).Keys(key.Substring(index + 1)) = value
					End If
				End If
			End Set
		End Property

		Friend ReadOnly Property Level() As Integer
			Get
				Dim t As Tree = Me
				Dim i As Integer = 1
				Do While Not IsNothing(t.parent)
					t = t.parent
					i += 1
				Loop
				Return i
			End Get
		End Property

		Friend ReadOnly Property Child(ByVal index As Integer) As Tree
			Get
				If IsNothing(children) Then
					Return Nothing
				Else
					Return CType(children(index), Tree)
				End If
			End Get
		End Property

		Public ReadOnly Property Child(ByVal name As String) As Tree
			Get
				Dim index As Integer = IndexOf(name)
				If index = -1 Then
					Return Nothing
				Else
					Return Child(index)
				End If
			End Get
		End Property

		Protected Overrides Sub Finalize()
			If Not IsNothing(children) Then
				children.Clear()
				children = Nothing
			End If
			data = Nothing
			parent = Nothing
			MyBase.Finalize()
		End Sub

	End Class
End Namespace