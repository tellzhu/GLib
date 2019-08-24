Imports Microsoft.Win32
Imports dotNet.util

Namespace io
	Friend Class Registry

		Friend Shared Sub Save(ByRef t As Tree, ByVal registryPath As String)
			If IsNothing(t) Then
				Return
			End If

			Dim rootKey As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey(registryPath)
			SaveKey(t, rootKey)
			rootKey.Close()
			rootKey = Nothing
		End Sub

		Private Shared Sub SaveKey(ByRef parentTree As Tree, ByRef parentKey As RegistryKey)
			Dim sonKey As RegistryKey
			For i As Integer = 0 To parentTree.ChildCount - 1
				With parentTree.Child(i)
					If .ChildCount = 1 And .Child(0).ChildCount = 0 Then
						parentKey.SetValue(.Name, .Child(0).Name, RegistryValueKind.String)
					Else
						sonKey = parentKey.CreateSubKey(.Name)
						SaveKey(parentTree.Child(i), sonKey)
						sonKey.Close()
						sonKey = Nothing
					End If
				End With
			Next
		End Sub

		Friend Shared Function Load(ByVal registryPath As String) As Tree
			Dim rootKey As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(registryPath)
			If rootKey Is Nothing Then
				Return Nothing
			End If

			Dim t As Tree = New Tree(rootKey.Name)
			LoadKey(t, rootKey)
			rootKey.Close()
			rootKey = Nothing
			Return t
		End Function

		Private Shared Sub LoadKey(ByRef parentTree As Tree, ByRef parentKey As RegistryKey)
			Dim sonKey As RegistryKey
			Dim strs() As String = parentKey.GetSubKeyNames()
			For i As Integer = 0 To strs.Length - 1
				sonKey = parentKey.OpenSubKey(strs(i))
				LoadKey(parentTree.Add(strs(i)), sonKey)
				sonKey.Close()
				sonKey = Nothing
			Next
			strs = Nothing

			strs = parentKey.GetValueNames()
			For i As Integer = 0 To strs.Length - 1
				parentTree.Keys(strs(i)) = CStr(parentKey.GetValue(strs(i)))
			Next
			strs = Nothing
		End Sub

	End Class
End Namespace
