Imports dotNet.i18n.Languager

Namespace text
	Public Class Name

		Private lt As List(Of String)
		Private dictChn As Dictionary(Of String, String)
		Private dictEng As Dictionary(Of String, String)

		Public ReadOnly Property Count() As Integer
			Get
				Return lt.Count
			End Get
		End Property

		Private Sub Load(ByRef chnNames As String(), ByRef engNames As String())
			If chnNames IsNot Nothing Then
				If lt.Count <> chnNames.Length Then
					Return
				End If
				For i As Integer = 0 To lt.Count - 1
					dictChn.Add(lt.Item(i), chnNames(i))
				Next
			End If
			If engNames IsNot Nothing Then
				If lt.Count <> engNames.Length Then
					Return
				End If
				For i As Integer = 0 To lt.Count - 1
					dictEng.Add(lt.Item(i), engNames(i))
				Next
			End If
		End Sub

		Public Sub New(ByRef IDs As String(), ByRef chnNames As String(), ByRef engNames As String())
			Me.New()
			If IDs Is Nothing Or (chnNames Is Nothing And engNames Is Nothing) Then
				Return
			End If
			lt.AddRange(IDs)
			Load(chnNames, engNames)
		End Sub

		Public ReadOnly Property IntegerID(ByVal index As Integer) As Integer
			Get
                Return CInt(StringID(index))
			End Get
		End Property

		Public ReadOnly Property StringID(ByVal index As Integer) As String
			Get
				Return lt.Item(index)
			End Get
		End Property

		Public Function IndexOf(ByVal name As String) As Integer
			If Language = LanguageCategory.CHINESE Then
				For Each pair As KeyValuePair(Of String, String) In dictChn
					If pair.Value = name Then
						Return lt.IndexOf(pair.Key)
					End If
				Next
				Return -1
			Else
				For Each pair As KeyValuePair(Of String, String) In dictEng
					If pair.Value = name Then
						Return lt.IndexOf(pair.Key)
					End If
				Next
				Return -1
			End If
		End Function

		Public Sub New(ByRef IDs As Integer(), ByRef chnNames As String(), ByRef engNames As String())
			Me.New()
			If IDs Is Nothing Or (chnNames Is Nothing And engNames Is Nothing) Then
				Return
			End If
			For i As Integer = 0 To IDs.Length - 1
				lt.Add(CStr(IDs(i)))
			Next
			Load(chnNames, engNames)
		End Sub

		Public ReadOnly Property Name(ByVal id As String) As String
			Get
				If Language = LanguageCategory.CHINESE Then
					Return dictChn.Item(id)
				Else
					Return dictEng.Item(id)
				End If
			End Get
		End Property

		Public ReadOnly Property Name(ByVal id As Integer) As String
			Get
				Return Name(CStr(id))
			End Get
		End Property

		Public ReadOnly Property NameAt(ByVal index As Integer) As String
			Get
				Return Name(lt.Item(index))
			End Get
		End Property

		Public Sub New()
			lt = New List(Of String)
			dictChn = New Dictionary(Of String, String)
			dictEng = New Dictionary(Of String, String)
		End Sub

		Protected Overrides Sub Finalize()
			lt.Clear()
			lt = Nothing
			dictChn.Clear()
			dictChn = Nothing
			dictEng.Clear()
			dictEng = Nothing
			MyBase.Finalize()
		End Sub
	End Class
End Namespace
