Imports dotNet.i18n.Languager

Namespace util
	Public Class Dictionary

		Private dict As Dictionary(Of String, String)

		Public Sub LoadWords(ByRef chnWords As String(), ByRef engWords As String())
            If chnWords Is Nothing Or engWords Is Nothing Then
                Return
            End If
			If chnWords.Length <> engWords.Length Then
				Return
			End If
			For i As Integer = 0 To chnWords.Length - 1
				dict.Add(chnWords(i), engWords(i))
			Next
		End Sub

		Public ReadOnly Property Word(ByVal keyWord As String) As String
			Get
				If Language = LanguageCategory.CHINESE Then
					Return keyWord
				Else
					Return dict.Item(keyWord)
				End If
			End Get
		End Property

		Public Sub New()
			dict = New Dictionary(Of String, String)
		End Sub

		Protected Overrides Sub Finalize()
			dict.Clear()
			dict = Nothing
			MyBase.Finalize()
		End Sub
	End Class
End Namespace
