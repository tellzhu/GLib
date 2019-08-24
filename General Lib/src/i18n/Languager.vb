Namespace i18n
    Public Class Languager

        Enum LanguageCategory
            CHINESE
            ENGLISH
        End Enum

		Friend Shared ReadOnly Property FontName() As String
			Get
				Return currentFontName
			End Get
		End Property

        Public Shared Property Language() As LanguageCategory
            Get
                Return currentLanguage
            End Get
            Set(ByVal value As LanguageCategory)
                currentLanguage = value
			End Set
        End Property

        Private Shared currentLanguage As LanguageCategory
		Private Shared currentFontName As String = "Times New Roman"
    End Class
End Namespace
