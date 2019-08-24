Imports dotNet.io

Namespace util
    Public Class Configuration

        Private Shared path As String
        Private Shared type As StoredType = StoredType.XML

		Private Enum StoredType
			XML
			REGISTRY
		End Enum

		Private Shared WriteOnly Property StoreType() As StoredType
			Set(ByVal value As StoredType)
				type = value
			End Set
		End Property

		Public Shared WriteOnly Property StorePath() As String
			Set(ByVal value As String)
				path = value
			End Set
		End Property

		Public Shared Function Load() As Tree
			Select Case type
				Case StoredType.XML
                    Return XMLController.Load(path)
                Case StoredType.REGISTRY
                    Return Registry.Load(path)
            End Select
            Return Nothing
        End Function

        Public Shared Sub Save(ByRef config As Tree)
            Select Case type
                Case StoredType.XML
                    XMLController.Save(config, path)
				Case StoredType.REGISTRY
					Registry.Save(config, path)
			End Select
		End Sub

	End Class
End Namespace
