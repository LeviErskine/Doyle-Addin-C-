Imports System.IO
Imports System.Xml.Serialization

Namespace Options
	Public Class UserOptions
		Public Property PrintExportLocation As String = ""
		Public Property DxfExportLocation As String = ""

		' Feature toggles
		Public Property EnableObsoletePrint As Boolean = True

		Public Shared ReadOnly _
			OptionsFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
			                                         "DoyleAddinOptions.xml")

		Public Sub Save()
			Dim serializer As New XmlSerializer(GetType(UserOptions))
			Using writer As New StreamWriter(OptionsFilePath)
				serializer.Serialize(writer, Me)
			End Using
		End Sub

		Public Shared Function Load() As UserOptions
			If File.Exists(OptionsFilePath) Then
				Dim serializer As New XmlSerializer(GetType(UserOptions))
				Using reader As New StreamReader(OptionsFilePath)
					Return CType(serializer.Deserialize(reader), UserOptions)
				End Using
			Else
				Return New UserOptions()
			End If
		End Function
	End Class
End NameSpace