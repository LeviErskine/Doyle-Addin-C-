Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms
Imports stdole
Imports Svg

Friend MustInherit Class PictureConverter
	Inherits AxHost

	Private Sub New()
		MyBase.New(String.Empty)
	End Sub

	Private Shared Function ImageToPictureDisp(image As Image) As IPictureDisp
		Return CType(GetIPictureDispFromPicture(image), IPictureDisp)
	End Function

	' Find a group by "name" or "inkscape:label" attribute
	Public Shared Function SvgResourceToPictureDisp _
		(resourceName As String, width As Integer, height As Integer, layerName As String) As IPictureDisp
		Using stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName)
			If stream Is Nothing Then Throw New FileNotFoundException($"Resource '{resourceName}' not found.")

			Dim svgDoc = SvgDocument.Open (Of SvgDocument)(stream)
			' Search for <g> with name or inkscape:label
			Dim layer = svgDoc.Descendants().OfType (Of SvgGroup)().FirstOrDefault _
				    (Function(g)
					    Return _
					    (g.CustomAttributes.ContainsKey("inkscape:label") AndAlso g.CustomAttributes("inkscape:label") = layerName) OrElse
					    (g.CustomAttributes.ContainsKey("http://www.inkscape.org/namespaces/inkscape:label") AndAlso
					     g.CustomAttributes("http://www.inkscape.org/namespaces/inkscape:label") = layerName)
				    End Function)

			If layer Is Nothing Then
				Throw New ArgumentException($"Layer '{layerName}' not found in SVG.")
			End If

			' Create a new SVG document with just the selected layer
			Dim newDoc As New SvgDocument With {
				    .Width = svgDoc.Width,
				    .Height = svgDoc.Height
				    }
			newDoc.Children.Add(layer.DeepCopy())

			' Calculate scaling
			Dim scale = Math.Min(width/newDoc.Width.Value, height/newDoc.Height.Value)
			Dim newWidth = CInt(newDoc.Width.Value*scale)
			Dim newHeight = CInt(newDoc.Height.Value*scale)

			Using finalBitmap = New Bitmap(width, height)
				Using g = Graphics.FromImage(finalBitmap)
					g.InterpolationMode = InterpolationMode.HighQualityBicubic
					g.Clear(Color.Transparent)
					Dim x = (width - newWidth)/2
					Dim y = (height - newHeight)/2
					g.DrawImage(CType(newDoc.Draw(newWidth, newHeight), Image), CInt(x), CInt(y))
				End Using
				Return ImageToPictureDisp(finalBitmap)
			End Using
		End Using
	End Function
End Class