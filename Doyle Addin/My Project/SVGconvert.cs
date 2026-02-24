namespace DoyleAddin.My_Project;

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Svg;

// ReSharper disable once IdentifierTypo
// ReSharper disable once InconsistentNaming
internal static class SVGconvert

{
	private static object ImageToPictureDisp(Image image)
	{
		// Use reflection to access the internal AxHost.GetIPictureDispFromPicture method
		var axHostType = typeof(AxHost);
		var method = axHostType.GetMethod("GetIPictureDispFromPicture",
			BindingFlags.Static | BindingFlags.NonPublic);

		if (method == null) throw new InvalidOperationException("Cannot access GetIPictureDispFromPicture method");

		return method.Invoke(null, new object[] { image });
	}

	// Find a group by "name" or "inkscape:label" attribute
	public static object SvgResourceToPictureDisp(string resourceName, int width, int height,
		string layerName)
	{
		using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName) ??
		                   throw new FileNotFoundException($"Resource '{resourceName}' not found.");
		var svgDoc = SvgDocument.Open<SvgDocument>(stream);
		// Search for <g> with name or inkscape:label
		var layer = svgDoc.Descendants().OfType<SvgGroup>().FirstOrDefault(g =>
			(g.CustomAttributes.ContainsKey("inkscape:label") &&
			 (g.CustomAttributes["inkscape:label"] ?? "") == (layerName ?? "")) ||
			(g.CustomAttributes.ContainsKey("http://www.inkscape.org/namespaces/inkscape:label") &&
			 (g.CustomAttributes["http://www.inkscape.org/namespaces/inkscape:label"] ?? "") ==
			 (layerName ?? ""))) ?? throw new ArgumentException($"Layer '{layerName}' not found in SVG.");

		// Create a new SVG document with just the selected layer
		var newDoc = new SvgDocument
		{
			Width  = svgDoc.Width,
			Height = svgDoc.Height
		};
		newDoc.Children.Add(layer.DeepCopy());

		// Calculate scaling
		var scale     = Math.Min(width / newDoc.Width.Value, height / newDoc.Height.Value);
		var newWidth  = (int)Math.Round(newDoc.Width.Value * scale);
		var newHeight = (int)Math.Round(newDoc.Height.Value * scale);

		using var finalBitmap = new Bitmap(width, height);
		using (var g = Graphics.FromImage(finalBitmap))
		{
			g.InterpolationMode = InterpolationMode.HighQualityBicubic;
			g.Clear(Color.Transparent);
			var x = (width - newWidth) / 2d;
			var y = (height - newHeight) / 2d;
			g.DrawImage(newDoc.Draw(newWidth, newHeight), (int)Math.Round(x), (int)Math.Round(y));
		}

		return ImageToPictureDisp(finalBitmap);
	}
}