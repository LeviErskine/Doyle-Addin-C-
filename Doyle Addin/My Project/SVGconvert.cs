using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Svg;

namespace DoyleAddin.My_Project;

// ReSharper disable once IdentifierTypo
// ReSharper disable once InconsistentNaming
internal static class SVGconvert

{
	private static IPictureDisp ImageToPictureDisp(Image image)
	{
		return (IPictureDisp)AxHostConverter.GetIPictureDispFromPicture(image);
	}

	// Find a group by "name" or "inkscape:label" attribute
	public static IPictureDisp SvgResourceToPictureDisp(string resourceName, int width, int height,
		string layerName)
	{
		using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
		if (stream is null)
			throw new FileNotFoundException($"Resource '{resourceName}' not found.");

		var svgDoc = SvgDocument.Open<SvgDocument>(stream);
		// Search for <g> with name or inkscape:label
		var layer = svgDoc.Descendants().OfType<SvgGroup>().FirstOrDefault(g =>
			(g.CustomAttributes.ContainsKey("inkscape:label") &&
			 (g.CustomAttributes["inkscape:label"] ?? "") == (layerName ?? "")) ||
			(g.CustomAttributes.ContainsKey("http://www.inkscape.org/namespaces/inkscape:label") &&
			 (g.CustomAttributes["http://www.inkscape.org/namespaces/inkscape:label"] ?? "") ==
			 (layerName ?? "")));

		if (layer is null) throw new ArgumentException($"Layer '{layerName}' not found in SVG.");

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

	/// <summary>
	///     Helper class to convert .NET Image to COM IPictureDisp
	/// </summary>
	#pragma warning disable CS0189 // Class is never instantiated
	// ReSharper disable once ClassNeverInstantiated.Local
	private sealed class AxHostConverter : AxHost
	{
		private AxHostConverter() : base("")
		{
		}

		public new static stdole.IPictureDisp GetIPictureDispFromPicture(Image image)
		{
			return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
		}
	}
	#pragma warning restore CS0189
}