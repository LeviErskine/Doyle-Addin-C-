using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using stdole;
using Svg;

namespace Doyle_Addin
{

    internal abstract class PictureConverter : AxHost
    {

        private PictureConverter() : base(string.Empty)
        {
        }

        private static IPictureDisp ImageToPictureDisp(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }

        // Find a group by "name" or "inkscape:label" attribute
        public static IPictureDisp SvgResourceToPictureDisp(string resourceName, int width, int height, string layerName)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream is null)
                    throw new FileNotFoundException($"Resource '{resourceName}' not found.");

                var svgDoc = SvgDocument.Open<SvgDocument>(stream);
                // Search for <g> with name or inkscape:label
                var layer = svgDoc.Descendants().OfType<SvgGroup>().FirstOrDefault(g => g.CustomAttributes.ContainsKey("inkscape:label") && (g.CustomAttributes["inkscape:label"] ?? "") == (layerName ?? "") || g.CustomAttributes.ContainsKey("http://www.inkscape.org/namespaces/inkscape:label") && (g.CustomAttributes["http://www.inkscape.org/namespaces/inkscape:label"] ?? "") == (layerName ?? ""));


                if (layer is null)
                {
                    throw new ArgumentException($"Layer '{layerName}' not found in SVG.");
                }

                // Create a new SVG document with just the selected layer
                var newDoc = new SvgDocument()
                {
                    Width = svgDoc.Width,
                    Height = svgDoc.Height
                };
                newDoc.Children.Add(layer.DeepCopy());

                // Calculate scaling
                float scale = Math.Min(width / newDoc.Width.Value, height / newDoc.Height.Value);
                int newWidth = (int)Math.Round(newDoc.Width.Value * scale);
                int newHeight = (int)Math.Round(newDoc.Height.Value * scale);

                using (var finalBitmap = new Bitmap(width, height))
                {
                    using (var g = Graphics.FromImage(finalBitmap))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.Clear(Color.Transparent);
                        double x = (width - newWidth) / 2d;
                        double y = (height - newHeight) / 2d;
                        g.DrawImage(newDoc.Draw(newWidth, newHeight), (int)Math.Round(x), (int)Math.Round(y));
                    }
                    return ImageToPictureDisp(finalBitmap);
                }
            }
        }
    }
}