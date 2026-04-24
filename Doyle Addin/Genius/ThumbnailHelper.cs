namespace DoyleAddin.Genius;

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Inventor;
using Image = System.Drawing.Image;

/// <summary>
///     Provides utility methods for working with Inventor document thumbnails.
/// </summary>
public static class ThumbnailHelper
{
	/// <summary>
	///     Retrieves the raw thumbnail object (IPictureDisp) from an Inventor document.
	/// </summary>
	public static object GetThumbnailRaw(Document document)
	{
		if (document == null) return null;

		try
		{
			// Try getting thumbnail via PropertySets (Summary Information)
			var summaryProps = document.PropertySets.Cast<PropertySet>()
			                           .FirstOrDefault(ps =>
				                           ps.InternalName == "{F29F5501-2E01-11D0-A6E1-00A0C922E752}" ||
				                           ps.Name == GeniusConstants.SummaryInformation);

			if (summaryProps == null) return document.GetType().GetProperty("Thumbnail")?.GetValue(document);
			foreach (var prop in summaryProps.Cast<Property>()
			                                 .Where(prop => prop.Name == "Thumbnail" && prop.Value != null))
				return prop.Value;

			// Fallback: direct property access
			return document.GetType().GetProperty("Thumbnail")?.GetValue(document);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"ThumbnailHelper: Error getting raw thumbnail: {ex.Message}");
			return null;
		}
	}

	/// <summary>
	///     Converts an IPictureDisp object to System.Drawing.Image.
	/// </summary>
	public static Image ConvertIPictureToImage(object pictureDisp)
	{
		if (pictureDisp == null) return null;

		try
		{
			var      axHostType  = typeof(AxHost);
			string[] methodNames = ["GetImageFromIPicture", "GetImageFromIPictureDisp", "GetPicture"];

			foreach (var methodName in methodNames)
			{
				var method = axHostType.GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static);
				if (method == null) continue;
				try
				{
					if (method.Invoke(null, [pictureDisp]) is Image img) return img;
				}
				catch
				{
					/* continue */
				}
			}

			// Brute force search
			return axHostType.GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
			                 .Where(m => m.Name.Contains("Image") || m.Name.Contains("Picture"))
			                 .Select(m =>
			                 {
				                 try
				                 {
					                 return m.Invoke(null, [pictureDisp]) as Image;
				                 }
				                 catch
				                 {
					                 return null;
				                 }
			                 })
			                 .FirstOrDefault(img => img != null);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"ThumbnailHelper: Error converting IPicture: {ex.Message}");
			return null;
		}
	}

	/// <summary>
	///     Converts a System.Drawing.Image to WPF BitmapImage.
	/// </summary>
	public static BitmapImage ConvertToBitmapImage(Image image)
	{
		if (image == null) return null;

		try
		{
			using var memoryStream = new MemoryStream();
			using (var clone = new Bitmap(image))
			{
				clone.Save(memoryStream, ImageFormat.Png);
			}

			memoryStream.Position = 0;

			var bitmapImage = new BitmapImage();
			bitmapImage.BeginInit();
			bitmapImage.StreamSource = memoryStream;
			bitmapImage.CacheOption  = BitmapCacheOption.OnLoad;
			bitmapImage.EndInit();
			if (bitmapImage.CanFreeze) bitmapImage.Freeze();
			return bitmapImage;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"ThumbnailHelper: Error converting to BitmapImage: {ex.Message}");
			return null;
		}
	}
}