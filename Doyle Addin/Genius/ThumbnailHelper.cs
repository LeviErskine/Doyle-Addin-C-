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
	///     Must be called on STA thread (UI thread) due to COM property access.
	/// </summary>
	/// <param name="document">The Inventor document to retrieve the thumbnail from.</param>
	/// <returns>The IPictureDisp object, or null if unavailable.</returns>
	public static object GetThumbnailRaw(Document document)
	{
		try
		{
			if (document == null) return null;

			// Validate document type
			if (document.DocumentType is not kPartDocumentObject and
			    not kAssemblyDocumentObject)
				return null;

			// Try getting thumbnail via PropertySets (Inventor Summary Information)
			// Iterating avoids COMException when the property or set doesn't exist
			try
			{
				foreach (var ps in document.PropertySets.Cast<PropertySet>().Where(ps =>
					         string.Equals(ps.Name, "Inventor Summary Information",
						         StringComparison.OrdinalIgnoreCase)))
				{
					foreach (var prop in from Property prop in ps
					         where string.Equals(prop.Name, "Thumbnail", StringComparison.OrdinalIgnoreCase)
					         where prop.Value != null
					         select prop) return prop.Value;

					break;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"ThumbnailHelper: PropertySets approach failed: {ex.Message}");
			}

			// Fallback: try direct property access via reflection
			var thumbnailPropInfo =
				document.GetType().GetProperty("Thumbnail", BindingFlags.Public | BindingFlags.Instance);
			if (thumbnailPropInfo == null) return null;
			var value = thumbnailPropInfo.GetValue(document);
			return value;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"ThumbnailHelper: Error getting raw thumbnail: {ex.Message}");
			return null;
		}
	}

	/// <summary>
	///     Converts an IPictureDisp object to System.Drawing.Image.
	///     Must be called on STA thread (UI thread) due to COM interop requirements.
	/// </summary>
	/// <param name="pictureDisp">The IPictureDisp COM object.</param>
	/// <returns>A System.Drawing.Image, or null if conversion fails.</returns>
	public static Image ConvertIPictureToImage(object pictureDisp)
	{
		if (pictureDisp == null) return null;

		try
		{
			var axHostType = typeof(AxHost);

			// Try common method names for IPicture conversion in AxHost
			string[] methodNames = ["GetImageFromIPicture", "GetImageFromIPictureDisp", "GetPicture"];
			foreach (var methodName in methodNames)
			{
				var method = axHostType.GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static);
				if (method == null) continue;
				try
				{
					var result = method.Invoke(null, [pictureDisp]);
					if (result is Image img) return img;
				}
				catch
				{
					// Continue to next method
				}
			}

			// Search all methods if specific ones failed
			var methods = axHostType.GetMethods(BindingFlags.NonPublic | BindingFlags.Static);
			foreach (var method in methods)
			{
				if (!method.Name.Contains("Image") && !method.Name.Contains("Picture")) continue;
				try
				{
					var parameters = method.GetParameters();
					if (parameters.Length == 1)
					{
						var result = method.Invoke(null, [pictureDisp]);
						if (result is Image img) return img;
					}
				}
				catch
				{
					// Continue trying other methods
				}
			}

			return null;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"ThumbnailHelper: Error converting IPicture - {ex.GetType().Name}: {ex.Message}");
			return null;
		}
	}

	/// <summary>
	///     Converts a System.Drawing.Image to WPF BitmapImage.
	/// </summary>
	/// <param name="image">The System.Drawing.Image to convert.</param>
	/// <returns>A WPF BitmapImage.</returns>
	public static BitmapImage ConvertToBitmapImage(Image image)
	{
		if (image == null) return null;

		try
		{
			using var memoryStream = new MemoryStream();
			// Use a clone to avoid issues with the source image's state
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