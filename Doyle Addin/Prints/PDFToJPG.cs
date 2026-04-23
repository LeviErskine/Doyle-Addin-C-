namespace DoyleAddin.Prints;

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;

/// <summary>
///     Provides functionality to convert a PDF file to an image file.
///     This utility is primarily focused on exporting the first page of a PDF
///     as an image for various processing or usage purposes.
/// </summary>
public static class PdfToImage
{
	/// <summary>
	///     Exports the first page of a PDF file as an image in JPG format.
	/// </summary>
	/// <param name="pdfFilePath">The file path of the input PDF.</param>
	/// <param name="imageFilePath">The file path where the generated image will be saved.</param>
	public static void ExportFirstPageAsImage(string pdfFilePath, string imageFilePath)
	{
		using var pathManager = new PathManager();

		try
		{
			// Set desired DPI or pixel dimensions
			const int dpi = 3200;

			using var docReader = DocLib.Instance.GetDocReader(pdfFilePath, new PageDimensions(dpi, dpi));
			ExportPageAsImage(docReader, 0, imageFilePath);
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
		}
	}

	/// <summary>
	///     Exports a specific page of a PDF file as an image in JPG format.
	/// </summary>
	/// <param name="docReader">The PDF document reader.</param>
	/// <param name="pageIndex">The zero-based index of the page to export.</param>
	/// <param name="imageFilePath">The file path where the generated image will be saved.</param>
	private static void ExportPageAsImage(IDocReader docReader, int pageIndex, string imageFilePath)
	{
		try
		{
			using var pageReader = docReader.GetPageReader(pageIndex);
			var       width      = pageReader.GetPageWidth();
			var       height     = pageReader.GetPageHeight();
			var       rawBytes   = pageReader.GetImage();

			// Create a bitmap from the raw BGRA bytes
			using var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
			var bmpData = bmp.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly,
				bmp.PixelFormat);
			Marshal.Copy(rawBytes, 0, bmpData.Scan0, rawBytes.Length);
			bmp.UnlockBits(bmpData);

			// Composite onto a white background
			using var whiteBmp = new Bitmap(width, height, PixelFormat.Format24bppRgb);
			using (var g = Graphics.FromImage(whiteBmp))
			{
				g.Clear(Color.White);
				g.DrawImage(bmp, 0, 0);
			}

			whiteBmp.Save(imageFilePath, ImageFormat.Jpeg);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Failed to export page {pageIndex}: {ex.Message}");
		}
	}

	/// <summary>
	///     Exports multiple pages of a PDF as individual images, using page-specific part numbers.
	/// </summary>
	/// <param name="pdfPath">The file path of the input PDF.</param>
	/// <param name="outputPath">The directory where images will be saved.</param>
	/// <param name="pageCount">The number of pages to export.</param>
	/// <param name="dpi">The DPI for image export.</param>
	/// <param name="drawingPartNumber">Fallback part number if a page-specific one isn't found.</param>
	/// <param name="getPartNumberForPage">Function to get part number for a specific page.</param>
	public static void ExportMultiPagePartImages(string pdfPath, string outputPath, int pageCount, int dpi,
		string drawingPartNumber, Func<int, string> getPartNumberForPage)
	{
		using var pathManager = new PathManager();

		try
		{
			using var docReader = DocLib.Instance.GetDocReader(pdfPath, new PageDimensions(dpi, dpi));

			for (var pageIndex = 0; pageIndex < pageCount; pageIndex++)
			{
				var pagePartNumber = getPartNumberForPage(pageIndex);

				// Export this page as an image with the page-specific part number
				ExportPageAsImage(docReader, pageIndex,
					!string.IsNullOrEmpty(pagePartNumber)
						? Path.Combine(outputPath, pagePartNumber + ".jpg")
						// Fallback to drawing part number if page-specific part number not found
						: Path.Combine(outputPath, drawingPartNumber + $"_page{pageIndex + 1}.jpg"));
			}
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Failed to export multi-page images: {ex.Message}");
		}
	}

	/// <summary>
	///     Manages the PATH environment variable for native DLL access.
	/// </summary>
	private sealed class PathManager : IDisposable
	{
		private readonly string _originalPath;
		private bool _disposed;

		public PathManager()
		{
			_originalPath = Environment.GetEnvironmentVariable("PATH");

			try
			{
				// Get the directory where your add-in DLL is located (bin\Debug or bin\Release)
				var assemblyLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

				if (assemblyLocation == null) return;
				var nativeDllPath = Path.Combine(assemblyLocation, "runtimes", "win-x64", "native");

				if (Directory.Exists(nativeDllPath))
					// Add the native DLL path to the beginning of the PATH environment variable
					Environment.SetEnvironmentVariable("PATH", nativeDllPath + ";" + _originalPath);
			}
			catch
			{
				// If PATH setup fails, continue without it
			}
		}

		public void Dispose()
		{
			if (_disposed) return;
			// Restore the full original PATH
			if (!string.IsNullOrEmpty(_originalPath))
				Environment.SetEnvironmentVariable("PATH", _originalPath);
			_disposed = true;
		}
	}
}