using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Docnet.Core;
using Docnet.Core.Models;

namespace Doyle_Addin.Prints;

/// <summary>
/// Provides functionality to convert a PDF file to an image file.
/// This utility is primarily focused on exporting the first page of a PDF
/// as an image for various processing or usage purposes.
/// </summary>
public static class PdfToImage
{
    /// <summary>
    /// Exports the first page of a PDF file as an image in JPG format.
    /// </summary>
    /// <param name="pdfFilePath">The file path of the input PDF.</param>
    /// <param name="imageFilePath">The file path where the generated image will be saved.</param>
    public static void ExportFirstPageAsImage(string pdfFilePath, string imageFilePath)
    {
        Environment.GetEnvironmentVariable(@"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin");
        var fullOriginalPath = Environment.GetEnvironmentVariable("PATH");

        try
        {
            // Get the directory where your add-in DLL is located (bin\Debug or bin\Release)
            var assemblyLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Construct the path to the native DLLs within the runtimes folder
            // IMPORTANT: Adjust 'win-x64' if your target platform is x86 or another.
            Debug.Assert(assemblyLocation != null, nameof(assemblyLocation) + " != null");
            var nativeDllPath = Path.Combine(assemblyLocation, "runtimes", "win-x64", "native");

            if (!Directory.Exists(nativeDllPath)) return;
            // Add the native DLL path to the beginning of the PATH environment variable
            // This makes it discoverable for the duration of this process.
            // Before changing PATH, save the full original PATH
            // Prepend the native DLL path
            Environment.SetEnvironmentVariable("PATH", nativeDllPath + ";" + fullOriginalPath);

            // Set desired DPI or pixel dimensions
            const int dpi = 3200;

            using var docReader = DocLib.Instance.GetDocReader(pdfFilePath, new PageDimensions(dpi, dpi));
            using var pageReader = docReader.GetPageReader(0);
            var width = pageReader.GetPageWidth();
            var height = pageReader.GetPageHeight();
            var rawBytes = pageReader.GetImage();

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
            // MsgBox.Show($"Native Docnet DLL path Not found: {nativeDllPath}", "Error", MsgBoxResult.Ok, MsgBoxResult.Cancel)
        }

        // MsgBox.Show($"An error occurred: {ex.Message}", "Docnet Error", MessageBoxButtons.OK, MsgBoxIcon.Error"")

        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
        finally
        {
            // Restore the full original PATH
            if (!string.IsNullOrEmpty(fullOriginalPath))
            {
                Environment.SetEnvironmentVariable("PATH", fullOriginalPath);
            }
        }
    }
}