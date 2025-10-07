using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Docnet.Core;
using Docnet.Core.Models;

namespace Doyle_Addin
{


    /// <summary>
    /// 
    /// </summary>
    public static class PdfToImage
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pdfFilePath"></param>
        /// <param name="imageFilePath"></param>
        public static void ExportFirstPageAsImage(string pdfFilePath, string imageFilePath)
        {
            Environment.GetEnvironmentVariable(@"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin");
            string nativeDllPath;
            string fullOriginalPath = Environment.GetEnvironmentVariable("PATH");

            try
            {
                // Get the directory where your add-in DLL is located (bin\Debug or bin\Release)
                string assemblyLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                // Construct the path to the native DLLs within the runtimes folder
                // IMPORTANT: Adjust 'win-x64' if your target platform is x86 or another.
                nativeDllPath = Path.Combine(assemblyLocation, "runtimes", "win-x64", "native");

                if (Directory.Exists(nativeDllPath))
                {
                    // Add the native DLL path to the beginning of the PATH environment variable
                    // This makes it discoverable for the duration of this process.
                    // Before changing PATH, save the full original PATH
                    // Prepend the native DLL path
                    Environment.SetEnvironmentVariable("PATH", nativeDllPath + ";" + fullOriginalPath);

                    // Set desired DPI or pixel dimensions
                    const int dpi = 3200;

                    using (var docReader = DocLib.Instance.GetDocReader(pdfFilePath, new PageDimensions(dpi, dpi)))
                    {
                        using (var pageReader = docReader.GetPageReader(0))
                        {
                            int width = pageReader.GetPageWidth();
                            int height = pageReader.GetPageHeight();
                            byte[] rawBytes = pageReader.GetImage();

                            // Create a bitmap from the raw BGRA bytes
                            using (var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                            {
                                var bmpData = bmp.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, bmp.PixelFormat);
                                Marshal.Copy(rawBytes, 0, bmpData.Scan0, rawBytes.Length);
                                bmp.UnlockBits(bmpData);

                                // Composite onto a white background
                                using (var whiteBmp = new Bitmap(width, height, PixelFormat.Format24bppRgb))
                                {
                                    using (var g = Graphics.FromImage(whiteBmp))
                                    {
                                        g.Clear(Color.White);
                                        g.DrawImage(bmp, 0, 0);
                                    }
                                    whiteBmp.Save(imageFilePath, ImageFormat.Jpeg);
                                }
                            }
                        }
                    }
                }

                else
                {
                    // MsgBox.Show($"Native Docnet DLL path Not found: {nativeDllPath}", "Error", MsgBoxResult.Ok, MsgBoxResult.Cancel)
                }
            }

            // MsgBox.Show($"An error occurred: {ex.Message}", "Docnet Error", MessageBoxButtons.OK, MsgBoxIcon.Error"")

            catch (Exception ex)
            {
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
}