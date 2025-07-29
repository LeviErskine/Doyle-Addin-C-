Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Docnet.Core.Models


Public Module PDFToImage
    ''' <summary>
    ''' Converts the first page of a PDF to a JPEG image using Docnet.Core, with a white background.
    ''' </summary>
    ''' <param name="pdfFilePath">The path to the input PDF file.</param>
    ''' <param name="imageFilePath">The path to the output image file.</param>
    Public Sub ExportFirstPageAsImage(pdfFilePath As String, imageFilePath As String)
        Dim originalPath As String = Environment.GetEnvironmentVariable("C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin")
        Dim nativeDllPath As String = "C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\runtimes\win-x64\native"
        Dim fullOriginalPath As String = Environment.GetEnvironmentVariable("PATH")

        Try
            ' Get the directory where your add-in DLL is located (bin\Debug or bin\Release)
            Dim assemblyLocation As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)

            ' Construct the path to the native DLLs within the runtimes folder
            ' IMPORTANT: Adjust 'win-x64' if your target platform is x86 or another.
            nativeDllPath = Path.Combine(assemblyLocation, "runtimes", "win-x64", "native")

            If Directory.Exists(nativeDllPath) Then
                ' Add the native DLL path to the beginning of the PATH environment variable
                ' This makes it discoverable for the duration of this process.
                ' Before changing PATH, save the full original PATH
                ' Prepend the native DLL path
                Environment.SetEnvironmentVariable("PATH", nativeDllPath & ";" & fullOriginalPath)

                ' Set desired DPI or pixel dimensions
                Dim dpi As Integer = 3200

                Using docReader = DocLib.Instance.GetDocReader(pdfFilePath, New PageDimensions(dpi, dpi))
                    Using pageReader = docReader.GetPageReader(0)
                        Dim width = pageReader.GetPageWidth()
                        Dim height = pageReader.GetPageHeight()
                        Dim rawBytes = pageReader.GetImage()

                        ' Create a bitmap from the raw BGRA bytes
                        Using bmp As New Bitmap(width, height, PixelFormat.Format32bppArgb)
                            Dim bmpData = bmp.LockBits(New Rectangle(0, 0, width, height), Imaging.ImageLockMode.WriteOnly, bmp.PixelFormat)
                            Marshal.Copy(rawBytes, 0, bmpData.Scan0, rawBytes.Length)
                            bmp.UnlockBits(bmpData)

                            ' Composite onto a white background
                            Using whiteBmp As New Bitmap(width, height, PixelFormat.Format24bppRgb)
                                Using g As Graphics = Graphics.FromImage(whiteBmp)
                                    g.Clear(Color.White)
                                    g.DrawImage(bmp, 0, 0)
                                End Using
                                whiteBmp.Save(imageFilePath, ImageFormat.Jpeg)
                            End Using
                        End Using
                    End Using
                End Using

            Else
                '  MsgBox.Show($"Native Docnet DLL path Not found: {nativeDllPath}", "Error", MsgBoxResult.Ok, MsgBoxResult.Cancel)
            End If

        Catch ex As Exception
            ' MsgBox.Show($"An error occurred: {ex.Message}", "Docnet Error", MessageBoxButtons.OK, MsgBoxIcon.Error"")

        Finally
            ' Restore the full original PATH
            If Not String.IsNullOrEmpty(fullOriginalPath) Then
                Environment.SetEnvironmentVariable("PATH", fullOriginalPath)
            End If
        End Try
    End Sub
End Module