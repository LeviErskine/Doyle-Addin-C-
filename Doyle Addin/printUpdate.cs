using Docnet.Core;
using Docnet.Core.Models;
using Doyle_Addin.Options;
using Inventor;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using File = System.IO.File;
using Application = Inventor.Application;

namespace Doyle_Addin
{
    internal static class PrintUpdate
    {
        public static void RunPrintUpdate(Application? application)
        {
            var app = application;
            if (app is null) return;

            // Check if the current document is a drawing, show error if not
            if (app.ActiveDocument.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
            {
                Interaction.MsgBox("ONLY FOR USE IN DRAWING DOCUMENTS", (MsgBoxStyle)Conversions.ToInteger("Ilogic"));
                return;
            }

            // Set reference to active document
            var oDDoc = (DrawingDocument)app.ActiveDocument;

            // Gets referenced model document type (part or assembly)
            var refDocType = oDDoc.ReferencedDocuments[1].DocumentType;

            var oFilePath = UserOptions.Load().PrintExportLocation;
            var pn = oDDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();

            // Always export PDF
            if (string.IsNullOrEmpty(oFilePath))
            {
                Interaction.MsgBox(
                    "This file has not been saved yet or save location cannot be found" + Constants.vbCrLf,
                    (MsgBoxStyle)64, "Save error");
                return;
            }

            var fileName = pn + ".pdf";
            var pdfPath = oFilePath + @"\" + fileName;
            var oPdfAddin =
                (TranslatorAddIn)app.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"];
            Document oDocument = app.ActiveDocument;
            var oContext = app.TransientObjects.CreateTranslationContext();
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            var oOptions = app.TransientObjects.CreateNameValueMap();
            var oDataMedium = app.TransientObjects.CreateDataMedium();

            // Set PDF options
            oOptions.Value["All_Color_AS_Black"] = 0;
            oOptions.Value["Remove_Line_Weights"] = 0;
            oOptions.Value["Vector_Resolution"] = 4800;
            oOptions.Value["Sheet_Range"] = PrintRangeEnum.kPrintAllSheets;

            // Set a PDF target file name
            oDataMedium.FileName = pdfPath;
            try
            {
                // Publish document
                oPdfAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium);
                PdfToImage.ExportFirstPageAsImage(pdfPath, oFilePath + @"\" + pn + ".jpg");
            }
            catch
            {
                Interaction.MsgBox("Failed to Export PDF (Someone might have this file open)");
                return;
            }

            // If the referenced document is a part, convert PDF to JPG
            if (refDocType != DocumentTypeEnum.kPartDocumentObject) return;
            try
            {
                const int dpi = 3200;
                int pageCount;
                try
                {
                    using var docReader = DocLib.Instance.GetDocReader(pdfPath, new PageDimensions(dpi, dpi));
                    pageCount = docReader.GetPageCount();
                }
                catch
                {
                    // If we can't determine page count, assume a single page
                    pageCount = 1;
                }

                // Only convert and delete if single page
                if (pageCount == 1)
                {
                    PdfToImage.ExportFirstPageAsImage(pdfPath, oFilePath + @"\" + pn + ".jpg");
                    if (File.Exists(pdfPath))
                    {
                        File.Delete(pdfPath);
                    }
                }
                else
                {
                    // Optionally, notify the user or handle multipage PDFs as needed
                    // MsgBox("PDF has multiple pages and will not be converted to JPG or deleted.", MsgBoxStyle.Information)
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox("Failed to convert PDF to JPG: " + ex.Message);
            }
        }
    }
}