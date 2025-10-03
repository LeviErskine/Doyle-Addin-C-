using System;
using File = System.IO.File;
using Docnet.Core;
using Docnet.Core.Models;
using Inventor;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Doyle_Addin
{

    static class PrintUpdate
    {
        public static void RunPrintUpdate(Application thisApplication)
        {
            // Check if the current document is a drawing, show error if not
            if (thisApplication.ActiveDocument.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
            {
                Interaction.MsgBox("ONLY FOR USE IN DRAWING DOCUMENTS", (MsgBoxStyle)Conversions.ToInteger("Ilogic"));
                return;
            }

            // Set reference to active document
            DrawingDocument oDDoc = (DrawingDocument)thisApplication.ActiveDocument;

            // Gets referenced model document type (part or assembly)
            var refDocType = oDDoc.ReferencedDocuments[1].DocumentType;

            string oFilePath = Doyle_Addin.Options.UserOptions.Load().PrintExportLocation;
            string fileName;
            string pn = oDDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();

            // Always export PDF
            if (string.IsNullOrEmpty(oFilePath))
            {
                Interaction.MsgBox("This file has not been saved yet or save location cannot be found" + Constants.vbCrLf, (MsgBoxStyle)64, "Save error");
                return;
            }

            fileName = pn + ".pdf";
            string pdfPath = oFilePath + @"\" + fileName;
            TranslatorAddIn oPdfAddin = (TranslatorAddIn)thisApplication.ApplicationAddIns.get_ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}");
            Document oDocument = thisApplication.ActiveDocument;
            var oContext = thisApplication.TransientObjects.CreateTranslationContext();
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            var oOptions = thisApplication.TransientObjects.CreateNameValueMap();
            var oDataMedium = thisApplication.TransientObjects.CreateDataMedium();

            // Set PDF options
            oOptions.set_Value("All_Color_AS_Black", 0);
            oOptions.set_Value("Remove_Line_Weights", 0);
            oOptions.set_Value("Vector_Resolution", 4800);
            oOptions.set_Value("Sheet_Range", PrintRangeEnum.kPrintAllSheets);

            // Set a PDF target file name
            oDataMedium.FileName = pdfPath;
            try
            {
                // Publish document
                oPdfAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium);
                Doyle_Addin.PdfToImage.ExportFirstPageAsImage(pdfPath, oFilePath + @"\" + pn + ".jpg");
            }
            catch
            {
                Interaction.MsgBox("Failed to Export PDF (Someone might have this file open)", MsgBoxStyle.OkOnly);
                return;
            }

            // If the referenced document is a part, convert PDF to JPG
            if (refDocType == DocumentTypeEnum.kPartDocumentObject)
            {
                try
                {
                    const int dpi = 3200;
                    int pageCount;
                    try
                    {
                        using (var docReader = DocLib.Instance.GetDocReader(pdfPath, new PageDimensions(dpi, dpi)))
                        {
                            pageCount = docReader.GetPageCount();
                        }
                    }
                    catch
                    {
                        // If we can't determine page count, assume a single page
                        pageCount = 1;
                    }

                    // Only convert and delete if single page
                    if (pageCount == 1)
                    {
                        Doyle_Addin.PdfToImage.ExportFirstPageAsImage(pdfPath, oFilePath + @"\" + pn + ".jpg");
                        if (File.Exists(pdfPath))
                        {
                            File.Delete(pdfPath);
                        }
                    }
                    else
                    {
                        // Optionally, notify user or handle multipage PDFs as needed
                        // MsgBox("PDF has multiple pages and will not be converted to JPG or deleted.", MsgBoxStyle.Information)
                    }
                }
                catch (Exception ex)
                {
                    Interaction.MsgBox("Failed to convert PDF to JPG: " + ex.Message, MsgBoxStyle.OkOnly);
                }
            }
        }
    }
}