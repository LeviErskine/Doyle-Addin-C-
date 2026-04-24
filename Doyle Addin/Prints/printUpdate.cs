namespace DoyleAddin.Prints;

using System.Windows.Forms;
using Docnet.Core;
using Docnet.Core.Models;
using Inventor;
using Options;
using Environment = Environment;
using File = File;
using Path = Path;

internal static class PrintUpdate
{
	public static void RunPrintUpdate()
	{
		// Check if the current document is a drawing, show error if not
		if (ThisApplication.ActiveDocument.DocumentType != kDrawingDocumentObject)
		{
			MessageBox.Show("ONLY FOR USE IN DRAWING DOCUMENTS", "Ilogic", MessageBoxButtons.OK,
				MessageBoxIcon.Warning);
			return;
		}

		// Set reference to active document

		// Gets referenced model document type (part or assembly)
		if (ThisApplication.ActiveDocument is not DrawingDocument oDDoc) return;
		var refDocType = oDDoc.ReferencedDocuments[1].DocumentType;

		var oFilePath = UserOptions.Load().PrintExportLocation;
		var pn        = oDDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();

		// Check if Part Number matches Filename
		var fileNameWithoutExt = Path.GetFileNameWithoutExtension(oDDoc.FullFileName);
		if (!string.IsNullOrEmpty(fileNameWithoutExt) &&
		    !string.Equals(pn, fileNameWithoutExt, StringComparison.OrdinalIgnoreCase))
		{
			var result = MessageBox.Show(
				$"The document's Part Number '{pn}' does not match the filename '{fileNameWithoutExt}'." +
				Environment.NewLine +
				"Update to match the filename?" +
				Environment.NewLine +
				Environment.NewLine +
				"If this was intentional select 'No'",
				"Part Number Mismatch",
				MessageBoxButtons.YesNoCancel,
				MessageBoxIcon.Warning);

			switch (result)
			{
				case DialogResult.Yes:
					oDDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value = fileNameWithoutExt;
					pn = fileNameWithoutExt; // Update pn variable for subsequent use
					break;
				case DialogResult.Cancel:
					return; // Stop the export process
			}
			// If result is DialogResult.No, continue with the current pn without updating
		}

		// Always export PDF
		if (string.IsNullOrEmpty(oFilePath))
		{
			MessageBox.Show(
				"This file has not been saved yet or save location cannot be found" + Environment.NewLine,
				"Save error", MessageBoxButtons.OK, MessageBoxIcon.Information);
			return;
		}

		var fileName = pn + ".pdf";
		var pdfPath  = Path.Combine(oFilePath, fileName);
		var oPdfAddin =
			(TranslatorAddIn)ThisApplication.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"];
		Document oDocument = ThisApplication.ActiveDocument;
		var      oContext  = ThisApplication.TransientObjects.CreateTranslationContext();
		oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;
		var oOptions    = ThisApplication.TransientObjects.CreateNameValueMap();
		var oDataMedium = ThisApplication.TransientObjects.CreateDataMedium();

		// Set PDF options
		oOptions.Value["All_Color_AS_Black"]  = 0;
		oOptions.Value["Remove_Line_Weights"] = 0;
		oOptions.Value["Vector_Resolution"]   = 4800;
		oOptions.Value["Sheet_Range"]         = PrintRangeEnum.kPrintAllSheets;

		// Set a PDF target file name
		oDataMedium.FileName = pdfPath;
		try
		{
			// Publish document
			oPdfAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium);
			PdfToImage.ExportFirstPageAsImage(pdfPath, Path.Combine(oFilePath, pn + ".jpg"));
		}
		catch
		{
			MessageBox.Show("Failed to Export PDF (Someone might have this file open)", "Export failed",
				MessageBoxButtons.OK, MessageBoxIcon.Error);
			return;
		}

		// If the referenced document is a part, convert PDF to JPG
		if (refDocType != kPartDocumentObject) return;
		try
		{
			const int dpi = 3200;
			int       pageCount;
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

			if (pageCount == 1)
			{
				// Single page - use existing logic
				PdfToImage.ExportFirstPageAsImage(pdfPath, Path.Combine(oFilePath, pn + ".jpg"));
				if (File.Exists(pdfPath)) File.Delete(pdfPath);
			}
			else
			{
				// Multipage - export each page with its own part number but keep the PDF
				PdfToImage.ExportMultiPagePartImages(pdfPath, oFilePath, pageCount, dpi, pn,
					pageIndex => GetPartNumberForPage(oDDoc, pageIndex));
				// Keep the PDF file for multipage documents
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Failed to convert PDF to JPG: " + ex.Message, "Conversion failed", MessageBoxButtons.OK,
				MessageBoxIcon.Error);
		}
	}

	private static string GetPartNumberForPage(DrawingDocument drawingDoc, int pageIndex)
	{
		try
		{
			// Get the sheet for this page (0-based index)
			if (pageIndex >= drawingDoc.Sheets.Count) return string.Empty;

			var sheet = drawingDoc.Sheets[pageIndex + 1]; // Sheets are 1-based in Inventor

			// Look for drawing views on this sheet
			foreach (DrawingView view in sheet.DrawingViews)
				try
				{
					// Get the referenced document for this view
					if (view.ReferencedDocumentDescriptor?.ReferencedDocument is not Document refDocument) continue;
					// Get the part-number from the referenced document
					var partNumberProp = refDocument.PropertySets["Design Tracking Properties"]["Part Number"];
					if (partNumberProp?.Value != null) return partNumberProp.Value.ToString();
				}
				catch
				{
					// Continue to next view if this one fails
				}
		}
		catch
		{
			// Return empty if any error occurs
		}

		return string.Empty;
	}
}