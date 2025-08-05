Imports Doyle_Addin.Options
Imports Inventor
Imports File = System.IO.File

Module PrintUpdate
    Public Sub RunPrintUpdate(thisApplication As Application)
        ' Check if the current document is a drawing, show error if not
        If thisApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
            MsgBox("ONLY FOR USE IN DRAWING DOCUMENTS", "Ilogic")
            Return
        End If

        ' Set reference to active document
        Dim oDDoc As DrawingDocument = thisApplication.ActiveDocument

        ' Gets referenced model document type (part or assembly)
        Dim refDocType As DocumentTypeEnum = oDDoc.ReferencedDocuments(1).DocumentType

        Dim oFilePath As String = UserOptions.Load.PrintExportLocation
        Dim fileName As String
        Dim pn As String = oDDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString

        ' Always export PDF
        If String.IsNullOrEmpty(oFilePath) Then
            MsgBox("This file has not been saved yet or save location cannot be found" & vbCrLf, 64, "Save error")
            Return
        End If

        fileName = pn & ".pdf"
        Dim pdfPath As String = oFilePath & "\" & fileName
        Dim oPdfAddin As TranslatorAddIn =
                thisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
        Dim oDocument As Document = thisApplication.ActiveDocument
        Dim oContext As TranslationContext = thisApplication.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        Dim oOptions As NameValueMap = thisApplication.TransientObjects.CreateNameValueMap
        Dim oDataMedium As DataMedium = thisApplication.TransientObjects.CreateDataMedium

        ' Set PDF options
        oOptions.Value("All_Color_AS_Black") = 0
        oOptions.Value("Remove_Line_Weights") = 0
        oOptions.Value("Vector_Resolution") = 4800
        oOptions.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets

        ' Set PDF target file name
        oDataMedium.FileName = pdfPath
        Try
            ' Publish document
            oPdfAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
            ExportFirstPageAsImage(pdfPath, oFilePath & "\" & pn & ".jpg")
        Catch
            MsgBox("Failed to Export PDF (Someone might have this file open)", MsgBoxStyle.OkOnly)
            Return
        End Try

        ' If referenced document is a part, convert PDF to JPG and then delete the PDF
        If refDocType = DocumentTypeEnum.kPartDocumentObject Then
            Try
                ExportFirstPageAsImage(pdfPath, oFilePath & "\" & pn & ".jpg")
                ' Delete the PDF after conversion
                If File.Exists(pdfPath) Then
                    File.Delete(pdfPath)
                End If
            Catch ex As Exception
                MsgBox("Failed to convert PDF to JPG: " & ex.Message, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub
End Module