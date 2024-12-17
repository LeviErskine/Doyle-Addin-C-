Imports Inventor

Module PrintUpdate

    Public Sub RunPrintUpdate(ThisApplication As Inventor.Application)
        ' Check if current document is a drawing, show error if not
        If ThisApplication.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
            MsgBox("ONLY FOR USE IN DRAWING DOCUMENTS", "Ilogic")
            Return
        End If

        ' Set reference to active document
        Dim oDDoc As DrawingDocument = ThisApplication.ActiveDocument

        ' Get referenced model document type (part or assembly)
        Dim refDocType As Inventor.DocumentTypeEnum = oDDoc.ReferencedDocuments(1).DocumentType

        Dim oFilePath As String = "P:\"
        Dim FileName As String
        Dim PN As String = oDDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
        Dim DocName As String = oDDoc.DisplayName

        ' Check if document name matches part number
        If DocName <> (PN & ".idw") Then
            MsgBox("DOCUMENT NAME IS DIFFERENT FROM PART NUMBER")
            Return
        End If

        ' Count sheets
        Dim SheetCount As Integer = oDDoc.Sheets.Count

        ' Always export PDF
        If String.IsNullOrEmpty(oFilePath) Then
            MsgBox("This file has not yet been saved and doesn't exist on disk!" & vbCrLf & "Please save it first", 64, "Formsprag iLogic")
            Return
        End If

        FileName = PN & ".pdf"
        Dim pdfPath As String = oFilePath & "\" & FileName
        Dim oPDFAddin As Inventor.TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
        Dim oDocument As Inventor.Document = ThisApplication.ActiveDocument
        Dim oContext As Inventor.TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        Dim oOptions As Inventor.NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap
        Dim oDataMedium As Inventor.DataMedium = ThisApplication.TransientObjects.CreateDataMedium

        ' Set PDF options
        oOptions.Value("All_Color_AS_Black") = 0
        oOptions.Value("Remove_Line_Weights") = 1
        oOptions.Value("Vector_Resolution") = 400
        oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets

        ' Set PDF target file name
        oDataMedium.FileName = pdfPath
        Try
            ' Publish document
            oPDFAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
        Catch
            MsgBox("Failed to Export PDF (Someone might have this file open)", MsgBoxStyle.OkOnly)
            Return
        End Try

        ' If referenced document is a part, convert PDF to JPG and then delete the PDF
        If refDocType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
            Try
                ExportFirstPageAsImage(pdfPath, oFilePath & "\" & PN & ".jpg")
                ' Delete the PDF after conversion
                If IO.File.Exists(pdfPath) Then
                    IO.File.Delete(pdfPath)
                End If
            Catch ex As Exception
                MsgBox("Failed to convert PDF to JPG: " & ex.Message, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub

End Module