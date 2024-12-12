Imports Inventor

Module PrintUpdate

    Public Sub RunPrintUpdate()
        ' Check if current document is a drawing, show error if not
        If g_inventorApplication.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
            MsgBox("ONLY FOR USE IN DRAWING DOCUMENTS", "Ilogic")
            Return
        End If

        ' Set reference to active document
        Dim oDDoc As DrawingDocument = g_inventorApplication.ActiveDocument

        'temporary
        Dim oFilePath As String = "W:\Blake\TESTFILES\TESTPRINTS"
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
        FileName = PN & ".jpg"

        If SheetCount = 1 Then
            ' Export JPEG only
            oDDoc.Sheets.Item(1).Activate()
            Try
                Dim oPrintMgr As PrintManager = oDDoc.PrintManager
                oPrintMgr.Printer = "BullZip Pdf Printer"
                oPrintMgr.AllColorsAsBlack = False
                oPrintMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation
                oPrintMgr.Scale = PrintScaleModeEnum.kPrintBestFitScale
                oPrintMgr.SubmitPrint()
            Catch ex As Exception
                MsgBox("Failed to Export JPG", "Test")
                Return
            End Try
        Else
            ' Export PDF
            If String.IsNullOrEmpty(oFilePath) Then
                MsgBox("This file has not yet been saved and doesn't exist on disk!" & vbCrLf & "Please save it first", 64, "Formsprag iLogic")
                Return
            End If

            FileName = PN & ".pdf"
            Dim oPDFAddin As Inventor.TranslatorAddIn = g_inventorApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
            Dim oDocument As Inventor.Document = g_inventorApplication.ActiveDocument
            Dim oContext As Inventor.TranslationContext = g_inventorApplication.TransientObjects.CreateTranslationContext
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
            Dim oOptions As Inventor.NameValueMap = g_inventorApplication.TransientObjects.CreateNameValueMap
            Dim oDataMedium As Inventor.DataMedium = g_inventorApplication.TransientObjects.CreateDataMedium

            ' Set PDF options
            oOptions.Value("All_Color_AS_Black") = 0
            oOptions.Value("Remove_Line_Weights") = 1
            oOptions.Value("Vector_Resolution") = 400
            oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets

            ' Set PDF target file name
            oDataMedium.FileName = oFilePath & "\" & FileName
            Try
                ' Publish document
                oPDFAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
            Catch
                MsgBox("Failed to Export PDF", "Test")
                Return
            End Try
        End If
    End Sub

End Module