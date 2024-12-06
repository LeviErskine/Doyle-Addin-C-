Imports Inventor

Module PrintUpdate

    Public Sub RunPrintUpdate()

        'Check if current document is a drawing, show error if not
        If g_inventorApplication.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

            MsgBox("ONLY FOR USE IN DRAWING DOCUMENTS", "Ilogic")
            Return
        Else
        End If

        'Set reference to active document
        Dim oDDoc As DrawingDocument
        oDDoc = g_inventorApplication.ActiveDocument
        Dim oFilePath As String = "P:\"
        Dim FileName As String
        Dim PN As String
        Dim oCamera As Camera = g_inventorApplication.ActiveView.Camera

        'Set filename to document part number
        PN = oDDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString

        Dim DocName As String = oDDoc.DisplayName

        If DocName <> (PN & ".idw") Then
            MsgBox("DOCUMENT NAME IS DIFFERENT FROM PART NUMBER")
            Return
        Else
        End If

        'Count Sheets
        Dim SheetCount As Integer = oDDoc.Sheets.Count
        FileName = PN & ".jpg"

        If SheetCount = 1 Then

            'Export JPEG only	

            'Activate Sheet 1
            oDDoc.Sheets.Item(1).Activate()

            Try
                Dim oPrintMgr As PrintManager
                oPrintMgr = oDDoc.PrintManager
                'specify your printer name
                oPrintMgr.Printer = "Bullzip PDF Printer"

                'oPrintMgr.ColorMode = kPrintAllColorsAsBlack
                oPrintMgr.AllColorsAsBlack = False
                oPrintMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation
                oPrintMgr.Scale = PrintScaleModeEnum.kPrintBestFitScale

                oPrintMgr.SubmitPrint()
            Catch ex As Exception

            End Try

        Else

            'Export both

            'Export JPEG

            'Activate Sheet 1
            oDDoc.Sheets.Item(1).Activate()

            Try

                Dim oPrintMgr As PrintManager
                oPrintMgr = oDDoc.PrintManager
                'specify your printer name
                oPrintMgr.Printer = "Bullzip PDF Printer"

                'oPrintMgr.ColorMode = kPrintAllColorsAsBlack
                oPrintMgr.AllColorsAsBlack = False
                oPrintMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation
                oPrintMgr.Scale = PrintScaleModeEnum.kPrintBestFitScale
                oPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeA4
                oPrintMgr.SubmitPrint()

            Catch ex As Exception

            End Try

            'Export PDF

            ' Check that this file has been saved and actually exists on disk
            If String.IsNullOrEmpty(oFilePath) Then
                MsgBox("This file has not yet been saved and doesn't exist on disk!" _
& vbLf & "Please save it first", 64, "Formsprag iLogic")
                Return
            End If

            FileName = PN & ".pdf"
            Dim oPDFAddin As Inventor.TranslatorAddIn
            oPDFAddin = g_inventorApplication.ApplicationAddIns.ItemById _
("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
            Dim oDocument As Inventor.Document
            oDocument = g_inventorApplication.ActiveDocument
            Dim oContext As Inventor.TranslationContext
            oContext = g_inventorApplication.TransientObjects.CreateTranslationContext
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
            Dim oOptions As Inventor.NameValueMap
            oOptions = g_inventorApplication.TransientObjects.CreateNameValueMap
            Dim oDataMedium As Inventor.DataMedium
            oDataMedium = g_inventorApplication.TransientObjects.CreateDataMedium

            'set PDF Options
            oOptions.Value("All_Color_AS_Black") = 0
            oOptions.Value("Remove_Line_Weights") = 1
            oOptions.Value("Vector_Resolution") = 400
            oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets

            'Set the PDF target file name
            oDataMedium.FileName = oFilePath & "\" & FileName
            Try
                'Publish document
                oPDFAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
            Catch

                MsgBox("Failed to Export PDF", "Test")

                Return
            End Try
        End If

    End Sub

End Module
