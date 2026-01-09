Imports System.IO
Imports Inventor
Imports File = System.IO.File
Imports Path = System.IO.Path

Module ObsoletePrint
	' This is the main routine that runs.
	Sub ApplyObsoletePrint(thisApplication As Application)
		' Get the active drawing document
		Dim drawingDoc As DrawingDocument = TryCast(thisApplication.ActiveDocument, DrawingDocument)
		If drawingDoc Is Nothing Then
			Return
		End If

		For Each sheet As Sheet In drawingDoc.Sheets
			' Get the appropriate symbol name for this sheet size
			Dim symbolName As String = GetSymbolNameForSheetSize(sheet.Size)
			If String.IsNullOrEmpty(symbolName) Then
				Continue For ' Skip unsupported sheet sizes
			End If

			' Get or load the symbol definition
			Dim symbolDefinition As SketchedSymbolDefinition = GetSymbolDefinition(symbolName, drawingDoc, thisApplication)
			If symbolDefinition Is Nothing Then
				Continue For ' Skip if the symbol cannot be found or loaded
			End If

			' Delete existing instances of this symbol on the sheet
			DeleteExistingSymbolInstances(sheet, symbolName)

			' Place the symbol at the center of the sheet
			PlaceSymbolAtSheetCenter(sheet, symbolDefinition, thisApplication)
		Next
	End Sub

	' Determines the appropriate OBSOLETE symbol name based on sheet size
	Private Function GetSymbolNameForSheetSize(sheetSize As DrawingSheetSizeEnum) As String
		Select Case sheetSize
			Case DrawingSheetSizeEnum.kADrawingSheetSize
				Return "OBSOLETE A"
			Case DrawingSheetSizeEnum.kBDrawingSheetSize
				Return "OBSOLETE B"
			Case DrawingSheetSizeEnum.kCDrawingSheetSize
				Return "OBSOLETE C"
			Case DrawingSheetSizeEnum.kDDrawingSheetSize
				Return "OBSOLETE D"
			Case DrawingSheetSizeEnum.kEDrawingSheetSize
				Return "OBSOLETE E"
			Case Else
				Return String.Empty
		End Select
	End Function

	' Gets the symbol definition from the document or library
	Private Function GetSymbolDefinition _
		(symbolName As String, drawingDoc As DrawingDocument, thisApplication As Application) As SketchedSymbolDefinition
		' Step 1: Try to get the symbol from the active document itself
		Dim symbolDefinition As SketchedSymbolDefinition = Nothing
		Try
			symbolDefinition = drawingDoc.SketchedSymbolDefinitions.Item(symbolName)
		Catch
			' Symbol not in the document, need to get it from the library
		End Try

		If symbolDefinition IsNot Nothing Then
			Return symbolDefinition
		End If

		' Step 2: Search loaded libraries
		Dim symbolLibrary As SketchedSymbolDefinitionLibrary
		Try
			symbolLibrary = FindFromLibraries(symbolName, drawingDoc.SketchedSymbolDefinitions.SketchedSymbolDefinitionLibraries)
		Catch
			' No libraries loaded or error searching
			symbolLibrary = Nothing
		End Try

		' Step 3: If not found in loaded libraries, copy the library and restart
		If symbolLibrary Is Nothing Then
			Dim libraryPath As String = CopyObsoleteLibrary()
			If Not String.IsNullOrEmpty(libraryPath) Then
				' Library copied successfully, restart the entire process
				ApplyObsoletePrint(thisApplication)
			End If
			Return Nothing
		End If

		' Step 4: Add the symbol from the library to the document
		Try
			symbolDefinition = drawingDoc.SketchedSymbolDefinitions.AddFromLibrary(symbolLibrary, symbolName, True)
		Catch
			' Failed to add a symbol from the library
			Return Nothing
		End Try

		Return symbolDefinition
	End Function

	' Deletes all existing instances of a symbol with the specified name from the sheet
	Private Sub DeleteExistingSymbolInstances(sheet As Sheet, symbolName As String)
		Try
			' Iterate through all sketched symbols on the sheet (backwards to avoid collection modification issues)
			For i As Integer = sheet.SketchedSymbols.Count To 1 Step - 1
				Dim sketchedSymbol As SketchedSymbol = sheet.SketchedSymbols.Item(i)
				If sketchedSymbol.Definition.Name = symbolName Then
					sketchedSymbol.Delete()
				End If
			Next
		Catch
			' Silently ignore any errors during deletion
		End Try
	End Sub

	' Places the symbol at the center of the specified sheet
	Private Sub PlaceSymbolAtSheetCenter _
		(sheet As Sheet, symbolDefinition As SketchedSymbolDefinition, thisApplication As Application)
		' Get the center point of the sheet
		Dim transientGeometry As TransientGeometry = thisApplication.TransientGeometry
		Dim centerPoint As Point2d = transientGeometry.CreatePoint2d(sheet.Width/2, sheet.Height/2)

		Dim pointCollection As ObjectCollection = thisApplication.TransientObjects.CreateObjectCollection
		pointCollection.Add(centerPoint)

		' Add an instance of the symbol to the sheet at the center point
		sheet.SketchedSymbols.AddWithLeader(symbolDefinition, pointCollection, , , , False, True)
	End Sub

	' Function to copy the ObsoleteLibrary.idw file to the Symbol Library folder
	' Returns the full path to the copied file if successful, or empty string if failed
	Private Function CopyObsoleteLibrary() As String
		Dim sourcePath As String = "C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\Resources\ObsoleteLibrary.idw"
		Dim destinationPath As String =
			    "C:\Users\Public\Documents\Autodesk\Inventor 2025\Design Data\Symbol Library\ObsoleteLibrary.idw"

		Try
			' Check if the source file exists
			If Not File.Exists(sourcePath) Then
				Return String.Empty
			End If

			' Ensure destination directory exists
			Dim destinationDir As String = Path.GetDirectoryName(destinationPath)
			If Not Directory.Exists(destinationDir) Then
				Directory.CreateDirectory(destinationDir)
			End If

			' Copy the file (overwrite if exists)
			File.Copy(sourcePath, destinationPath, True)
			Return destinationPath
		Catch
			Return String.Empty
		End Try
	End Function

	Private Function FindFromLibraries(symbolDefinitionName As String, allLibraries As SketchedSymbolDefinitionLibraries) _
		As SketchedSymbolDefinitionLibrary
		For Each library As SketchedSymbolDefinitionLibrary In allLibraries
			' Search for the definition within this specific library
			Dim foundDefinition As LibrarySketchedSymbolDefinition = SearchDefinitions _
				    (symbolDefinitionName, library.SketchedSymbolDefinitions)
			If foundDefinition IsNot Nothing Then
				' If found, return the library object and exit the function
				Return library
			End If
		Next

		' If the loop finishes, the symbol was not found in any library.
		Return Nothing
	End Function

	' Helper function to search for a definition by name within a collection.
	Private Function SearchDefinitions(searchDefinitionName As String, definitions As LibrarySketchedSymbolDefinitions) _
		As LibrarySketchedSymbolDefinition
		Return _
			definitions.Cast (Of LibrarySketchedSymbolDefinition)().FirstOrDefault _
				(Function(libraryDefinition) libraryDefinition.Name = searchDefinitionName)

		' If the loop finishes, the definition was not found.
	End Function
End Module