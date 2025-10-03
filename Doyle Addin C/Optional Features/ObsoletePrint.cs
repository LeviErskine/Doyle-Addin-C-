using System.IO;
using File = System.IO.File;
using Path = System.IO.Path;
using System.Linq;
using Inventor;

namespace Doyle_Addin
{

    static class ObsoletePrint
    {
        // This is the main routine that runs.
        public static void ApplyObsoletePrint(Application thisApplication)
        {
            // Get the active drawing document
            DrawingDocument drawingDoc = thisApplication.ActiveDocument as DrawingDocument;
            if (drawingDoc is null)
            {
                return;
            }

            foreach (Sheet sheet in drawingDoc.Sheets)
            {
                // Get the appropriate symbol name for this sheet size
                string symbolName = GetSymbolNameForSheetSize(sheet.Size);
                if (string.IsNullOrEmpty(symbolName))
                {
                    continue; // Skip unsupported sheet sizes
                }

                // Get or load the symbol definition
                var symbolDefinition = GetSymbolDefinition(symbolName, drawingDoc, thisApplication);
                if (symbolDefinition is null)
                {
                    continue; // Skip if the symbol cannot be found or loaded
                }

                // Delete existing instances of this symbol on the sheet
                DeleteExistingSymbolInstances(sheet, symbolName);

                // Place the symbol at the center of the sheet
                PlaceSymbolAtSheetCenter(sheet, symbolDefinition, thisApplication);
            }
        }

        // Determines the appropriate OBSOLETE symbol name based on sheet size
        private static string GetSymbolNameForSheetSize(DrawingSheetSizeEnum sheetSize)
        {
            switch (sheetSize)
            {
                case DrawingSheetSizeEnum.kADrawingSheetSize:
                    {
                        return "OBSOLETE A";
                    }
                case DrawingSheetSizeEnum.kBDrawingSheetSize:
                    {
                        return "OBSOLETE B";
                    }
                case DrawingSheetSizeEnum.kCDrawingSheetSize:
                    {
                        return "OBSOLETE C";
                    }
                case DrawingSheetSizeEnum.kDDrawingSheetSize:
                    {
                        return "OBSOLETE D";
                    }
                case DrawingSheetSizeEnum.kEDrawingSheetSize:
                    {
                        return "OBSOLETE E";
                    }

                default:
                    {
                        return string.Empty;
                    }
            }
        }

        // Gets the symbol definition from the document or library
        private static SketchedSymbolDefinition GetSymbolDefinition(string symbolName, DrawingDocument drawingDoc, Application thisApplication)
        {
            // Step 1: Try to get the symbol from the active document itself
            SketchedSymbolDefinition symbolDefinition = null;
            try
            {
                symbolDefinition = drawingDoc.SketchedSymbolDefinitions[symbolName];
            }
            catch
            {
                // Symbol not in the document, need to get it from the library
            }

            if (symbolDefinition is not null)
            {
                return symbolDefinition;
            }

            // Step 2: Search loaded libraries
            SketchedSymbolDefinitionLibrary symbolLibrary;
            try
            {
                symbolLibrary = FindFromLibraries(symbolName, drawingDoc.SketchedSymbolDefinitions.SketchedSymbolDefinitionLibraries);
            }
            catch
            {
                // No libraries loaded or error searching
                symbolLibrary = null;
            }

            // Step 3: If not found in loaded libraries, copy the library and restart
            if (symbolLibrary is null)
            {
                string libraryPath = CopyObsoleteLibrary();
                if (!string.IsNullOrEmpty(libraryPath))
                {
                    // Library copied successfully, restart the entire process
                    ApplyObsoletePrint(thisApplication);
                }
                return null;
            }

            // Step 4: Add the symbol from the library to the document
            try
            {
                symbolDefinition = drawingDoc.SketchedSymbolDefinitions.AddFromLibrary(symbolLibrary, symbolName, true);
            }
            catch
            {
                // Failed to add a symbol from the library
                return null;
            }

            return symbolDefinition;
        }

        // Deletes all existing instances of a symbol with the specified name from the sheet
        private static void DeleteExistingSymbolInstances(Sheet sheet, string symbolName)
        {
            try
            {
                // Iterate through all sketched symbols on the sheet (backwards to avoid collection modification issues)
                for (int i = sheet.SketchedSymbols.Count; i >= 1; i -= 1)
                {
                    var sketchedSymbol = sheet.SketchedSymbols[i];
                    if ((sketchedSymbol.Definition.Name ?? "") == (symbolName ?? ""))
                    {
                        sketchedSymbol.Delete();
                    }
                }
            }
            catch
            {
                // Silently ignore any errors during deletion
            }
        }

        // Places the symbol at the center of the specified sheet
        private static void PlaceSymbolAtSheetCenter(Sheet sheet, SketchedSymbolDefinition symbolDefinition, Application thisApplication)
        {
            // Get the center point of the sheet
            var transientGeometry = thisApplication.TransientGeometry;
            var centerPoint = transientGeometry.CreatePoint2d(sheet.Width / 2d, sheet.Height / 2d);

            var pointCollection = thisApplication.TransientObjects.CreateObjectCollection();
            pointCollection.Add(centerPoint);

            // Add an instance of the symbol to the sheet at the center point
            sheet.SketchedSymbols.AddWithLeader(symbolDefinition, pointCollection, SymbolClipping: false, Static: true);
        }

        // Function to copy the ObsoleteLibrary.idw file to the Symbol Library folder
        // Returns the full path to the copied file if successful, or empty string if failed
        private static string CopyObsoleteLibrary()
        {
            string sourcePath = @"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\Resources\ObsoleteLibrary.idw";
            string destinationPath = @"C:\Users\Public\Documents\Autodesk\Inventor 2025\Design Data\Symbol Library\ObsoleteLibrary.idw";

            try
            {
                // Check if the source file exists
                if (!File.Exists(sourcePath))
                {
                    return string.Empty;
                }

                // Ensure destination directory exists
                string destinationDir = Path.GetDirectoryName(destinationPath);
                if (!Directory.Exists(destinationDir))
                {
                    Directory.CreateDirectory(destinationDir);
                }

                // Copy the file (overwrite if exists)
                File.Copy(sourcePath, destinationPath, true);
                return destinationPath;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static SketchedSymbolDefinitionLibrary FindFromLibraries(string symbolDefinitionName, SketchedSymbolDefinitionLibraries allLibraries)
        {
            foreach (SketchedSymbolDefinitionLibrary library in allLibraries)
            {
                // Search for the definition within this specific library
                var foundDefinition = SearchDefinitions(symbolDefinitionName, library.SketchedSymbolDefinitions);
                if (foundDefinition is not null)
                {
                    // If found, return the library object and exit the function
                    return library;
                }
            }

            // If the loop finishes, the symbol was not found in any library.
            return null;
        }

        // Helper function to search for a definition by name within a collection.
        private static LibrarySketchedSymbolDefinition SearchDefinitions(string searchDefinitionName, LibrarySketchedSymbolDefinitions definitions)
        {
            return definitions.Cast<LibrarySketchedSymbolDefinition>().FirstOrDefault(libraryDefinition => (libraryDefinition.Name ?? "") == (searchDefinitionName ?? ""));


            // If the loop finishes, the definition was not found.
        }
    }
}