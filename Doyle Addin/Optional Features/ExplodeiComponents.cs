namespace DoyleAddin.Optional_Features;

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ClosedXML.Excel;
using Inventor;
using IOFile = File;
using IOPath = Path;

/// <summary>
///     Provides functionality to explode iPart and iAssembly factories into individual standard parts and assemblies
/// </summary>
public static class ExplodeiComponents
{
	/// <summary>
	///     Main entry point to explode iPart/iAssembly into individual standard parts/assemblies
	/// </summary>
	public static void ExplodeiComponentsAction()
	{
		Debug.WriteLine("=== ExplodeiComponentsAction Started ===");

		var activeDoc = ThisApplication.ActiveDocument;

		switch (activeDoc)
		{
			case PartDocument partDoc:
			{
				Debug.WriteLine($"Active document: {partDoc.FullFileName}");
				Debug.WriteLine($"Is iPart Factory: {partDoc.ComponentDefinition.IsiPartFactory}");

				if (partDoc.ComponentDefinition.IsiPartFactory)
				{
					ExplodeFactory(partDoc);
				}
				else
				{
					Debug.WriteLine("Document is not an iPart factory");
					MessageBox.Show("The active document is not an iPart factory.",
						"Explode iComponents", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}

				break;
			}
			case AssemblyDocument asmDoc:
			{
				Debug.WriteLine($"Active document: {asmDoc.FullFileName}");
				Debug.WriteLine($"Is iAssembly Factory: {asmDoc.ComponentDefinition.IsiAssemblyFactory}");

				if (asmDoc.ComponentDefinition.IsiAssemblyFactory)
				{
					ExplodeFactory(asmDoc);
				}
				else
				{
					Debug.WriteLine("Document is not an iAssembly factory");
					MessageBox.Show("The active document is not an iAssembly factory.",
						"Explode iComponents", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}

				break;
			}
			default:
				Debug.WriteLine("No active Part or Assembly document found");
				MessageBox.Show(
					"Please open an iPart or an iAssembly.",
					"Explode iComponents", MessageBoxButtons.OK, MessageBoxIcon.Information);
				break;
		}

		Debug.WriteLine("=== ExplodeiComponentsAction Finished ===");
		ThisApplication.Documents.CloseAll(true);
	}

	/// <summary>
	///     Explodes an iPart or iAssembly factory into individual standard components
	/// </summary>
	private static void ExplodeFactory(dynamic factoryDoc)
	{
		try
		{
			Debug.WriteLine("=== ExplodeFactory Started ===");

			switch (factoryDoc)
			{
				case PartDocument partDoc:
					ExplodeiPartFactory(partDoc);
					break;

				case AssemblyDocument asmDoc:
					ExplodeiAssemblyFactory(asmDoc);
					break;

				default:
					throw new ArgumentException("Unsupported document type.");
			}
		}
		catch (Exception ex)
		{
			ThisApplication.Documents.CloseAll(true);
			Debug.WriteLine($"ExplodeFactory failed: {ex}");
			MessageBox.Show($"Error during component explosion: {ex.Message}",
				"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
	}

	// ===================================================================
	// iPART – New refactored approach with working copies
	// ===================================================================
	private static void ExplodeiPartFactory(PartDocument factoryDoc)
	{
		try
		{
			Debug.WriteLine("=== Explode iPart Factory Started ===");

			var factory = factoryDoc.ComponentDefinition.iPartFactory;

			if (factory == null)
			{
				MessageBox.Show("Could not access the iPart factory.",
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			var xlsxOutputDirectory = GetComponentDirectory((dynamic)factoryDoc);
			var tempXlsxPath        = CreateExcelBackup(factory, xlsxOutputDirectory, "iPart Members");
			Debug.WriteLine($"Excel backup created at: {tempXlsxPath}");

			var processedCount = 0;
			var totalCount     = factory.TableRows.Count;

			// Process each member
			for (var i = 1; i <= totalCount; i++)
				try
				{
					Debug.WriteLine($"--- Processing iPart member {i} of {totalCount} ---");

					factory = factoryDoc.ComponentDefinition.iPartFactory; // Refresh factory reference
					if (factory == null)
						throw new InvalidOperationException("Could not access iPart factory.");

					var row = factory.TableRows[i];
					factory.DefaultRow = row;
					factoryDoc.Update();

					var partNumber = GetPartNumber((dynamic)factoryDoc);
					Debug.WriteLine($"Part number: {partNumber}");

					// Get the original file path where this member would be generated
					var originalPath = IOPath.Combine(GetComponentDirectory((dynamic)factoryDoc),
						$"{partNumber}.ipt");
					// Overwrite any existing target file
					if (IOFile.Exists(originalPath))
					{
						Debug.WriteLine("Deleting existing target file...");
						IOFile.Delete(originalPath);
					}

					var workingCopyPath =
						IOPath.Combine(IOPath.GetTempPath(), $"iPart_WorkingCopy_{Guid.NewGuid():N}.ipt");
					Debug.WriteLine($"Creating working copy: {workingCopyPath}");
					factoryDoc.SaveAs(workingCopyPath, true);

					dynamic workingDoc = (PartDocument)ThisApplication.Documents.Open(workingCopyPath, false);
					workingDoc.Update();

					// Set the part number iProperty to ensure it's retained
					try
					{
						var designTrackingProps = workingDoc.PropertySets["Design Tracking Properties"];
						designTrackingProps["Part Number"].Value = partNumber;
						Debug.WriteLine($"Set Part Number iProperty to: {partNumber}");
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"Failed to set Part Number iProperty: {ex.Message}");
					}

					try
					{
						ConvertToStandardComponent(workingDoc);

						// Ensure the part number iProperty is set correctly after conversion
						try
						{
							var designTrackingProps = workingDoc.PropertySets["Design Tracking Properties"];
							designTrackingProps["Part Number"].Value = partNumber;
							Debug.WriteLine($"Re-set Part Number iProperty to: {partNumber}");
						}
						catch (Exception ex)
						{
							Debug.WriteLine($"Failed to re-set Part Number iProperty: {ex.Message}");
						}

						Debug.WriteLine("Saving standard copy...");
						SaveAsStandard(workingDoc, originalPath, "ipt");

						processedCount++;

						// Update progress
						ThisApplication.StatusBarText =
							$"Processed {processedCount} of {totalCount} members: {partNumber}";
						Debug.WriteLine($"Member {i} processed successfully");
					}
					catch (InvalidOperationException ex)
					{
						Debug.WriteLine($"Error processing member {i}: {ex}");
						MessageBox.Show(
							$"Failed to convert iPart to standard component for member {i}:\n{ex.Message}",
							"Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"Unexpected error processing member {i}: {ex}");
						MessageBox.Show($"Unexpected error processing member {i}:\n{ex.Message}",
							"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					finally
					{
						try
						{
							workingDoc.Close();
						}
						catch (Exception ex)
						{
							Debug.WriteLine($"Failed to close working copy document: {ex}");
						}
						finally
						{
							// Release the COM reference to prevent memory leaks
							try
							{
								Marshal.ReleaseComObject(workingDoc);
							}
							catch (Exception ex)
							{
								Debug.WriteLine($"Failed to release COM object: {ex}");
							}
						}

						try
						{
							if (IOFile.Exists(workingCopyPath))
								IOFile.Delete(workingCopyPath);
						}
						catch (Exception ex)
						{
							Debug.WriteLine($"Failed to delete working copy file: {ex}");
						}
					}
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"Error processing member {i}: {ex}");
					MessageBox.Show($"Error processing member {i}: {ex.Message}",
						"Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}

			Debug.WriteLine($"=== Explode iPart Completed - Processed {processedCount} members ===");
			MessageBox.Show(
				$"Successfully exploded {processedCount} iParts into standard parts.\n\n" +
				$"Table backup saved here:\n{xlsxOutputDirectory}",
				"Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Explode iPart failed: {ex}");
			MessageBox.Show($"Error during iPart explosion: {ex.Message}",
				"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
	}

	// ===================================================================
	// iASSEMBLY – Simplified version (exactly what you said works great)
	// ===================================================================
	private static void ExplodeiAssemblyFactory(AssemblyDocument factoryDoc)
	{
		var factory = factoryDoc.ComponentDefinition.iAssemblyFactory;
		ExplodeFactoryCoreSimplified(factoryDoc, factory, "iam");
	}

	private static void ExplodeFactoryCoreSimplified(dynamic factoryDoc, dynamic factory, string extension)
	{
		var totalCount = factory.TableRows.Count;

		// Clear the cache directory before generating members
		var cacheDir = factory.MemberCacheDir ?? IOPath.GetDirectoryName(factoryDoc.FullFileName) ?? string.Empty;
		if (Directory.Exists(cacheDir))
			foreach (var file in Directory.GetFiles(cacheDir))
				IOFile.Delete(file);

		// First, generate all members in the cache
		for (var i = 1; i <= totalCount; i++)
		{
			var row = factory.TableRows[i];
			factory.DefaultRow = row;
			factoryDoc.Save();
			factory.CreateMember(row);
			Marshal.ReleaseComObject(row);
		}

		var memberFiles = Directory.GetFiles(cacheDir, $"*.{extension}");

		foreach (var cacheFilePath in memberFiles)
		{
			var fileName = IOPath.GetFileName(cacheFilePath);
			var destinationPath =
				IOPath.Combine(IOPath.GetDirectoryName(factoryDoc.FullFileName) ?? string.Empty, fileName);

			// Open the member from the cache
			var memberDoc = ThisApplication.Documents.Open(cacheFilePath, false);

			// Break the link to make it a standard component
			BreakFactoryLink(memberDoc);

			// Save and move to the permanent location
			memberDoc.Save();
			memberDoc.Close();
			ThisApplication.Documents.CloseAll(true);
			Marshal.ReleaseComObject(memberDoc);

			if (IOFile.Exists(destinationPath)) IOFile.Delete(destinationPath);
			if (destinationPath != null) IOFile.Copy(cacheFilePath, destinationPath);
		}

		Debug.WriteLine($"=== Explode iAssembly Completed - Processed {memberFiles.Length} members ===");
		MessageBox.Show(
			$"Successfully exploded {memberFiles.Length} iAssemblies into standard assemblies.",
			"Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
		ThisApplication.Documents.CloseAll(true);
	}

	private static void BreakFactoryLink(dynamic memberDoc)
	{
		switch (memberDoc)
		{
			case PartDocument pDoc when pDoc.ComponentDefinition.IsiPartMember:
				pDoc.ComponentDefinition.iPartMember.BreakLinkToFactory();
				break;
			case AssemblyDocument aDoc when aDoc.ComponentDefinition.IsiAssemblyMember:
				aDoc.ComponentDefinition.iAssemblyMember.BreakLinkToFactory();
				break;
		}
	}

	// ====================== Excel Backup and Helper Methods ======================

	/// <summary>
	///     Creates an Excel backup of the iComponent member table
	/// </summary>
	private static string CreateExcelBackup(dynamic factory, string outputDirectory, string sheetName)
	{
		if (string.IsNullOrWhiteSpace(outputDirectory)) return string.Empty;

		var tempPath = IOPath.Combine(outputDirectory, $"iComponent_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

		try
		{
			if (!Directory.Exists(outputDirectory))
				Directory.CreateDirectory(outputDirectory);

			using var workbook  = new XLWorkbook();
			var       worksheet = workbook.Worksheets.Add(sheetName);

			var tableColumns = factory.TableColumns;
			var tableRows    = factory.TableRows;
			int columnCount  = tableColumns.Count;
			int rowCount     = tableRows.Count;

			Debug.WriteLine($"Creating Excel backup: {columnCount} columns, {rowCount} rows");

			// Write headers (1-based indexing for Inventor COM collections)
			for (var col = 1; col <= columnCount; col++)
			{
				var column = tableColumns[col];
				worksheet.Cell(1, col).Value = column.DisplayHeading ?? $"Column_{col}";
			}

			// Style headers
			var headerRow = worksheet.Range(1, 1, 1, columnCount);
			headerRow.Style.Font.Bold            = true;
			headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

			// Write data for each row
			for (var row = 1; row <= rowCount; row++)
			{
				var tableRow = tableRows[row];
				for (var col = 1; col <= columnCount; col++)
				{
					var cellValue = tableRow[col].Value;
					if (cellValue != null)
						worksheet.Cell(row + 1, col).Value = cellValue.ToString();
				}
			}

			worksheet.Columns().AdjustToContents();
			workbook.SaveAs(tempPath);
			Debug.WriteLine($"Excel backup saved: {tempPath}");
			return tempPath;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"CreateExcelBackup exception: {ex}");
			MessageBox.Show($"Warning: Could not create Excel backup: {ex.Message}",
				"Backup Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return tempPath;
		}
	}

	// ====================== Remaining original helper methods (unchanged) ======================

	private static string GetComponentDirectory(Document factoryDoc)
	{
		var fullFileName = factoryDoc.FullFileName ?? string.Empty;
		var factoryDir   = IOPath.GetDirectoryName(fullFileName);
		var factoryBase  = IOPath.GetFileNameWithoutExtension(fullFileName);
		return IOPath.Combine(factoryDir ?? string.Empty, factoryBase);
	}

	/// <summary>
	///     Saves the document as a standard component with the specified filename
	/// </summary>
	private static void SaveAsStandard(dynamic doc, string filePath, string extension)
	{
		Debug.WriteLine($"SaveAsStandard called with path: {filePath}");

		// Ensure directory exists
		var directory = IOPath.GetDirectoryName(filePath);
		Debug.WriteLine($"Target directory: {directory}");

		if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
		{
			Debug.WriteLine($"Creating directory: {directory}");
			Directory.CreateDirectory(directory);
		}

		// Save as copy using temp file to avoid SaveAs issues
		Debug.WriteLine("Saving document to " + filePath);
		var tempFilePath = IOPath.Combine(IOPath.GetTempPath(), $"Export_{Guid.NewGuid():N}.{extension}");
		Debug.WriteLine("Saving document to temp path " + tempFilePath);
		doc.SaveAs(tempFilePath, true);
		Debug.WriteLine("Temp save successful; copying to target");
		IOFile.Copy(tempFilePath, filePath, true);
		IOFile.Delete(tempFilePath);
		Debug.WriteLine("Copy to target successful");
	}

	private static void ConvertToStandardComponent(Document doc)
	{
		Debug.WriteLine("ConvertToStandardComponent called");

		switch (doc)
		{
			case PartDocument partDoc:
				Debug.WriteLine($"Is iPart Factory before conversion: {partDoc.ComponentDefinition.IsiPartFactory}");
				if (partDoc.ComponentDefinition.IsiPartFactory)
					try
					{
						// Update document to ensure consistent state
						partDoc.Update();

						// Save the document before deleting the factory
						partDoc.Save();

						// Proceed directly to factory deletion
						Debug.WriteLine("Attempting to delete iPart factory...");
						partDoc.ComponentDefinition.iPartFactory.Delete();
						Debug.WriteLine("iPart factory deleted");

						// Try to update after deletion
						try
						{
							partDoc.Update();
							Debug.WriteLine("Document updated after factory deletion");
						}
						catch (COMException updateEx)
						{
							Debug.WriteLine(
								$"Post-deletion update failed, but factory was deleted: {updateEx.Message}");
						}
					}
					catch (COMException ex)
					{
						Debug.WriteLine($"COM Exception deleting iPart factory: {ex.Message}");
						throw new InvalidOperationException($"Failed to delete iPart factory: {ex.Message}", ex);
					}
				else
					Debug.WriteLine("Document is not an iPart factory - nothing to delete");

				break;

			case AssemblyDocument asmDoc:
				Debug.WriteLine(
					$"Is iAssembly Factory before conversion: {asmDoc.ComponentDefinition.IsiAssemblyFactory}");
				if (asmDoc.ComponentDefinition.IsiAssemblyFactory)
					try
					{
						// Update document to ensure consistent state
						asmDoc.Update();

						// Save the document before deleting the factory
						asmDoc.Save();

						// Suppress all constraints to prevent issues during factory deletion
						Debug.WriteLine("Suppressing all assembly constraints...");
						foreach (AssemblyConstraint constraint in asmDoc.ComponentDefinition.Constraints)
							try
							{
								constraint.Suppressed = true;
							}
							catch (Exception ex)
							{
								Debug.WriteLine($"Failed to suppress constraint {constraint.Name}: {ex.Message}");
							}

						Debug.WriteLine("All constraints suppressed");

						// Update the document to ensure consistency after suppressing constraints
						try
						{
							asmDoc.Update();
							Debug.WriteLine("Document updated after suppressing constraints");
						}
						catch (Exception ex)
						{
							Debug.WriteLine($"Failed to update document after suppressing constraints: {ex.Message}");
						}

						// Collect excluded occurrences before deleting factory
						var suppressedOccurrences = new List<ComponentOccurrence>();
						foreach (var occ in asmDoc.ComponentDefinition.Occurrences.Cast<ComponentOccurrence>()
						                          .Where(occ => occ.Excluded))
						{
							suppressedOccurrences.Add(occ);
							Debug.WriteLine($"Collected excluded occurrence: {occ.Name}");
						}

						Debug.WriteLine($"Collected {suppressedOccurrences.Count} excluded occurrences");

						// Delete the collected excluded occurrences to allow factory deletion
						Debug.WriteLine("Deleting collected excluded occurrences...");
						if (suppressedOccurrences.Count > 0)
							foreach (var occ in suppressedOccurrences)
								try
								{
									Debug.WriteLine($"Attempting to delete occurrence: {occ.Name}");
									occ.Delete();
									Debug.WriteLine($"Deleted excluded occurrence: {occ.Name}");
								}
								catch (Exception ex)
								{
									Debug.WriteLine($"Failed to delete excluded occurrence: {ex.Message}");
								}

						Debug.WriteLine("Deletion of excluded occurrences complete");

						// Proceed directly to factory deletion
						Debug.WriteLine("Attempting to delete iAssembly factory...");
						asmDoc.ComponentDefinition.iAssemblyFactory.Delete();
						Debug.WriteLine("iAssembly factory deleted");
					}
					catch (COMException ex)
					{
						Debug.WriteLine($"COM Exception deleting iAssembly factory: {ex.Message}");
						throw new InvalidOperationException($"Failed to delete iAssembly factory: {ex.Message}", ex);
					}
				else
					Debug.WriteLine("Document is not an iAssembly factory - nothing to delete");

				break;

			default:
				Debug.WriteLine($"Unsupported document type: {doc.GetType().Name}");
				break;
		}
	}

	private static string GetPartNumber(Document doc)
	{
		try
		{
			var designTrackingProps = doc.PropertySets["Design Tracking Properties"];
			var raw                 = designTrackingProps["Part Number"]?.Value?.ToString() ?? string.Empty;
			var sanitized           = SanitizeFileName(raw);
			return !string.IsNullOrWhiteSpace(sanitized)
				? sanitized
				: SanitizeFileName(IOPath.GetFileNameWithoutExtension(doc.FullFileName ?? string.Empty));
		}
		catch
		{
			return SanitizeFileName(IOPath.GetFileNameWithoutExtension(doc.FullFileName ?? string.Empty)) ?? "Member";
		}
	}

	private static string SanitizeFileName(string value)
	{
		if (string.IsNullOrWhiteSpace(value)) return string.Empty;
		value = IOPath.GetInvalidFileNameChars().Aggregate(value, (current, c) => current.Replace(c, '_'));
		return value.Trim();
	}
}