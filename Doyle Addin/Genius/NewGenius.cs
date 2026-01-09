#region

using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using Brushes = System.Windows.Media.Brushes;
using MessageBox = System.Windows.Forms.MessageBox;
using MessageBoxButtons = System.Windows.Forms.MessageBoxButtons;
using MessageBoxIcon = System.Windows.Forms.MessageBoxIcon;
using FileAttributes = System.IO.FileAttributes;

#endregion

namespace Doyle_Addin.Genius;

public sealed class NewGenius : Window, IDisposable
{
	private const string Clid = "{A7F9C2B1-8D4E-4F5A-9B3C-6E7D8F9A0B1C}";

	// Timer to monitor visibility state
	private readonly DispatcherTimer visibilityTimer;

	// Inventor Application Reference

	private bool disposed;

	// Dockable Window Reference
	private DockableWindow docWin;

    /// <summary>
    ///     Private constructor. Use ShowGeniusPanel to instantiate.
    /// </summary>
    private NewGenius(Application mInventorApp, string internalName, string windowTitle, UIElement wpfControl,
		bool showTitle = true)
	{
		InitializeComponent();

		// WPF Window Settings
		Content = wpfControl;

		// Ensure handle is created before adding to Inventor
		var helper = new WindowInteropHelper(this);
		helper.EnsureHandle();
		var windowHandle = helper.Handle;

		// Create or Refresh the Inventor DockableWindow
		SetupDockableWindow(mInventorApp, internalName, windowTitle, windowHandle, showTitle);

		// Visibility Timer: Checks if the user closed the dockable window via the "X"
		visibilityTimer      =  new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
		visibilityTimer.Tick += OnVisibilityTimerTick;
		visibilityTimer.Start();

		// Event Subscription
		Closed += NewGenius_Closed;

		// Ensure visible immediately
		Show();
	}

	public void Dispose()
	{
		if (disposed) return;
		disposed = true;

		// 1. Stop Timers
		if (visibilityTimer != null)
		{
			visibilityTimer.Stop();
			visibilityTimer.Tick -= OnVisibilityTimerTick;
		}

		// 2. Unsubscribe Events
		Closed -= NewGenius_Closed;

		// 3. Clean up Content
		if (Content is IDisposable disposableContent)
			try
			{
				disposableContent.Dispose();
			}
			catch
			{
				/* Ignore */
			}

		Content = null; // Detach WPF visual tree

		// 4. Clean up Inventor Dockable Window (COM)
		if (docWin != null)
			try
			{
				// Ensure we don't crash if Inventor is already closing
				docWin.Visible = false;
				docWin.Delete();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[Genius] Error deleting dockable window: {ex.Message}");
			}
			finally
			{
				// Release COM reference
				Marshal.ReleaseComObject(docWin);
				docWin = null;
			}

		// 5. Close this WPF host
		try
		{
			if (IsVisible) Hide();
			Close();
		}
		catch
		{
			// Window might already be closed
		}

		GC.SuppressFinalize(this);
	}

	private void SetupDockableWindow(Application minventorApp, string internalName, string windowTitle,
		IntPtr childHandle,
		bool showTitle)
	{
		try
		{
			// Attempt to clean up existing window with this CLID
			var existingWindow = minventorApp.UserInterfaceManager.DockableWindows
			                                 .Cast<DockableWindow>()
			                                 .FirstOrDefault(w => w.ClientId == Clid);

			existingWindow?.Delete();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[Genius] Warning cleaning up old window: {ex.Message}");
		}

		try
		{
			docWin = minventorApp.UserInterfaceManager.DockableWindows.Add(Clid, internalName, windowTitle);
			docWin.AddChild(childHandle);
			docWin.ShowTitleBar = showTitle;
			docWin.Visible      = true;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[Genius] Failed to create DockableWindow: {ex.Message}");
			// Fallback: Just show the WPF window as a standalone tool window
			Show();
		}
	}

	public static async Task ShowGeniusPanel(Application minventorApp)
	{
		// 1. Validate Document State
		if (!ValidateDocumentState(minventorApp)) return;

		// 2. Ensure we run on the UI Thread (WPF Requirement)
		if (System.Windows.Application.Current != null && !System.Windows.Application.Current.Dispatcher.CheckAccess())
		{
			await System.Windows.Application.Current.Dispatcher.InvokeAsync(() => ShowGeniusPanel(minventorApp));
			return;
		}

		try
		{
			var doc             = minventorApp.ActiveDocument;
			var databaseService = new DatabaseService();

			// 3. Determine Context (Part vs Assembly vs iPart/iAssembly)
			var documentType = DetermineDocumentType(doc);

			UIElement content;
			string    title;

			switch (documentType)
			{
				case DocumentType.Factory:
					// Use iPartAssemblyPanel for factory documents (iParts and iAssemblies)
					var iPartAssemblyPanel = new iPartAssemblyPanel(minventorApp, null, null, false);
					await iPartAssemblyPanel.InitializePanel(); // Initialize synchronously where possible

					content = ExtractContentFromWindow(iPartAssemblyPanel);
					title   = "iPart/Assembly Panel";
					break;

				case DocumentType.RegularAssembly:
					// Use AssemblyPanel for regular assemblies with component parts
					var (calculatedPartInfo, hasCalculatedValues, componentParts) =
						await DocumentHandles.HandleAssembly(
							doc as AssemblyDocument, minventorApp, null, new Dictionary<string, string>());

					var assemblyPanel = new AssemblyPanel(minventorApp, calculatedPartInfo, null, hasCalculatedValues,
						componentParts);
					await assemblyPanel.InitializePanel();

					content = ExtractContentFromWindow(assemblyPanel);
					title   = "Assembly Panel";
					break;

				default:
					// Create the NewGenius wrapper first so we can pass it to the panel
					var newGeniusWindow = new NewGenius(minventorApp, "GeniusWindow", "Genius Panel", null);

					// Create Panel with parent window reference
					var geniusPanel = new GeniusPanel(minventorApp, databaseService, newGeniusWindow);

					content = ExtractContentFromWindow(geniusPanel);
					title   = "Genius Panel";

					// Update the window content with the extracted panel content
					newGeniusWindow.Content = content;
					break;
			}

			// 4. Create and Host the Window (for iPart/Assembly and Assembly cases)
			if (documentType is DocumentType.Factory or DocumentType.RegularAssembly)
			{
				var _ = new NewGenius(minventorApp, "GeniusWindow", title, content);
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error initializing panel: {ex.Message}", "Genius Error", MessageBoxButtons.OK,
				MessageBoxIcon.Error);
		}
	}

	private static bool ValidateDocumentState(Application minventorApp)
	{
		try
		{
			var activeDoc = minventorApp.ActiveDocument;
			if (activeDoc == null)
			{
				MessageBox.Show("No active document found.", "No Active Document", MessageBoxButtons.OK,
					MessageBoxIcon.Warning);
				return false;
			}

			if (string.IsNullOrEmpty(activeDoc.FullFileName)) return true;
			var fileInfo = new FileInfo(activeDoc.FullFileName);
			if (!fileInfo.Exists || (fileInfo.Attributes & FileAttributes.ReadOnly) != FileAttributes.ReadOnly)
				return true;
			MessageBox.Show("The document appears to be read-only or checked in.", "Document Access Error",
				MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return false;
		}
		catch
		{
			MessageBox.Show("Unable to access document properties.", "Document Access Error", MessageBoxButtons.OK,
				MessageBoxIcon.Warning);
			return false;
		}
	}

	private static DocumentType DetermineDocumentType(Document doc)
	{
		if (IsFactoryDocument(doc)) return DocumentType.Factory;

		return doc.DocumentType == kAssemblyDocumentObject ? DocumentType.RegularAssembly : DocumentType.Other;
	}

	private static bool IsFactoryDocument(Document doc)
	{
		return doc.DocumentType switch
		{
			kPartDocumentObject     => (doc as PartDocument)!.ComponentDefinition.IsiPartFactory,
			kAssemblyDocumentObject => (doc as AssemblyDocument)!.ComponentDefinition.IsiAssemblyFactory,
			_                       => false
		};
	}

    /// <summary>
    ///     Extracts the Content of a Window, detaches it, and returns it as a UIElement.
    ///     Replaces the unwieldy UnWrapWindow methods.
    /// </summary>
    private static UIElement ExtractContentFromWindow(Window sourceWindow)
	{
		if (sourceWindow == null) return null;

		// Get content
		var content = sourceWindow.Content as UIElement;

		// Detach content. 
		// Setting Content to null in WPF automatically removes the logical parent 
		// and visual parent links. Reflection is rarely needed here.
		sourceWindow.Content = null;

		// Close the source window as it is now an empty shell
		sourceWindow.Close();

		return content;
	}

	private void InitializeComponent()
	{
		WindowStyle   = WindowStyle.None;
		ResizeMode    = ResizeMode.NoResize;
		Background    = Brushes.White;
		ShowInTaskbar = false;
		Width         = 300; // Default width if docked fails
		Height        = 600;
	}

	public void CloseDockableWindow()
	{
		// Trigger cleanup
		Dispose();
	}

	private void OnVisibilityTimerTick(object sender, EventArgs e)
	{
		if (disposed) return;

		try
		{
			if (docWin == null) return;

			// If the user clicked the 'X' on the Inventor Dockable window, Visible becomes false
			if (docWin.Visible) return;
			visibilityTimer.Stop();
			Dispose();
		}
		catch (COMException)
		{
			// The underlying COM object might be gone
			visibilityTimer.Stop();
			Dispose();
		}
	}

	private void NewGenius_Closed(object sender, EventArgs e)
	{
		Dispose();
	}

	public void ForceDispose()
	{
		Dispose();
	}

	private enum DocumentType
	{
		Factory,
		RegularAssembly,
		Other
	}
}