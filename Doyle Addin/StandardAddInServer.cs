#region

using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Doyle_Addin.My_Project;
using Doyle_Addin.Optional_Features;
using Doyle_Addin.Options;

#endregion

namespace Doyle_Addin;

/// <inheritdoc />
[ProgId("Test2.StandardAddInServer")]
[Guid("513b9d7e-103e-4569-8eb5-ab3929cd33ad")]
public class StandardAddInServer : ApplicationAddInServer
{
	private const string TabAnnotateId = "id_TabAnnotate";
	private const string printupdate = "printUpdate";
	private const string dxfupdate = "dxfUpdate";
	private const string geniusprops = "geniusProps";

	// Event handler delegates to ensure proper unsubscription
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _dxfUpdateHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _obsoleteButtonHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _optionsButtonHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _printUpdateHandler;
	private readonly UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler _uiEventsResetRibbonInterfaceHandler;

	// Instance field to store the Inventor application reference
	private Application _application;

	// New Genius Properties button definition and property

	private UserInterfaceEvents uiEvents;

	public StandardAddInServer()
	{
		_dxfUpdateHandler                    = _ => DXFUpdate_OnExecute();
		_printUpdateHandler                  = _ => PrintUpdate_OnExecute();
		_optionsButtonHandler                = _ => OptionsButton_OnExecute();
		_obsoleteButtonHandler               = _ => ObsoleteButton_OnExecute();
		_uiEventsResetRibbonInterfaceHandler = UiEvents_OnResetRibbonInterface;
	}

	private ButtonDefinition DxfUpdate
	{
		[MethodImpl(MethodImplOptions.Synchronized)]
		get;
		[MethodImpl(MethodImplOptions.Synchronized)]
		set
		{
			field?.OnExecute -= _dxfUpdateHandler;

			field            =  value;
			field?.OnExecute += _dxfUpdateHandler;
		}
	}

	private ButtonDefinition PrintUpdate
	{
		[MethodImpl(MethodImplOptions.Synchronized)]
		get;
		[MethodImpl(MethodImplOptions.Synchronized)]
		set
		{
			field?.OnExecute -= _printUpdateHandler;

			field            =  value;
			field?.OnExecute += _printUpdateHandler;
		}
	}

	private ButtonDefinition OptionsButton
	{
		[MethodImpl(MethodImplOptions.Synchronized)]
		get;
		[MethodImpl(MethodImplOptions.Synchronized)]
		set
		{
			field?.OnExecute -= _optionsButtonHandler;

			field            =  value;
			field?.OnExecute += _optionsButtonHandler;
		}
	}

	private ButtonDefinition ObsoleteButton
	{
		[MethodImpl(MethodImplOptions.Synchronized)]
		get;
		[MethodImpl(MethodImplOptions.Synchronized)]
		set
		{
			field?.OnExecute -= _obsoleteButtonHandler;

			field            =  value;
			field?.OnExecute += _obsoleteButtonHandler;
		}
	}

	[MethodImpl(MethodImplOptions.Synchronized)]
	private void SetUiEvents(UserInterfaceEvents value)
	{
		uiEvents?.OnResetRibbonInterface -= _uiEventsResetRibbonInterfaceHandler;

		uiEvents                         =  value;
		uiEvents?.OnResetRibbonInterface += _uiEventsResetRibbonInterfaceHandler;
	}

	/// <summary>
	///     Gets the add-in installation directory dynamically
	/// </summary>
	/// <returns>The directory where the add-in is installed</returns>
	private static string GetAddInDirectory()
	{
		var assemblyLocation = Assembly.GetExecutingAssembly().Location;
		var directory        = Path.GetDirectoryName(assemblyLocation);
		return directory ?? throw new InvalidOperationException("Could not determine add-in directory");
	}

	#region ApplicationAddInServer Members

	// Inventor calls this method when it loads the AddIn. The AddInSiteObject provides access 
	// To the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
	// The first time. However, with the introduction of the ribbon, this argument is always true.
	/// <inheritdoc />
	public void Activate(ApplicationAddInSite AddInSiteObject, bool FirstTime)
	{
		try
		{
			// Initialize AddIn members.
			_application = AddInSiteObject.Application;

			// Initialize the static ThisApplication field for global access
			Initialize(_application);

			_ = CheckForUpdateAndDownloadAsync(_application).ConfigureAwait(false);

			// Get a reference to the ControlDefinitions object. 
			var controlDefs   = _application.CommandManager.ControlDefinitions;
			var oThemeManager = _application.ThemeManager;

			var oTheme = oThemeManager.ActiveTheme;

			switch (oTheme.Name ?? "")
			{
				case "LightTheme":
				case "DarkTheme":
				{
					var themeSuffix = oTheme.Name == "LightTheme" ? "Light" : "Dark";
					var iconSizes   = new[] { 16, 32 };
					var icons = new[]
					{
						new
						{
							Name         = "PrintUpdate", Icon = "Doyle_Addin.Resources.PrintUpdateIcon.svg",
							InternalName = printupdate
						},
						new
						{
							Name         = "DXFUpdate", Icon = "Doyle_Addin.Resources.DXFUpdateIcon.svg",
							InternalName = dxfupdate
						},
						new
						{
							Name         = "Settings", Icon = "Doyle_Addin.Resources.SettingsIcon.svg",
							InternalName = "userOptions"
						},
						new
						{
							Name         = obsoleteprint, Icon = "Doyle_Addin.Resources.ObsoletePrint.svg",
							InternalName = obsoleteprint
						},
						new
						{
							Name         = "Genius Properties", Icon = "Doyle_Addin.Resources.SettingsIcon.svg",
							InternalName = geniusprops
						}
					};

					foreach (var icon in icons)
					{
						var largeIcon =
							SVGconvert.SvgResourceToPictureDisp(icon.Icon, iconSizes[1], iconSizes[1],
								themeSuffix);
						var smallIcon =
							SVGconvert.SvgResourceToPictureDisp(icon.Icon, iconSizes[0], iconSizes[0],
								themeSuffix);

						// Try to remove the existing definition first if it exists
						try
						{
							var existingDef = controlDefs[icon.InternalName];
							existingDef?.Delete();
						}
						catch
						{
							// Definition doesn't exist, which is fine
						}

						switch (icon.Name ?? "")
						{
							case "PrintUpdate":
							{
								PrintUpdate = controlDefs.AddButtonDefinition("Print" + '\n' + "Update",
									printupdate, CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case "DXFUpdate":
							{
								DxfUpdate = controlDefs.AddButtonDefinition("DXF" + '\n' + "Update", dxfupdate,
									CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case "Settings":
							{
								OptionsButton = controlDefs.AddButtonDefinition("Options", "userOptions",
									CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case obsoleteprint:
							{
								ObsoleteButton = controlDefs.AddButtonDefinition("Obsolete" + '\n' + "Print",
									obsoleteprint, CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
						}
					}

					break;
				}
			}

			// Always add to the user interface (not just the first time)
			// This ensures buttons appear when add-in is reloaded
			AddToUserInterface();

			// Connect to the user-interface events to handle a ribbon reset.
			SetUiEvents(_application.UserInterfaceManager.UserInterfaceEvents);

			// Ensure the option file exists with default values if it doesn't exist
			if (File.Exists(UserOptions.OptionsFilePath)) return;
			var defaultOptions = new UserOptions
			{
				PrintExportLocation = @"P:\",
				DxfExportLocation   = @"X:\"
			};
			defaultOptions.Save();
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.Message);
		}
	}

	private static async Task CheckForUpdateAndDownloadAsync(Application application)
	{
		try
		{
			var localVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
			// System.Windows.Forms.MessageBox.Show($"Local version: {localVersion}", "Debug")

			var releaseNullable = await GetLatestReleaseFromGitHub();
			if (!releaseNullable.HasValue)
				// System.Windows.Forms.MessageBox.Show("Could not fetch release info from GitHub.", "Debug")
				return;

			var release = releaseNullable.Value;

			var latestVersion = release.GetProperty("tag_name").GetString()?.TrimStart('v');
			// System.Windows.Forms.MessageBox.Show($"Latest GitHub version: {latestVersion}", "Debug")

			if (localVersion != null)
			{
				var localVerObj = new Version(localVersion);
				if (latestVersion != null)
				{
					var latestVerObj = new Version(latestVersion);
					if (latestVerObj <= localVerObj) return;
				}
			}

			var result =
				MessageBox.Show(
					$"A new version of the Doyle AddIn is available ({latestVersion}) . Update now?",
					"Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			if (result == DialogResult.Yes)
			{
				await File.WriteAllTextAsync(
					Path.Combine(GetAddInDirectory(), "pending_update.txt"), "update");
				application.Quit();
			}
			else
			{
				await File.WriteAllTextAsync(
					Path.Combine(GetAddInDirectory(), "pending_update.txt"), "update");
				MessageBox.Show("The update will be installed after you close Inventor.", "Update Scheduled",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			// System.Windows.Forms.MessageBox.Show("You are running the latest version.", "Debug")
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
			// Optionally log or show error
		}
	}

	private static async Task<JsonElement?> GetLatestReleaseFromGitHub()
	{
		const string githubUser = "LeviErskine";
		const string githubRepo = "Doyle-Addin-C-";
		const string url        = $"https://api.github.com/repos/{githubUser}/{githubRepo}/releases";
		using var    client     = new HttpClient();
		client.DefaultRequestHeaders.UserAgent.ParseAdd("InventorAddinUpdater");
		var json = await client.GetStringAsync(url);
		var doc  = JsonDocument.Parse(json);
		var root = doc.RootElement;
		if (root.ValueKind == JsonValueKind.Array &&
		    root.GetArrayLength() > 0) return root[0]; // Use the first release (most recent)

		return null;
	}

	// Inventor calls this method when the AddIn is unloaded. The AddIn will be
	// unloaded either manually by the user or when the Inventor session is terminated.
	/// <inheritdoc />
	public void Deactivate()
	{
		// Clean up button definitions
		try
		{
			PrintUpdate?.Delete();
			PrintUpdate = null;
		}
		catch
		{
			// ignored
		}

		try
		{
			DxfUpdate?.Delete();
			DxfUpdate = null;
		}
		catch
		{
			// ignored
		}

		try
		{
			OptionsButton?.Delete();
			OptionsButton = null;
		}
		catch
		{
			// ignored
		}

		try
		{
			ObsoleteButton?.Delete();
			ObsoleteButton = null;
		}
		catch
		{
			// ignored
		}

		// Release objects.
		SetUiEvents(null);
		_application = null;

		// Check for pending update marker
		var updateMarker = Path.Combine(GetAddInDirectory(), "pending_update.txt");
		var updaterBat   = Path.Combine(GetAddInDirectory(), "Updater.bat");
		if (!File.Exists(updateMarker) || !File.Exists(updaterBat)) return;
		try
		{
			// Start Updater.bat in a detached process
			var psi = new ProcessStartInfo
			{
				FileName        = updaterBat,
				WindowStyle     = ProcessWindowStyle.Normal,
				UseShellExecute = true
			};
			Process.Start(psi);
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
			// Optionally log or show error
		}

		// Remove marker
		File.Delete(updateMarker);
	}

	// This property is provided to allow the AddIn to expose an API of its own to other 
	// Programs. Typically, this would be done by implementing the AddIn's API
	// interface in a class and returning that class object through this property.
	/// <inheritdoc />
	public object Automation => null;

	private static readonly string[] Item4 = ["id_PanelP_SheetMetalManageUnfold"];
	private static readonly string[] Item4Array = ["id_PanelP_ToolsOptions", addIns];
	private static readonly string[] Item4Array0 = [addIns];
	private static readonly string[] Item4Array2 = ["id_PanelP_ToolsOptions"];
	private static readonly string[] Item4Array3 = ["id_PanelP_FlatPatternExit"];
	private static readonly string[] Item4Array4 = ["id_PanelD_AnnotateRevision"];
	private const string drawing = "Drawing";
	private const string obsoleteprint = "ObsoletePrint";
	private const string addIns = "Add-Ins";

	/// <inheritdoc />
	public void ExecuteCommand(int CommandID)
	{
	}

	#endregion

	#region User interface definition

	// Sub where the user-interface creation is done.  This is called when
	// the add-in is loaded and also if the user interface is reset.
	private void AddToUserInterface()
	{
		// Load user options to check feature flags
		var options = UserOptions.Load();

		// Cache frequently used objects
		var uiManager = _application.UserInterfaceManager;

		// Define ribbon mappings for each document type
		var ribbonMappings = new Dictionary<string, Ribbon>
		{
			{ "Part", uiManager.Ribbons["Part"] }, { "Assembly", uiManager.Ribbons["Assembly"] },
			{ drawing, uiManager.Ribbons[drawing] }, { "ZeroDoc", uiManager.Ribbons["ZeroDoc"] }
		};

		// Define button configurations with document type specificity
		// DXF button only appears on Part documents
		var dxfButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>
		{
			Tuple.Create("id_TabSheetMetal", dxfupdate, DxfUpdate, Item4),
			Tuple.Create("id_TabFlatPattern", dxfupdate, DxfUpdate, Item4Array3),
			Tuple.Create("id_TabTools", dxfupdate, DxfUpdate, Item4Array2)
		};

		// Option button appears on all document types
		var optionsButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>
		{
			Tuple.Create("id_TabPlaceViews", printupdate, PrintUpdate, Item4Array0),
			Tuple.Create(TabAnnotateId, printupdate, PrintUpdate, Item4Array0),
			Tuple.Create("id_TabTools", "userOptions", OptionsButton, Item4Array)
		};

		// Obsolete Print button - only add if feature is enabled
		var obsoletePrintConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>();
		if (options.EnableObsoletePrint)
			obsoletePrintConfigs.Add(Tuple.Create(TabAnnotateId, obsoleteprint, ObsoleteButton, Item4Array4));

		// Add buttons to appropriate ribbons based on document context
		foreach (var (ribbonName, ribbon) in ribbonMappings)
		{
			// Add the DXF button only to Part ribbon
			if (ribbonName == "Part")
				foreach (var (tabName, panelName, buttonDef, fallbackPanels) in dxfButtonConfigs)
					AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);

			// Add Options buttons to all ribbons
			foreach (var (tabName, panelName, buttonDef, fallbackPanels) in optionsButtonConfigs)
				AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);

			AddObsoletePrintButton(ribbon, ribbonName, obsoletePrintConfigs);
		}
	}

	// Helper method to add obsolete print button only to Drawing ribbon
	private static void AddObsoletePrintButton(Ribbon ribbon, string ribbonName,
		List<Tuple<string, string, ButtonDefinition, string[]>> obsoletePrintConfigs)
	{
		if (ribbonName != drawing) return;

		foreach (var (tabName, panelName, buttonDef, fallbackPanels) in obsoletePrintConfigs)
			AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);
	}

	// Helper method to add a button to specific ribbon with fallback handling
	private static void AddButtonToRibbon(Ribbon ribbon, string tabName, string panelName,
		ButtonDefinition buttonDef, string[] fallbackPanels)
	{
		try
		{
			var tab = ribbon.RibbonTabs[tabName];
			if (tab is null) return;

			var panel = GetOrCreatePanel(tab, panelName, fallbackPanels);
			if (panel is null) return;

			if (!ButtonExistsInPanel(panel, buttonDef))
				panel.CommandControls.AddButton(buttonDef, true);
		}
		catch (Exception ex)
		{
			Debug.Print(
				$"Failed to add button '{(buttonDef is not null ? buttonDef.DisplayName : "Unknown")}' to tab '{tabName}' in ribbon '{ribbon.InternalName}': {ex.Message}");
		}
	}

	private static RibbonPanel GetOrCreatePanel(RibbonTab tab, string panelName, string[] fallbackPanels)
	{
		try
		{
			var panel = tab.RibbonPanels[panelName];
			if (panel is not null) return panel;
		}
		catch
		{
			// Continue to fallback logic
		}

		return TryFallbackPanels(tab, panelName, fallbackPanels);
	}

	private static RibbonPanel TryFallbackPanels(RibbonTab tab, string panelName, string[] fallbackPanels)
	{
		return fallbackPanels.Select(fallbackPanel => TryGetFallbackPanel(tab, fallbackPanel, panelName))
		                     .FirstOrDefault(panel => panel is not null);
	}

	private static RibbonPanel TryGetFallbackPanel(RibbonTab tab, string fallbackPanel, string panelName)
	{
		try
		{
			return fallbackPanel == addIns
				? tab.RibbonPanels.Add(addIns, panelName, Globals.AddInClientId())
				: tab.RibbonPanels[fallbackPanel];
		}
		catch
		{
			return null;
		}
	}

	private static bool ButtonExistsInPanel(RibbonPanel panel, ButtonDefinition buttonDef)
	{
		foreach (CommandControl ctrl in panel.CommandControls)
			try
			{
				if (ctrl?.ControlDefinition == null || buttonDef is null) continue;
				if ((ctrl.ControlDefinition.InternalName ?? "") == (buttonDef.InternalName ?? ""))
					return true;
			}
			catch
			{
				// Skip this control if there's any issue
			}

		return false;
	}

	private void UiEvents_OnResetRibbonInterface(NameValueMap context)
	{
		AddToUserInterface();
	}

	private static void DXFUpdate_OnExecute()
	{
		DXFs.DxfUpdate.RunDxfUpdate();
	}

	private static void PrintUpdate_OnExecute()
	{
		Prints.PrintUpdate.RunPrintUpdate();
	}

	private void OptionsButton_OnExecute()
	{
		var optionsForm = new UserOptionsWindow();
		var result      = optionsForm.ShowDialog();

		// Refresh the ribbon after options are saved
		if (result is true) RefreshRibbon();
	}

	private void ObsoleteButton_OnExecute()
	{
		new Action(() => ObsoletePrint.ApplyObsoletePrint(_application))();
	}

	// Helper method to refresh the ribbon UI
	private void RefreshRibbon()
	{
		try
		{
			// Remove the obsolete print button from all ribbons first
			RemoveButtons(_application);

			// Re-add buttons with updated settings
			AddToUserInterface();
		}

		// MessageBox.Show("Ribbon updated successfully. Changes are now active.",
		// "Settings Applied",
		// MessageBoxButtons.OK,
		// MessageBoxIcon.Information)'
		catch (Exception ex)
		{
			Console.WriteLine(ex);
			// MessageBox.Show($"Error refreshing ribbon: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
		}
	}

	// Helper method to remove the obsolete print button from ribbons
	private static void RemoveButtons(Application application)
	{
		try
		{
			var ribbon = application.UserInterfaceManager.Ribbons[drawing];
			if (ribbon is null) return;

			RemoveButtonsFromAnnotateTab(ribbon);
		}
		catch (Exception ex)
		{
			Debug.Print($"Error removing obsolete print button: {ex.Message}");
		}
	}

	private static void RemoveButtonsFromAnnotateTab(Ribbon ribbon)
	{
		try
		{
			var tab = ribbon.RibbonTabs[TabAnnotateId];
			if (tab is null) return;

			ProcessPanelsForObsoleteButtons(tab.RibbonPanels);
		}
		catch (Exception ex)
		{
			Debug.Print($"Could not remove button from Annotate tab: {ex.Message}");
		}
	}

	private static void ProcessPanelsForObsoleteButtons(RibbonPanels panels)
	{
		foreach (var controlsToRemove in from RibbonPanel panel in panels select GetControlsToRemove(panel))
			RemoveObsoleteControls(controlsToRemove);
	}

	private static List<CommandControl> GetControlsToRemove(RibbonPanel panel)
	{
		return panel.CommandControls.Cast<CommandControl>().Where(IsObsoletePrintControl).ToList();
	}

	private static bool IsObsoletePrintControl(CommandControl ctrl)
	{
		try
		{
			return ctrl?.ControlDefinition is { InternalName: obsoleteprint };
		}
		catch
		{
			return false;
		}
	}

	private static void RemoveObsoleteControls(IEnumerable<CommandControl> controls)
	{
		foreach (var ctrl in controls) TryDeleteControl(ctrl);
	}

	private static void TryDeleteControl(CommandControl ctrl)
	{
		try
		{
			ctrl.Delete();
		}
		catch (Exception ex)
		{
			Debug.Print($"Failed to delete control: {ex.Message}");
		}
	}

	#endregion
}

/// <summary>
/// </summary>
public static class Globals
{
	/// <summary>
	///     This function uses reflection to get the GuidAttribute associated with the add-in.
	/// </summary>
	/// <returns></returns>
	public static string AddInClientId()
	{
		var guid = "";
		try
		{
			var t                = typeof(StandardAddInServer);
			var customAttributes = t.GetCustomAttributes(typeof(GuidAttribute), false);
			var guidAttribute    = (GuidAttribute)customAttributes[0];
			guid = "{" + guidAttribute.Value + "}";
		}
		catch
		{
			// ignored
		}

		return guid;
	}
}