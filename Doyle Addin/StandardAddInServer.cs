namespace DoyleAddin;

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Genius;
using Inventor;
using My_Project;
using Optional_Features;
using Options;
using File = File;
using Path = Path;

/// <inheritdoc />
[ProgId("DoyleAddin.StandardAddInServer")]
[Guid("513b9d7e-103e-4569-8eb5-ab3929cd33ad")]
public class StandardAddInServer : ApplicationAddInServer
{
	private static readonly string[] SheetMetalManageUnfold = ["id_PanelP_SheetMetalManageUnfold"];
	private static readonly string[] ToolsOptions = ["id_PanelP_ToolsOptions", "id_PanelA_ToolsOptions"];
	private static readonly string[] AddinTab = ["Add-Ins"];
	private static readonly string[] FlatPatternExit = ["id_PanelP_FlatPatternExit"];

	private static readonly string[] AnnotateRevision = ["id_PanelD_AnnotateRevision"];
	private static readonly string[] ManagePanels = ["id_PanelP_Manage", "id_PanelA_Manage"];

	private static PanelWrapper _geniusPanelWrapper;

	// Event handler delegates to ensure proper unsubscription
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _dxfUpdateHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _explodeiComponentsHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _geniusPanelHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _obsoleteButtonHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _optionsButtonHandler;
	private readonly ButtonDefinitionSink_OnExecuteEventHandler _printUpdateHandler;
	private readonly UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler _uiEventsResetRibbonInterfaceHandler;

	private UserInterfaceEvents uiEvents;

	/// <summary>
	///     Represents the primary class for the add-in, implementing the core application add-in lifecycle
	///     and providing custom functionality to extend Inventor's behavior.
	/// </summary>
	/// <remarks>
	///     This class implements the `ApplicationAddInServer` interface, handling the activation,
	///     deactivation, and execution of the add-in functionality within the Inventor environment.
	///     It includes event handling for user interface updates and button commands.
	/// </remarks>
	public StandardAddInServer()
	{
		_dxfUpdateHandler                    = _ => DXFUpdate_OnExecute();
		_printUpdateHandler                  = _ => PrintUpdate_OnExecute();
		_optionsButtonHandler                = _ => OptionsButton_OnExecute();
		_obsoleteButtonHandler               = _ => ObsoleteButton_OnExecute();
		_explodeiComponentsHandler           = _ => ExplodeiComponents_OnExecute();
		_geniusPanelHandler                  = _ => GeniusPanel_OnExecute();
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

	private ButtonDefinition ExplodeiComponentsButton
	{
		[MethodImpl(MethodImplOptions.Synchronized)]
		get;
		[MethodImpl(MethodImplOptions.Synchronized)]
		set
		{
			field?.OnExecute -= _explodeiComponentsHandler;

			field            =  value;
			field?.OnExecute += _explodeiComponentsHandler;
		}
	}

	private ButtonDefinition GeniusPanelButton
	{
		[MethodImpl(MethodImplOptions.Synchronized)]
		get;
		[MethodImpl(MethodImplOptions.Synchronized)]
		set
		{
			field?.OnExecute -= _geniusPanelHandler;

			field            =  value;
			field?.OnExecute += _geniusPanelHandler;
		}
	}

	// Inventor calls this method when it loads the AddIn. The AddInSiteObject provides access 
	// To the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
	// The first time. However, with the introduction of the ribbon, this argument is always true.
	/// <inheritdoc />
	public void Activate(ApplicationAddInSite AddInSiteObject, bool FirstTime)
	{
		try
		{
			Initialize(AddInSiteObject.Application);
			_ = CheckForUpdateAndDownloadAsync().ConfigureAwait(false);

			// Get a reference to the ControlDefinitions object. 
			var controlDefs   = ThisApplication.CommandManager.ControlDefinitions;
			var oThemeManager = ThisApplication.ThemeManager;

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
							Name         = "PrintUpdate", Icon = "DoyleAddin.Resources.PrintUpdateIcon.svg",
							InternalName = "PrintUpdate"
						},
						new
						{
							Name         = "DXFUpdate", Icon = "DoyleAddin.Resources.DXFUpdateIcon.svg",
							InternalName = "DXFUpdate"
						},
						new
						{
							Name         = "Settings", Icon = "DoyleAddin.Resources.SettingsIcon.svg",
							InternalName = "Settings"
						},
						new
						{
							Name         = "ObsoletePrint", Icon = "DoyleAddin.Resources.ObsoletePrint.svg",
							InternalName = "ObsoletePrint"
						},
						new
						{
							Name         = "Explode iComponents", Icon = "DoyleAddin.Resources.ExplodeiPart.svg",
							InternalName = "Explode iComponents"
						},
						new
						{
							Name         = "Genius Panel", Icon = "DoyleAddin.Resources.GeniusPanel.svg",
							InternalName = "Genius Panel"
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
									"PrintUpdate", CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case "DXFUpdate":
							{
								DxfUpdate = controlDefs.AddButtonDefinition("DXF" + '\n' + "Update", "DXFUpdate",
									CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case "Settings":
							{
								OptionsButton = controlDefs.AddButtonDefinition("Options", "Settings",
									CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case "ObsoletePrint":
							{
								ObsoleteButton = controlDefs.AddButtonDefinition("Obsolete" + '\n' + "Print",
									"ObsoletePrint", CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case "Explode iComponents":
							{
								ExplodeiComponentsButton = controlDefs.AddButtonDefinition(
									"Explode" + '\n' + "iComponents",
									"Explode iComponents", CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(),
									StandardIcon: smallIcon, LargeIcon: largeIcon);
								break;
							}
							case "Genius Panel":
							{
								GeniusPanelButton = controlDefs.AddButtonDefinition(
									"Genius" + '\n' + "Panel",
									"Genius Panel", CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(),
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
			SetUiEvents(ThisApplication.UserInterfaceManager.UserInterfaceEvents);

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

		try
		{
			ExplodeiComponentsButton?.Delete();
			ExplodeiComponentsButton = null;
		}
		catch
		{
			// ignored
		}

		try
		{
			GeniusPanelButton?.Delete();
			GeniusPanelButton = null;
		}
		catch
		{
			// ignored
		}

		// Release objects.
		SetUiEvents(null);

		// Clean up the genius panel wrapper
		try
		{
			_geniusPanelWrapper?.Close();
			_geniusPanelWrapper = null;
		}
		catch
		{
			// ignored
		}

		ThisApplication = null;

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

	/// <inheritdoc />
	public void ExecuteCommand(int CommandID)
	{
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

	private static async Task CheckForUpdateAndDownloadAsync()
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
				ThisApplication.Quit();
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

	// Sub where the user-interface creation is done.  This is called when
	// the add-in is loaded and also if the user interface is reset.
	private void AddToUserInterface()
	{
		// Load user options to check feature flags
		var options = UserOptions.Load();

		// Cache frequently used objects
		var uiManager = ThisApplication.UserInterfaceManager;

		// Define ribbon mappings for each document type
		var ribbonMappings = new Dictionary<string, Ribbon>
		{
			{ "Part", uiManager.Ribbons["Part"] }, { "Assembly", uiManager.Ribbons["Assembly"] },
			{ "Drawing", uiManager.Ribbons["Drawing"] }, { "ZeroDoc", uiManager.Ribbons["ZeroDoc"] }
		};

		// Define button configurations with document type specificity
		// DXF button only appears on Part documents
		var dxfButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>
		{
			Tuple.Create("id_TabSheetMetal", "DXFUpdate", DxfUpdate, SheetMetalManageUnfold),
			Tuple.Create("id_TabFlatPattern", "DXFUpdate", DxfUpdate, FlatPatternExit),
			Tuple.Create("id_TabTools", "DXFUpdate", DxfUpdate, ToolsOptions)
		};

		// Option button appears on all document types
		var optionsButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>
		{
			Tuple.Create("id_TabTools", "Settings", OptionsButton, ToolsOptions)
		};

		// Print Update button appears on Drawings document types
		var printButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>
		{
			Tuple.Create("id_TabPlaceViews", "PrintUpdate", PrintUpdate, AddinTab),
			Tuple.Create("id_TabAnnotate", "PrintUpdate", PrintUpdate, AddinTab)
		};

		// ExplodeiComponents button appears on Part documents
		var explodeButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>
		{
			Tuple.Create("id_TabManage", "Explode iComponents", ExplodeiComponentsButton, ManagePanels)
		};

		// Genius Panel button appears on all document types
		var geniusPanelConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>
		{
			Tuple.Create("id_TabTools", "Genius Panel", GeniusPanelButton, ToolsOptions)
		};

		// Obsolete Print button - only add if feature is enabled
		var obsoletePrintConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>();
		if (options.EnableObsoletePrint)
			obsoletePrintConfigs.Add(Tuple.Create("id_TabAnnotate", "ObsoletePrint", ObsoleteButton, AnnotateRevision));

		// Add buttons to appropriate ribbons based on document context
		foreach (var (ribbonName, ribbon) in ribbonMappings)
		{
			AddButtonsToRibbon(ribbon, ribbonName, "Part", dxfButtonConfigs);
			AddButtonsToRibbon(ribbon, ribbonName, "Assembly", geniusPanelConfigs);
			AddButtonsToRibbon(ribbon, ribbonName, "Part", geniusPanelConfigs);
			AddButtonsToRibbon(ribbon, ribbonName, "Drawing", printButtonConfigs);
			AddButtonsToRibbon(ribbon, ribbonName, null, optionsButtonConfigs);

			if (options.EnableExplodeiComponents)
				AddButtonsToRibbon(ribbon, ribbonName, "Part", explodeButtonConfigs);
			AddButtonsToRibbon(ribbon, ribbonName, "Assembly", explodeButtonConfigs);

			if (options.EnableObsoletePrint)
				AddButtonsToRibbon(ribbon, ribbonName, "Drawing", obsoletePrintConfigs);
		}
	}

	// Generic helper method to add buttons to ribbon based on document type filter
	private static void AddButtonsToRibbon(Ribbon ribbon, string ribbonName, object ribbonFilter,
		List<Tuple<string, string, ButtonDefinition, string[]>> buttonConfigs)
	{
		if (buttonConfigs.Count == 0) return;

		var shouldAdd = ribbonFilter switch
		{
			null                     => true,
			string singleRibbon      => ribbonName == singleRibbon,
			string[] multipleRibbons => multipleRibbons.Contains(ribbonName),
			_                        => false
		};

		if (!shouldAdd) return;

		foreach (var (tabName, panelName, buttonDef, fallbackPanels) in buttonConfigs)
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
			return fallbackPanel == "Add-Ins"
				? tab.RibbonPanels.Add("Add-Ins", panelName, Globals.AddInClientId())
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

	private static void ObsoleteButton_OnExecute()
	{
		ObsoletePrint.ApplyObsoletePrint();
	}

	private static void ExplodeiComponents_OnExecute()
	{
		ExplodeiComponents.ExplodeiComponentsAction();
	}

	private static void GeniusPanel_OnExecute()
	{
		try
		{
			if (ThisApplication == null) return;

			// Clean up existing wrapper if it exists and is disposed
			if (_geniusPanelWrapper is { IsDisposed: true }) _geniusPanelWrapper = null;

			// Always create a new instance for the panel
			_geniusPanelWrapper = new PanelWrapper();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error opening Genius Panel: {ex.Message}");
		}
	}

	// Helper method to refresh the ribbon UI
	private void RefreshRibbon()
	{
		try
		{
			RemoveButtons();
			AddToUserInterface();
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
		}
	}

	// Generic helper method to remove buttons from specific ribbon/tab combinations
	private static void RemoveButtons()
	{
		try
		{
			var buttonRemovalConfigs = new[]
			{
				new { RibbonName = "Drawing", TabName = "id_TabAnnotate", ButtonInternalName = "ObsoletePrint" },
				new { RibbonName = "Drawing", TabName = "id_TabAnnotate", ButtonInternalName = "Explode iComponents" }
			};

			foreach (var config in buttonRemovalConfigs)
				RemoveButtonFromRibbonTab(config.RibbonName, config.TabName, config.ButtonInternalName);
		}
		catch (Exception ex)
		{
			Debug.Print($"Error removing buttons: {ex.Message}");
		}
	}

	// Generic method to remove a specific button from a ribbon tab
	private static void RemoveButtonFromRibbonTab(string ribbonName, string tabName,
		string buttonInternalName)
	{
		try
		{
			var ribbon = ThisApplication.UserInterfaceManager.Ribbons[ribbonName];

			var tab = ribbon?.RibbonTabs[tabName];
			if (tab is null) return;

			RemoveButtonFromPanels(tab.RibbonPanels, buttonInternalName);
		}
		catch (Exception ex)
		{
			Debug.Print($"Could not remove button '{buttonInternalName}' from {ribbonName}/{tabName}: {ex.Message}");
		}
	}

	// Generic method to remove a button from all panels in a tab
	private static void RemoveButtonFromPanels(RibbonPanels panels, string buttonInternalName)
	{
		foreach (var ctrl in panels.Cast<RibbonPanel>().Select(panel => GetControlsToRemove(panel, buttonInternalName))
		                           .SelectMany(controlsToRemove => controlsToRemove)) TryDeleteControl(ctrl);
	}

	// Generic method to find controls matching a button internal name
	private static List<CommandControl> GetControlsToRemove(RibbonPanel panel, string buttonInternalName)
	{
		return
		[
			.. panel.CommandControls.Cast<CommandControl>().Where(ctrl => IsMatchingControl(ctrl, buttonInternalName))
		];
	}

	// Generic method to check if control matches a button internal name
	private static bool IsMatchingControl(CommandControl ctrl, string buttonInternalName)
	{
		try
		{
			return ctrl?.ControlDefinition?.InternalName == buttonInternalName;
		}
		catch
		{
			return false;
		}
	}

	// Generic method to safely delete a control
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