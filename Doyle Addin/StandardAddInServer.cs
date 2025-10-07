using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using Doyle_Addin.Optional_Features;
using Doyle_Addin.Options;
using Inventor;
using File = System.IO.File;

namespace Doyle_Addin
{
    /// <inheritdoc />
    [ProgId("Test2.StandardAddInServer")]
    [Guid("513b9d7e-103e-4569-8eb5-ab3929cd33ad")]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private static Inventor.Application? appInstance;

        private UserInterfaceEvents? uiEvents;

        private UserInterfaceEvents? UiEvents
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (uiEvents != null)
                {
                    uiEvents.OnResetRibbonInterface -= UiEvents_OnResetRibbonInterface;
                }

                uiEvents = value;
                if (uiEvents != null)
                {
                    uiEvents.OnResetRibbonInterface += UiEvents_OnResetRibbonInterface;
                }
            }
        }
        private ButtonDefinition? dxfUpdate;

        private ButtonDefinition? DxfUpdate
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get => dxfUpdate;

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (dxfUpdate != null)
                {
                    dxfUpdate.OnExecute -= DXFUpdate_OnExecute;
                }

                dxfUpdate = value;
                if (dxfUpdate != null)
                {
                    dxfUpdate.OnExecute += DXFUpdate_OnExecute;
                }
            }
        }
        private ButtonDefinition? printUpdate;

        private ButtonDefinition? PrintUpdate
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get => printUpdate;

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (printUpdate != null)
                {
                    printUpdate.OnExecute -= PrintUpdate_OnExecute;
                }

                printUpdate = value;
                if (printUpdate != null)
                {
                    printUpdate.OnExecute += PrintUpdate_OnExecute;
                }
            }
        }
        private ButtonDefinition? optionsButton;

        private ButtonDefinition? OptionsButton
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get => optionsButton;

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (optionsButton != null)
                {
                    optionsButton.OnExecute -= OptionsButton_OnExecute;
                }

                optionsButton = value;
                if (optionsButton != null)
                {
                    optionsButton.OnExecute += OptionsButton_OnExecute;
                }
            }
        }
        private ButtonDefinition? obsoleteButton;

        private ButtonDefinition? ObsoleteButton
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get => obsoleteButton;

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (obsoleteButton != null)
                {
                    obsoleteButton.OnExecute -= ObsoleteButton_OnExecute;
                }

                obsoleteButton = value;
                if (obsoleteButton != null)
                {
                    obsoleteButton.OnExecute += ObsoleteButton_OnExecute;
                }
            }
        }

        #region ApplicationAddInServer Members

        // Inventor calls this method when it loads the AddIn. The AddInSiteObject provides access 
        // To the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        // The first time. However, with the introduction of the ribbon, this argument is always true.
        /// <inheritdoc />
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {

            CheckForUpdateAndDownloadAsync();

            // Initialize AddIn members.
            var thisApplication = addInSiteObject.Application;
            appInstance = thisApplication;

            // Get a reference to the ControlDefinitions object. 
            var controlDefs = thisApplication.CommandManager.ControlDefinitions;
            var oThemeManager = thisApplication.ThemeManager;

            var oTheme = oThemeManager.ActiveTheme;

            switch (oTheme.Name ?? "")
            {
                case "LightTheme":
                case "DarkTheme":
                {
                    var themeSuffix = oTheme.Name == "LightTheme" ? "Light" : "Dark";
                    int[] iconSizes = [16, 32];
                    var icons = new[]
                    {
                        new { Name = "PrintUpdate", Icon = "Doyle_Addin.PrintUpdateIcon.svg", InternalName = "printUpdate" },
                        new { Name = "DXFUpdate", Icon = "Doyle_Addin.DXFUpdateIcon.svg", InternalName = "dxfUpdate" },
                        new { Name = "Settings", Icon = "Doyle_Addin.SettingsIcon.svg", InternalName = "userOptions" },
                        new
                        {
                            Name = "ObsoletePrint", Icon = "Doyle_Addin.ObsoletePrint.svg",
                            InternalName = "ObsoletePrint"
                        }
                    };

                        foreach (var icon in icons)
                        {
                            var largeIcon = PictureConverter.SvgResourceToPictureDisp(icon.Icon, iconSizes[1],
                                iconSizes[1], themeSuffix);
                            var smallIcon = PictureConverter.SvgResourceToPictureDisp(icon.Icon, iconSizes[0],
                                iconSizes[0], themeSuffix);

                            // Try to remove the existing definition first if it exists
                            try
                            {
                                var existingDef = controlDefs[icon.InternalName];
                                if (existingDef is not null)
                                {
                                    existingDef.Delete();
                                }
                            }
                            catch
                            {
                                // Definition doesn't exist, which is fine
                            }

                            switch (icon.Name)
                            {
                                case "PrintUpdate":
                                {
                                    PrintUpdate = controlDefs.AddButtonDefinition("Print" + '\n' + "Update",
                                        "printUpdate", CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(),
                                        StandardIcon: smallIcon, LargeIcon: largeIcon);
                                    break;
                                }
                                case "DXFUpdate":
                                {
                                    DxfUpdate = controlDefs.AddButtonDefinition("DXF" + '\n' + "Update", "dxfUpdate",
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
                                case "ObsoletePrint":
                                {
                                    ObsoleteButton = controlDefs.AddButtonDefinition("Obsolete" + '\n' + "Print",
                                        "ObsoletePrint", CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(),
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
           var uiManager = typeof(UserInterfaceManager).InvokeMember("Application",
                    BindingFlags.GetProperty | BindingFlags.Static |
                    BindingFlags.NonPublic,
                    null, null, null) as UserInterfaceManager;

            // Ensure the option file exists with default values if it doesn't exist
            if (File.Exists(UserOptions.OptionsFilePath)) return;
            var defaultOptions = new UserOptions()
            {
                PrintExportLocation = @"P:\",
                DxfExportLocation = @"X:\"
            };
            defaultOptions.Save();
        }

        private static async void CheckForUpdateAndDownloadAsync()
        {
            try
            {
                var localVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
                // System.Windows.Forms.MessageBox.Show($"Local version: {localVersion}", "Debug")

                var releaseNullable = await GetLatestReleaseFromGitHub();
                if (!releaseNullable.HasValue)
                {
                    // System.Windows.Forms.MessageBox.Show("Could not fetch release info from GitHub.", "Debug")
                    return;
                }
                var release = releaseNullable.Value;

                var latestVersion = release.GetProperty("tag_name").GetString()?.TrimStart('v');
                // System.Windows.Forms.MessageBox.Show($"Latest GitHub version: {latestVersion}", "Debug")

                if (localVersion == null) return;
                var localVerObj = new Version(localVersion);
                if (latestVersion == null) return;
                var latestVerObj = new Version(latestVersion);
                if (latestVerObj > localVerObj)
                {
                    var result =
                        MessageBox.Show(
                            $@"A new version of the Doyle AddIn is available ({latestVersion}) . Update now?",
                            @"Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        await File.WriteAllTextAsync(
                            @"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update");
                        appInstance?.Quit();
                    }
                    else
                    {
                        await File.WriteAllTextAsync(
                            @"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update");
                        MessageBox.Show(@"The update will be installed after you close Inventor.", @"Update Scheduled",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // System.Windows.Forms.MessageBox.Show("You are running the latest version.", "Debug")
                }
            }
            catch (Exception)
            {
                // Optionally log or show error
            }
        }

        private static async Task<JsonElement?> GetLatestReleaseFromGitHub()
        {
            const string url = "https://api.github.com/repos/Bmassner/Doyle-AddIn/releases";
            using var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd("InventorAddinUpdater");
            var json = await client.GetStringAsync(url);
            var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (root.ValueKind == JsonValueKind.Array && root.GetArrayLength() > 0)
            {
                return root[0]; // Use the first release (most recent)
            }
            else
            {
                return null;
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
                if (PrintUpdate is not null)
                {
                    PrintUpdate.Delete();
                    PrintUpdate = null;
                }
            }
            catch
            {
                // ignored
            }

            try
            {
                if (DxfUpdate is not null)
                {
                    DxfUpdate.Delete();
                    DxfUpdate = null;
                }
            }
            catch
            {
                // ignored
            }

            try
            {
                if (OptionsButton is not null)
                {
                    OptionsButton.Delete();
                    OptionsButton = null;
                }
            }
            catch
            {
                // ignored
            }

            try
            {
                if (ObsoleteButton is not null)
                {
                    ObsoleteButton.Delete();
                    ObsoleteButton = null;
                }
            }
            catch
            {
                // ignored
            }

            // Release objects.
            UiEvents = null;

            // Check for pending update marker
            const string updateMarker = @"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt";
            const string updaterBat = @"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\Updater.bat";
            if (File.Exists(updateMarker) && File.Exists(updaterBat))
            {
                try
                {
                    // Start Updater.bat in a detached process
                    var psi = new ProcessStartInfo()
                    {
                        FileName = updaterBat,
                        WindowStyle = ProcessWindowStyle.Normal,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
                catch (Exception)
                {
                    // Optionally log or show error
                }
                // Remove marker
                File.Delete(updateMarker);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // This property is provided to allow the AddIn to expose an API of its own to other 
        // Programs. Typically, this would be done by implementing the AddIn's API
        // interface in a class and returning that class object through this property.
        /// <inheritdoc />
        public object? Automation => null;

        private static readonly string[] Item4 = ["id_PanelP_SheetMetalManageUnfold"];
        private static readonly string[] Item4Array = ["id_PanelP_ToolsOptions", "Add-Ins"];
        private static readonly string[] Item4Array0 = ["Add-Ins"];
        private static readonly string[] Item4Array2 = ["id_PanelP_ToolsOptions"];
        private static readonly string[] Item4Array3 = ["id_PanelP_FlatPatternExit"];
        private static readonly string[] Item4Array4 = ["id_PanelD_AnnotateRevision"];

        /// <inheritdoc />
        public void ExecuteCommand(int commandId)
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
            var uiManager = appInstance?.UserInterfaceManager;
            if (uiManager is null) return;

            // Define ribbon mappings for each document type
            var ribbonMappings = new Dictionary<string, Ribbon>()
            {
                { "Part", uiManager.Ribbons["Part"] }, { "Assembly", uiManager.Ribbons["Assembly"] },
                { "Drawing", uiManager.Ribbons["Drawing"] }, { "ZeroDoc", uiManager.Ribbons["ZeroDoc"] }
            };

            // Define button configurations with document type specificity
            // DXF button only appears on Part documents
            var dxfButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>()
            {
                Tuple.Create("id_TabSheetMetal", "dxfUpdate", DxfUpdate, Item4)!,
                Tuple.Create("id_TabFlatPattern", "dxfUpdate", DxfUpdate, Item4Array3)!,
                Tuple.Create("id_TabTools", "dxfUpdate", DxfUpdate, Item4Array2)!
            };

            // Option button appears on all document types
            var optionsButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>()
            {
                Tuple.Create("id_TabPlaceViews", "printUpdate", PrintUpdate, Item4Array0)!,
                Tuple.Create("id_TabAnnotate", "printUpdate", PrintUpdate, Item4Array0)!,
                Tuple.Create("id_TabTools", "userOptions", OptionsButton, Item4Array)!
            };

            // Obsolete Print button - only add if feature is enabled
            var obsoletePrintConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>();
            if (options.EnableObsoletePrint)
            {
                obsoletePrintConfigs.Add(Tuple.Create("id_TabAnnotate", "ObsoletePrint", ObsoleteButton, Item4Array4)!);
            }

            // Add buttons to appropriate ribbons based on document context
            foreach (var kvp in ribbonMappings)
            {
                var ribbonName = kvp.Key;
                var ribbon = kvp.Value;

                // Add the DXF button only to Part ribbon
                if (ribbonName == "Part")
                {
                    foreach (var config in dxfButtonConfigs)
                    {
                        var tabName = config.Item1;
                        var panelName = config.Item2;
                        var buttonDef = config.Item3;
                        var fallbackPanels = config.Item4;

                        AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);
                    }
                }

                // Add Options buttons to all ribbons
                foreach (var config in optionsButtonConfigs)
                {
                    var tabName = config.Item1;
                    var panelName = config.Item2;
                    var buttonDef = config.Item3;
                    var fallbackPanels = config.Item4;

                    AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);
                }

                // Add the Obsolete Print button to Drawing ribbon only if enabled
                if (ribbonName == "Drawing")
                {
                    foreach (var config in obsoletePrintConfigs)
                    {
                        var tabName = config.Item1;
                        var panelName = config.Item2;
                        var buttonDef = config.Item3;
                        var fallbackPanels = config.Item4;

                        AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);
                    }
                }
            }
        }

        // Helper method to add a button to specific ribbon with fallback handling
        private static void AddButtonToRibbon(Ribbon ribbon, string tabName, string panelName,
            ButtonDefinition buttonDef, string[] fallbackPanels)
        {
            try
            {
                // Get the tab
                var tab = ribbon.RibbonTabs[tabName];
                if (tab is null)
                    return;

                // Try to get the specified panel
                RibbonPanel? panel = null;
                try
                {
                    panel = tab.RibbonPanels[panelName];
                }
                catch
                {
                    // Try fallback panels
                    foreach (var fallbackPanel in fallbackPanels)
                    {
                        try
                        {
                            if (fallbackPanel == "Add-Ins")
                            {
                                // Create a new Add-Ins panel if it doesn't exist
                                panel = tab.RibbonPanels.Add("Add-Ins", panelName, Globals.AddInClientId());
                                break;
                            }
                            else
                            {
                                panel = tab.RibbonPanels[fallbackPanel];
                                if (panel is not null)
                                    break;
                            }
                        }
                        catch
                        {
                            // ignored
                        }
                    }
                }

                // Check if the button already exists in this panel
                if (panel is null) return;
                var buttonExists = false;
                foreach (CommandControl ctrl in panel.CommandControls)
                {
                    try
                    {
                        // Add null checks before accessing properties
                        if (ctrl?.ControlDefinition is null || buttonDef is null ||
                            (ctrl.ControlDefinition.InternalName ?? "") != (buttonDef.InternalName ?? "")) continue;
                        buttonExists = true;
                        break;
                    }
                    catch
                    {
                        // Skip this control if there's any issue
                    }
                }

                // Add the button only if it doesn't already exist
                if (!buttonExists)
                {
                    panel.CommandControls.AddButton(buttonDef, true);
                }
            }

            catch (Exception ex)
            {
                // Log specific errors for debugging
                Debug.Print(
                    $"Failed to add button '{(buttonDef is not null ? buttonDef.DisplayName : "Unknown")}' to tab '{tabName}' in ribbon '{ribbon.InternalName}': {ex.Message}");
            }
        }

        private void UiEvents_OnResetRibbonInterface(NameValueMap context)
        {
            AddToUserInterface();
        }

        private static void DXFUpdate_OnExecute(NameValueMap context)
        {
            if (appInstance != null)
            {
                Doyle_Addin.DxfUpdate.RunDxfUpdate(null, appInstance);
            }
        }

        private static void PrintUpdate_OnExecute(NameValueMap context)
        {
            Doyle_Addin.PrintUpdate.RunPrintUpdate(appInstance);
        }

        private void OptionsButton_OnExecute(NameValueMap context)
        {
            var optionsForm = new UserOptionsForm();
            var hwnd = (nint)(appInstance?.MainFrameHWND ?? 0);
            var result = optionsForm.ShowDialog(new Globals.WindowWrapper(hwnd));

            // Refresh the ribbon after options are saved
            if (result == DialogResult.OK)
            {
                RefreshRibbon();
            }
        }

        private static void ObsoleteButton_OnExecute(NameValueMap context)
        {
            if (appInstance != null)
            {
                ObsoletePrint.ApplyObsoletePrint(appInstance);
            }
        }

        // Helper method to refresh the ribbon UI
        private void RefreshRibbon()
        {
            try
            {
                // Remove the obsolete print button from all ribbons first
                RemoveObsoletePrintButton();

                // Re-add buttons with updated settings
                AddToUserInterface();
            }

            // MessageBox.Show("Ribbon updated successfully. Changes are now active.",
            // "Settings Applied",
            // MessageBoxButtons.OK,
            // MessageBoxIcon.Information)'
            catch (Exception)
            {
                // MessageBox.Show($"Error refreshing ribbon: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            }
        }

        // Helper method to remove the obsolete print button from ribbons
        private static void RemoveObsoletePrintButton()
        {
            try
            {
                var uiManager = appInstance?.UserInterfaceManager;
                if (uiManager is null) return;
                                var ribbon = uiManager.Ribbons["Drawing"];

                if (ribbon is null) return;
                try
                {
                    var tab = ribbon.RibbonTabs["id_TabAnnotate"];
                    if (tab is null) return;
                    // Find all panels and remove the obsolete print button
                    foreach (RibbonPanel panel in tab.RibbonPanels)
                    {
                        var controlsToRemove = new List<CommandControl>();

                        // Collect controls to remove (with null checks)
                        foreach (CommandControl ctrl in panel.CommandControls)
                        {
                            try
                            {
                                if (ctrl?.ControlDefinition is { InternalName: "ObsoletePrint" })
                                {
                                    controlsToRemove.Add(ctrl);
                                }
                            }
                            catch
                            {
                                // Skip this control if there's any issue accessing it
                            }
                        }

                        // Remove collected controls
                        foreach (var ctrl in controlsToRemove)
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
                }
                catch (Exception ex)
                {
                    // Button might not exist, continue
                    Debug.Print($"Could not remove button from Annotate tab: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log error but don't show to the user during refresh
                Debug.Print($"Error removing obsolete print button: {ex.Message}");
            }
        }

        #endregion
    }


    /// <summary>
    /// 
    /// </summary>
    public static class Globals
    {

        #region Function to get the add-in client ID.
        // This function uses reflection to get the GuidAttribute associated with the add-in.
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string AddInClientId()
        {
            var guid = "";
            try
            {
                var t = typeof(StandardAddInServer);
                var customAttributes = t.GetCustomAttributes(typeof(GuidAttribute), false);
                var guidAttribute = (GuidAttribute)customAttributes[0];
                guid = "{" + guidAttribute.Value + "}";
            }
            catch
            {
                // ignored
            }

            return guid;
        }

        #endregion

        #region hWnd Wrapper Class
        // This class is used to wrap a Win32 hWnd as a .NET IWind32Window class.
        // This is primarily used for parenting a dialog to the Inventor window.
        /// <inheritdoc />
        public class WindowWrapper(nint handle) : IWin32Window
        {
            /// <inheritdoc />
            public nint Handle { get; private set; } = handle;
        }

        #endregion
    }
}