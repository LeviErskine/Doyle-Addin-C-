using System;
using System.Collections.Generic;
using System.Diagnostics;
using File = System.IO.File;
using System.Net.Http;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Doyle_Addin.Options;
using Inventor;

namespace Doyle_Addin
{
    /// <inheritdoc />
    [ProgId("Test2.StandardAddInServer")]
    [Guid("513b9d7e-103e-4569-8eb5-ab3929cd33ad")]
    public class StandardAddInServer : ApplicationAddInServer
    {

        private UserInterfaceEvents _uiEvents;

        private UserInterfaceEvents uiEvents
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _uiEvents;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_uiEvents != null)
                {
                    _uiEvents.OnResetRibbonInterface -= UiEvents_OnResetRibbonInterface;
                }

                _uiEvents = value;
                if (_uiEvents != null)
                {
                    _uiEvents.OnResetRibbonInterface += UiEvents_OnResetRibbonInterface;
                }
            }
        }
        private ButtonDefinition _dxfUpdate;

        private ButtonDefinition dxfUpdate
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _dxfUpdate;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_dxfUpdate != null)
                {
                    _dxfUpdate.OnExecute -= DXFUpdate_OnExecute;
                }

                _dxfUpdate = value;
                if (_dxfUpdate != null)
                {
                    _dxfUpdate.OnExecute += DXFUpdate_OnExecute;
                }
            }
        }
        private ButtonDefinition _printUpdate;

        private ButtonDefinition printUpdate
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _printUpdate;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_printUpdate != null)
                {
                    _printUpdate.OnExecute -= PrintUpdate_OnExecute;
                }

                _printUpdate = value;
                if (_printUpdate != null)
                {
                    _printUpdate.OnExecute += PrintUpdate_OnExecute;
                }
            }
        }
        private ButtonDefinition _optionsButton;

        private ButtonDefinition optionsButton
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _optionsButton;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_optionsButton != null)
                {
                    _optionsButton.OnExecute -= OptionsButton_OnExecute;
                }

                _optionsButton = value;
                if (_optionsButton != null)
                {
                    _optionsButton.OnExecute += OptionsButton_OnExecute;
                }
            }
        }
        private ButtonDefinition _obsoleteButton;

        private ButtonDefinition obsoleteButton
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _obsoleteButton;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_obsoleteButton != null)
                {
                    _obsoleteButton.OnExecute -= ObsoleteButton_OnExecute;
                }

                _obsoleteButton = value;
                if (_obsoleteButton != null)
                {
                    _obsoleteButton.OnExecute += ObsoleteButton_OnExecute;
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
            GlobalsHelpers.ThisApplication = addInSiteObject.Application;

            // Get a reference to the ControlDefinitions object. 
            var controlDefs = GlobalsHelpers.ThisApplication.CommandManager.ControlDefinitions;
            ThemeManager oThemeManager;
            oThemeManager = GlobalsHelpers.ThisApplication.ThemeManager;

            Theme oTheme;
            oTheme = oThemeManager.ActiveTheme;

            switch (oTheme.Name ?? "")
            {
                case "LightTheme":
                case "DarkTheme":
                    {
                        string themeSuffix = oTheme.Name == "LightTheme" ? "Light" : "Dark";
                        int[] iconSizes = new[] { 16, 32 };
                        var icons = new[] { new { Name = "PrintUpdate", Icon = "Doyle_Addin.PrintUpdateIcon.svg", InternalName = "printUpdate" }, new { Name = "DXFUpdate", Icon = "Doyle_Addin.DXFUpdateIcon.svg", InternalName = "dxfUpdate" }, new { Name = "Settings", Icon = "Doyle_Addin.SettingsIcon.svg", InternalName = "userOptions" }, new { Name = "ObsoletePrint", Icon = "Doyle_Addin.ObsoletePrint.svg", InternalName = "ObsoletePrint" } };

                        foreach (var icon in icons)
                        {
                            var largeIcon = PictureConverter.SvgResourceToPictureDisp(icon.Icon, iconSizes[1], iconSizes[1], themeSuffix);
                            var smallIcon = PictureConverter.SvgResourceToPictureDisp(icon.Icon, iconSizes[0], iconSizes[0], themeSuffix);

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

                            switch (icon.Name ?? "")
                            {
                                case "PrintUpdate":
                                    {
                                        printUpdate = controlDefs.AddButtonDefinition("Print" + '\n' + "Update", "printUpdate", CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(), StandardIcon: smallIcon, LargeIcon: largeIcon);
                                        break;
                                    }
                                case "DXFUpdate":
                                    {
                                        dxfUpdate = controlDefs.AddButtonDefinition("DXF" + '\n' + "Update", "dxfUpdate", CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId(), StandardIcon: smallIcon, LargeIcon: largeIcon);
                                        break;
                                    }
                                case "Settings":
                                    {
                                        optionsButton = controlDefs.AddButtonDefinition("Options", "userOptions", CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(), StandardIcon: smallIcon, LargeIcon: largeIcon);
                                        break;
                                    }
                                case "ObsoletePrint":
                                    {
                                        obsoleteButton = controlDefs.AddButtonDefinition("Obsolete" + '\n' + "Print", "ObsoletePrint", CommandTypesEnum.kNonShapeEditCmdType, Globals.AddInClientId(), StandardIcon: smallIcon, LargeIcon: largeIcon);
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
            uiEvents = GlobalsHelpers.ThisApplication.UserInterfaceManager.UserInterfaceEvents;

            // Ensure the option file exists with default values if it doesn't exist
            if (!File.Exists(UserOptions.OptionsFilePath))
            {
                var defaultOptions = new UserOptions()
                {
                    PrintExportLocation = @"P:\",
                    DxfExportLocation = @"X:\"
                };
                defaultOptions.Save();
            }
        }

        private static async void CheckForUpdateAndDownloadAsync()
        {
            try
            {
                string localVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                // System.Windows.Forms.MessageBox.Show($"Local version: {localVersion}", "Debug")

                var releaseNullable = await GetLatestReleaseFromGitHub();
                if (!releaseNullable.HasValue)
                {
                    // System.Windows.Forms.MessageBox.Show("Could not fetch release info from GitHub.", "Debug")
                    return;
                }
                var release = releaseNullable.Value;

                string latestVersion = release.GetProperty("tag_name").GetString().TrimStart('v');
                // System.Windows.Forms.MessageBox.Show($"Latest GitHub version: {latestVersion}", "Debug")

                var localVerObj = new Version(localVersion);
                var latestVerObj = new Version(latestVersion);
                if (latestVerObj > localVerObj)
                {
                    var result = MessageBox.Show($"A new version of the Doyle AddIn is available ({latestVersion}) . Update now?", "Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        File.WriteAllText(@"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update");
                        GlobalsHelpers.ThisApplication.Quit();
                    }
                    else
                    {
                        File.WriteAllText(@"C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update");
                        MessageBox.Show("The update will be installed after you close Inventor.", "Update Scheduled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // System.Windows.Forms.MessageBox.Show("You are running the latest version.", "Debug")
                }
            }
            catch (Exception ex)
            {
                // Optionally log or show error
            }
        }

        private static async Task<JsonElement?> GetLatestReleaseFromGitHub()
        {
            const string url = "https://api.github.com/repos/LeviErskine/Doyle-Addin-C-/releases";
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("InventorAddinUpdater");
                string json = await client.GetStringAsync(url);
                var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.Array && root.GetArrayLength() > 0)
                {
                    return root[0]; // Use the first release (most recent)
                }
                else
                {
                    return default;
                }
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
                if (printUpdate is not null)
                {
                    printUpdate.Delete();
                    printUpdate = null;
                }
            }
            catch
            {
            }

            try
            {
                if (dxfUpdate is not null)
                {
                    dxfUpdate.Delete();
                    dxfUpdate = null;
                }
            }
            catch
            {
            }

            try
            {
                if (optionsButton is not null)
                {
                    optionsButton.Delete();
                    optionsButton = null;
                }
            }
            catch
            {
            }

            try
            {
                if (obsoleteButton is not null)
                {
                    obsoleteButton.Delete();
                    obsoleteButton = null;
                }
            }
            catch
            {
            }

            // Release objects.
            uiEvents = null;
            GlobalsHelpers.ThisApplication = null;

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
                catch (Exception ex)
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
        public object Automation
        {
            get
            {
                return null;
            }
        }

        private static readonly string[] Item4 = new string[] { "id_PanelP_SheetMetalManageUnfold" };
        private static readonly string[] Item4Array = new string[] { "id_PanelP_ToolsOptions", "Add-Ins" };
        private static readonly string[] Item4Array0 = new string[] { "Add-Ins" };
        private static readonly string[] Item4Array2 = new string[] { "id_PanelP_ToolsOptions" };
        private static readonly string[] Item4Array3 = new string[] { "id_PanelP_FlatPatternExit" };
        private static readonly string[] Item4Array4 = new string[] { "id_PanelD_AnnotateRevision" };

        /// <inheritdoc />
        public void ExecuteCommand(int commandId)
        {
        }

        #endregion

        #region User interface definition
        // Sub where the user-interface creation is done.  This is called when
        // the add-in loaded and also if the user interface is reset.
        private void AddToUserInterface()
        {
            // Load user options to check feature flags
            var options = UserOptions.Load();

            // Cache frequently used objects
            var uiManager = GlobalsHelpers.ThisApplication.UserInterfaceManager;

            // Define ribbon mappings for each document type
            var ribbonMappings = new Dictionary<string, Ribbon>() { { "Part", uiManager.Ribbons["Part"] }, { "Assembly", uiManager.Ribbons["Assembly"] }, { "Drawing", uiManager.Ribbons["Drawing"] }, { "ZeroDoc", uiManager.Ribbons["ZeroDoc"] } };

            // Define button configurations with document type specificity
            // DXF button only appears on Part documents
            var dxfButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>() { Tuple.Create("id_TabSheetMetal", "dxfUpdate", dxfUpdate, Item4), Tuple.Create("id_TabFlatPattern", "dxfUpdate", dxfUpdate, Item4Array3), Tuple.Create("id_TabTools", "dxfUpdate", dxfUpdate, Item4Array2) };

            // Option button appears on all document types
            var optionsButtonConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>() { Tuple.Create("id_TabPlaceViews", "printUpdate", printUpdate, Item4Array0), Tuple.Create("id_TabAnnotate", "printUpdate", printUpdate, Item4Array0), Tuple.Create("id_TabTools", "userOptions", optionsButton, Item4Array) };

            // Obsolete Print button - only add if feature is enabled
            var obsoletePrintConfigs = new List<Tuple<string, string, ButtonDefinition, string[]>>();
            if (options.EnableObsoletePrint)
            {
                obsoletePrintConfigs.Add(Tuple.Create("id_TabAnnotate", "ObsoletePrint", obsoleteButton, Item4Array4));
            }

            // Add buttons to appropriate ribbons based on document context
            foreach (var kvp in ribbonMappings)
            {
                string ribbonName = kvp.Key;
                var ribbon = kvp.Value;

                // Add the DXF button only to Part ribbon
                if (ribbonName == "Part")
                {
                    foreach (var config in dxfButtonConfigs)
                    {
                        string tabName = config.Item1;
                        string panelName = config.Item2;
                        var buttonDef = config.Item3;
                        string[] fallbackPanels = config.Item4;

                        AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);
                    }
                }

                // Add Options buttons to all ribbons
                foreach (var config in optionsButtonConfigs)
                {
                    string tabName = config.Item1;
                    string panelName = config.Item2;
                    var buttonDef = config.Item3;
                    string[] fallbackPanels = config.Item4;

                    AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);
                }

                // Add the Obsolete Print button to Drawing ribbon only if enabled
                if (ribbonName == "Drawing")
                {
                    foreach (var config in obsoletePrintConfigs)
                    {
                        string tabName = config.Item1;
                        string panelName = config.Item2;
                        var buttonDef = config.Item3;
                        string[] fallbackPanels = config.Item4;

                        AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels);
                    }
                }
            }
        }

        // Helper method to add a button to specific ribbon with fallback handling
        private static void AddButtonToRibbon(Ribbon ribbon, string tabName, string panelName, ButtonDefinition buttonDef, string[] fallbackPanels)
        {
            try
            {
                // Get the tab
                var tab = ribbon.RibbonTabs[tabName];
                if (tab is null)
                    return;

                // Try to get the specified panel
                RibbonPanel panel = null;
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
                            continue;
                        }
                    }
                }

                // Check if the button already exists in this panel
                if (panel is not null)
                {
                    bool buttonExists = false;
                    foreach (CommandControl ctrl in panel.CommandControls)
                    {
                        try
                        {
                            // Add null checks before accessing properties
                            if (ctrl is not null && ctrl.ControlDefinition is not null && buttonDef is not null && (ctrl.ControlDefinition.InternalName ?? "") == (buttonDef.InternalName ?? ""))
                            {
                                buttonExists = true;
                                break;
                            }
                        }
                        catch
                        {
                            // Skip this control if there's any issue
                            continue;
                        }
                    }

                    // Add the button only if it doesn't already exist
                    if (!buttonExists)
                    {
                        panel.CommandControls.AddButton(buttonDef, true);
                    }
                }
            }

            catch (Exception ex)
            {
                // Log specific errors for debugging
                Debug.Print($"Failed to add button '{(buttonDef is not null ? buttonDef.DisplayName : "Unknown")}' to tab '{tabName}' in ribbon '{ribbon.InternalName}': {ex.Message}");

            }
        }

        private void UiEvents_OnResetRibbonInterface(NameValueMap context)
        {
            AddToUserInterface();
        }

        private static void DXFUpdate_OnExecute(NameValueMap context)
        {
            new Action(() => DxfUpdate.RunDxfUpdate(GlobalsHelpers.ThisApplication))();
        }

        private static void PrintUpdate_OnExecute(NameValueMap context)
        {
            new Action(() => PrintUpdate.RunPrintUpdate(GlobalsHelpers.ThisApplication))();
        }

        private void OptionsButton_OnExecute(NameValueMap context)
        {
            var optionsForm = new UserOptionsForm();
            var result = optionsForm.ShowDialog(new Globals.WindowWrapper(GlobalsHelpers.ThisApplication.MainFrameHWND));

            // Refresh the ribbon after options are saved
            if (result == DialogResult.OK)
            {
                RefreshRibbon();
            }
        }

        private void ObsoleteButton_OnExecute(NameValueMap context)
        {
            new Action(() => ObsoletePrint.ApplyObsoletePrint(GlobalsHelpers.ThisApplication))();
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
            catch (Exception ex)
            {
                // MessageBox.Show($"Error refreshing ribbon: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            }
        }

        // Helper method to remove the obsolete print button from ribbons
        private void RemoveObsoletePrintButton()
        {
            try
            {
                var uiManager = GlobalsHelpers.ThisApplication.UserInterfaceManager;
                var ribbon = uiManager.Ribbons["Drawing"];

                if (ribbon is not null)
                {
                    try
                    {
                        var tab = ribbon.RibbonTabs["id_TabAnnotate"];
                        if (tab is not null)
                        {
                            // Find all panels and remove the obsolete print button
                            foreach (RibbonPanel panel in tab.RibbonPanels)
                            {
                                var controlsToRemove = new List<CommandControl>();

                                // Collect controls to remove (with null checks)
                                foreach (CommandControl ctrl in panel.CommandControls)
                                {
                                    try
                                    {
                                        if (ctrl is not null && ctrl.ControlDefinition is not null && ctrl.ControlDefinition.InternalName == "ObsoletePrint")
                                        {
                                            controlsToRemove.Add(ctrl);
                                        }
                                    }
                                    catch
                                    {
                                        // Skip this control if there's any issue accessing it
                                        continue;
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
                    }
                    catch (Exception ex)
                    {
                        // Button might not exist, continue
                        Debug.Print($"Could not remove button from Annotate tab: {ex.Message}");
                    }
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
        
        /// <summary>
        /// This function uses reflection to get the GuidAttribute associated with the add-in. 
        /// </summary>
        /// <returns></returns>
        public static string AddInClientId()
        {
            string guid = "";
            try
            {
                var t = typeof(StandardAddInServer);
                object[] customAttributes = t.GetCustomAttributes(typeof(GuidAttribute), false);
                GuidAttribute guidAttribute = (GuidAttribute)customAttributes[0];
                guid = "{" + guidAttribute.Value.ToString() + "}";
            }
            catch
            {
            }

            return guid;
        }

        #endregion

        #region hWnd Wrapper Class
        // This class is used to wrap a Win32 hWnd as a .NET IWind32Window class.
        // This is primarily used for parenting a dialog to the Inventor window.
        /// <inheritdoc />
        public class WindowWrapper : IWin32Window
        {

            /// <summary>
            /// 
            /// </summary>
            /// <param name="handle"></param>
            public WindowWrapper(nint handle)
            {
                Handle = handle;
            }

            /// <inheritdoc />
            public nint Handle { get; private set; }
        }

        #endregion
    }
}