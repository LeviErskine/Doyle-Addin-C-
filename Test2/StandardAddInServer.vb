Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text.Json
Imports Inventor

Namespace DoyleAddin
    <ProgIdAttribute("Test2.StandardAddInServer"),
    GuidAttribute("513b9d7e-103e-4569-8eb5-ab3929cd33ad")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents UiEvents As UserInterfaceEvents
        Private WithEvents DXFUpdate As ButtonDefinition
        Private WithEvents PrintUpdate As ButtonDefinition
        Private WithEvents OptionsButton As ButtonDefinition

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            CheckForUpdateAndDownloadAsync()

            ' Initialize AddIn members.
            ThisApplication = addInSiteObject.Application

            ' Get a reference to the UserInterfaceManager object. 
            Dim UIManager As Inventor.UserInterfaceManager = ThisApplication.UserInterfaceManager

            ' Get a reference to the ControlDefinitions object. 
            Dim controlDefs As ControlDefinitions = ThisApplication.CommandManager.ControlDefinitions
            Dim oThemeManager As Inventor.ThemeManager
            oThemeManager = ThisApplication.ThemeManager

            Dim oTheme As Inventor.Theme
            oTheme = oThemeManager.ActiveTheme

            Select Case oTheme.Name
                    Case "LightTheme", "DarkTheme"
                    Dim themeSuffix As String = If(oTheme.Name = "LightTheme", "Light", "Dark")
                    Dim PrintUpdateIcon As String = "Doyle_Addin.PrintUpdateIcon.svg"
                    Dim DXFUpdateIcon As String = "Doyle_Addin.DXFUpdateIcon.svg"
                    Dim SettingsIcon As String = "Doyle_Addin.SettingsIcon.svg"
                    Dim PrintUpdateIconLarge As stdole.IPictureDisp = PictureConverter.SvgResourceToPictureDisp(PrintUpdateIcon, 32, 32, themeSuffix)
                    Dim PrintUpdateIconSmall As stdole.IPictureDisp = PictureConverter.SvgResourceToPictureDisp(PrintUpdateIcon, 16, 16, themeSuffix)
                    Dim DXFUpdateIconSmall As stdole.IPictureDisp = PictureConverter.SvgResourceToPictureDisp(DXFUpdateIcon, 16, 16, themeSuffix)
                    Dim DXFUpdateIconLarge As stdole.IPictureDisp = PictureConverter.SvgResourceToPictureDisp(DXFUpdateIcon, 32, 32, themeSuffix)
                    Dim SettingsIconSmall As stdole.IPictureDisp = PictureConverter.SvgResourceToPictureDisp(SettingsIcon, 16, 16, themeSuffix)
                    Dim SettingsIconLarge As stdole.IPictureDisp = PictureConverter.SvgResourceToPictureDisp(SettingsIcon, 32, 32, themeSuffix)
                    DXFUpdate = controlDefs.AddButtonDefinition("DXF Update", "dxfUpdate", CommandTypesEnum.kShapeEditCmdType, AddInClientID, , , DXFUpdateIconSmall, DXFUpdateIconLarge)
                    PrintUpdate = controlDefs.AddButtonDefinition("Print Update", "printUpdate", CommandTypesEnum.kShapeEditCmdType, AddInClientID, , , PrintUpdateIconSmall, PrintUpdateIconLarge)
                    OptionsButton = controlDefs.AddButtonDefinition("Options", "userOptions", CommandTypesEnum.kNonShapeEditCmdType, AddInClientID, , , SettingsIconSmall, SettingsIconLarge)
            End Select

                ' Add to the user interface, if it's the first time.
                If firstTime Then
                AddToUserInterface()
            End If

            ' Connect to the user-interface events to handle a ribbon reset.
            UiEvents = ThisApplication.UserInterfaceManager.UserInterfaceEvents

            ' Ensure the options file exists with default values if it doesn't exist
            If Not IO.File.Exists(UserOptions.OptionsFilePath) Then
                Dim defaultOptions As New UserOptions With {
                    .PrintExportLocation = "P:\",
                    .DXFExportLocation = "X:\"
                }
                defaultOptions.Save()
            End If
        End Sub

        Private Async Sub CheckForUpdateAndDownloadAsync()
            Try
                Dim localVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString()
                ' System.Windows.Forms.MessageBox.Show($"Local version: {localVersion}", "Debug")

                Dim releaseNullable = Await GetLatestReleaseFromGitHub()
                If Not releaseNullable.HasValue Then
                    '  System.Windows.Forms.MessageBox.Show("Could not fetch release info from GitHub.", "Debug")
                    Exit Sub
                End If
                Dim release = releaseNullable.Value

                Dim latestVersion As String = release.GetProperty("tag_name").GetString().TrimStart("v"c)
                '  System.Windows.Forms.MessageBox.Show($"Latest GitHub version: {latestVersion}", "Debug")

                Dim localVerObj As New Version(localVersion)
                Dim latestVerObj As New Version(latestVersion)
                If latestVerObj > localVerObj Then
                    Dim result = System.Windows.Forms.MessageBox.Show(
                $"A new version of the Doyle AddIn is available ({latestVersion}) . Update now?",
                "Update Available",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question)
                    If result = System.Windows.Forms.DialogResult.Yes Then
                        IO.File.WriteAllText("C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update")
                        ThisApplication.Quit()
                    Else
                        IO.File.WriteAllText("C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update")
                        System.Windows.Forms.MessageBox.Show("The update will be installed after you close Inventor.", "Update Scheduled", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                    End If
                Else
                    ' System.Windows.Forms.MessageBox.Show("You are running the latest version.", "Debug")
                End If
            Catch ex As Exception
                ' Optionally log or show error
            End Try
        End Sub

        Private Async Function GetLatestReleaseFromGitHub() As Task(Of JsonElement?)
            Dim url As String = "https://api.github.com/repos/Bmassner/Doyle-AddIn/releases"
            Using client As New HttpClient()
                client.DefaultRequestHeaders.UserAgent.ParseAdd("InventorAddinUpdater")
                Dim json = Await client.GetStringAsync(url)
                Dim doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                If root.ValueKind = JsonValueKind.Array AndAlso root.GetArrayLength() > 0 Then
                    Return root(0) ' Use the first release (most recent)
                Else
                    Return Nothing
                End If
            End Using
        End Function

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' Release objects.
            UiEvents = Nothing
            ThisApplication = Nothing

            ' Check for pending update marker
            Dim updateMarker As String = "C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt"
            Dim updaterBat As String = "C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\Updater.bat"
            If IO.File.Exists(updateMarker) AndAlso IO.File.Exists(updaterBat) Then
                Try
                    ' Start Updater.bat in a detached process
                    Dim psi As New ProcessStartInfo With {
                        .FileName = updaterBat,
                        .WindowStyle = ProcessWindowStyle.Normal,
                        .UseShellExecute = True
                    }
                    Process.Start(psi)
                Catch ex As Exception
                    ' Optionally log or show error
                End Try
                ' Remove marker
                IO.File.Delete(updateMarker)
            End If

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            ' Cache frequently used objects
            Dim uiManager = ThisApplication.UserInterfaceManager

            ' Define ribbon mappings for each document type
            Dim ribbonMappings = New Dictionary(Of String, Ribbon) From {
                {"Part", uiManager.Ribbons.Item("Part")},
                {"Assembly", uiManager.Ribbons.Item("Assembly")},
                {"Drawing", uiManager.Ribbons.Item("Drawing")},
                {"ZeroDoc", uiManager.Ribbons.Item("ZeroDoc")}
            }

            ' Define button configurations with document type specificity
            ' DXF button only appears on Part documents
            Dim dxfButtonConfigs = New List(Of Tuple(Of String, String, ButtonDefinition, String())) From {
                Tuple.Create("id_TabSheetMetal", "dxfUpdate", DXFUpdate, New String() {"id_PanelP_SheetMetalManageUnfold"}),
                Tuple.Create("id_TabFlatPattern", "dxfUpdate", DXFUpdate, New String() {"id_PanelP_FlatPatternExit"}),
                Tuple.Create("id_TabTools", "dxfUpdate", DXFUpdate, New String() {"id_PanelP_ToolsOptions"})
            }

            ' Options button appears on all document types
            Dim optionsButtonConfigs = New List(Of Tuple(Of String, String, ButtonDefinition, String())) From {
                Tuple.Create("id_TabPlaceViews", "printUpdate", PrintUpdate, New String() {"Add-Ins"}),
                Tuple.Create("id_TabAnnotate", "printUpdate", PrintUpdate, New String() {"Add-Ins"}),
                Tuple.Create("id_TabTools", "userOptions", OptionsButton, New String() {"id_PanelP_ToolsOptions", "Add-Ins"})
            }

            ' Add buttons to appropriate ribbons based on document context
            For Each kvp In ribbonMappings
                Dim ribbonName = kvp.Key
                Dim ribbon = kvp.Value

                ' Add DXF button only to Part ribbon
                If ribbonName = "Part" Then
                    For Each config In dxfButtonConfigs
                        Dim tabName = config.Item1
                        Dim panelName = config.Item2
                        Dim buttonDef = config.Item3
                        Dim fallbackPanels = config.Item4

                        AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels)
                    Next
                End If

                ' Add Options button to all ribbons
                For Each config In optionsButtonConfigs
                    Dim tabName = config.Item1
                    Dim panelName = config.Item2
                    Dim buttonDef = config.Item3
                    Dim fallbackPanels = config.Item4

                    AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels)
                Next
            Next
        End Sub

        ' Helper method to add a button to a specific ribbon with fallback handling
        Private Sub AddButtonToRibbon(ribbon As Ribbon, tabName As String, panelName As String, buttonDef As ButtonDefinition, fallbackPanels As String())
            Try
                ' Get the tab
                Dim tab As RibbonTab = ribbon.RibbonTabs.Item(tabName)
                If tab Is Nothing Then Return

                ' Try to get the specified panel
                Dim panel As RibbonPanel = Nothing
                Try
                    panel = tab.RibbonPanels.Item(panelName)
                Catch
                    ' Try fallback panels
                    For Each fallbackPanel In fallbackPanels
                        Try
                            If fallbackPanel = "Add-Ins" Then
                                ' Create new Add-Ins panel if it doesn't exist
                                panel = tab.RibbonPanels.Add("Add-Ins", panelName, AddInClientID())
                                Exit For
                            Else
                                panel = tab.RibbonPanels.Item(fallbackPanel)
                                If panel IsNot Nothing Then Exit For
                            End If
                        Catch
                            Continue For
                        End Try
                    Next
                End Try

                ' Add the button if panel was found
                If panel IsNot Nothing Then
                    panel.CommandControls.AddButton(buttonDef, True)
                End If

            Catch ex As Exception
                ' Log specific errors for debugging
                Debug.Print($"Failed to add button '{buttonDef.DisplayName}' to tab '{tabName}' in ribbon '{ribbon.InternalName}': {ex.Message}")
            End Try
        End Sub

        Private Sub UiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles UiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        Private Sub DXFUpdate_OnExecute(Context As NameValueMap) Handles DXFUpdate.OnExecute
            Call Sub() RunDxfUpdate(ThisApplication)
            'Call Sub() userName()
        End Sub

        Private Sub PrintUpdate_OnExecute(Context As NameValueMap) Handles PrintUpdate.OnExecute
            Call Sub() RunPrintUpdate(ThisApplication)
        End Sub
        Private Sub OptionsButton_OnExecute(Context As NameValueMap) Handles OptionsButton.OnExecute
            Dim optionsForm As New UserOptionsForm()
            optionsForm.ShowDialog(New WindowWrapper(ThisApplication.MainFrameHWND))
        End Sub
#End Region

    End Class
End Namespace


Public Module Globals

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(DoyleAddin.StandardAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(ThisApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private ReadOnly _hwnd As IntPtr
    End Class
#End Region
End Module
