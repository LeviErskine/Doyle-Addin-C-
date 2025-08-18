Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text.Json
Imports System.Windows.Forms
Imports Doyle_Addin.Options
Imports Inventor
Imports IPictureDisp = stdole.IPictureDisp
Imports File = System.IO.File


<ProgId("Test2.StandardAddInServer"),
	GuidAttribute("513b9d7e-103e-4569-8eb5-ab3929cd33ad")>
Public Class StandardAddInServer
	Implements ApplicationAddInServer

	Private WithEvents _uiEvents As UserInterfaceEvents
	Private WithEvents _dxfUpdate As ButtonDefinition
	Private WithEvents _printUpdate As ButtonDefinition
	Private WithEvents _optionsButton As ButtonDefinition

#Region "ApplicationAddInServer Members"

	' Inventor calls this method when it loads the AddIn. The AddInSiteObject provides access  
	' To the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
	' The first time. However, with the introduction of the ribbon, this argument is always true.
	Public Sub Activate(addInSiteObject As ApplicationAddInSite, firstTime As Boolean) _
		Implements ApplicationAddInServer.Activate

		CheckForUpdateAndDownloadAsync()

		' Initialize AddIn members.
		ThisApplication = addInSiteObject.Application

		' Get a reference to the ControlDefinitions object. 
		Dim controlDefs As ControlDefinitions = ThisApplication.CommandManager.ControlDefinitions
		Dim oThemeManager As ThemeManager
		oThemeManager = ThisApplication.ThemeManager

		Dim oTheme As Theme
		oTheme = oThemeManager.ActiveTheme

		Select Case oTheme.Name
			Case "LightTheme", "DarkTheme"
				Dim themeSuffix As String = If(oTheme.Name = "LightTheme", "Light", "Dark")
				Dim iconSizes = {16, 32}
				Dim icons = { _
					            New With {.Name = "PrintUpdate", .Icon = "Doyle_Addin.PrintUpdateIcon.svg"},
					            New With {.Name = "DXFUpdate", .Icon = "Doyle_Addin.DXFUpdateIcon.svg"},
					            New With {.Name = "Settings", .Icon = "Doyle_Addin.SettingsIcon.svg"}
				            }

				For Each icon In icons
					Dim largeIcon As IPictureDisp = PictureConverter.SvgResourceToPictureDisp(icon.Icon, iconSizes(1),
					                                                                          iconSizes(1), themeSuffix)
					Dim smallIcon As IPictureDisp = PictureConverter.SvgResourceToPictureDisp(icon.Icon, iconSizes(0),
					                                                                          iconSizes(0), themeSuffix)

					Select Case icon.Name
						Case "PrintUpdate"
							_printUpdate = controlDefs.AddButtonDefinition("Print Update", "printUpdate",
							                                               CommandTypesEnum.kShapeEditCmdType,
							                                               AddInClientId, , , smallIcon, largeIcon)
						Case "DXFUpdate"
							_dxfUpdate = controlDefs.AddButtonDefinition("DXF Update", "dxfUpdate",
							                                             CommandTypesEnum.kShapeEditCmdType,
							                                             AddInClientId, , , smallIcon, largeIcon)
						Case "Settings"
							_optionsButton = controlDefs.AddButtonDefinition("Export Options", "userOptions",
																			 CommandTypesEnum.kNonShapeEditCmdType,
																			 AddInClientId, , , smallIcon, largeIcon)
					End Select
				Next
		End Select

		' Add to the user interface if it's the first time.
		If firstTime Then
			AddToUserInterface()
		End If

		' Connect to the user-interface events to handle a ribbon reset.
		_uiEvents = ThisApplication.UserInterfaceManager.UserInterfaceEvents

		' Ensure the option file exists with default values if it doesn't exist
		If Not File.Exists(UserOptions.OptionsFilePath) Then
			Dim defaultOptions As New UserOptions With {
				    .PrintExportLocation = "P:\",
				    .DxfExportLocation = "X:\"
				    }
			defaultOptions.Save()
		End If
	End Sub

	Private Shared Async Sub CheckForUpdateAndDownloadAsync()
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
				Dim result = MessageBox.Show(
					$"A new version of the Doyle AddIn is available ({latestVersion}) . Update now?",
					"Update Available",
					MessageBoxButtons.YesNo,
					MessageBoxIcon.Question)
				If result = DialogResult.Yes Then
					File.WriteAllText("C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update")
					ThisApplication.Quit()
				Else
					File.WriteAllText("C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update")
					MessageBox.Show("The update will be installed after you close Inventor.", "Update Scheduled",
					                MessageBoxButtons.OK, MessageBoxIcon.Information)
				End If
			Else
				' System.Windows.Forms.MessageBox.Show("You are running the latest version.", "Debug")
			End If
		Catch ex As Exception
			' Optionally log or show error
		End Try
	End Sub

	Private Shared Async Function GetLatestReleaseFromGitHub() As Task(Of JsonElement?)
		Const url = "https://api.github.com/repos/Bmassner/Doyle-AddIn/releases"
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

	' Inventor calls this method when the AddIn is unloaded. The AddIn will be
	' unloaded either manually by the user or when the Inventor session is terminated.
	Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate

		' Release objects.
		_uiEvents = Nothing
		ThisApplication = Nothing

		' Check for pending update marker
		Const updateMarker = "C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt"
		Const updaterBat = "C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\Updater.bat"
		If File.Exists(updateMarker) AndAlso File.Exists(updaterBat) Then
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
			File.Delete(updateMarker)
		End If

		GC.Collect()
		GC.WaitForPendingFinalizers()
	End Sub

	' This property is provided to allow the AddIn to expose an API of its own to other 
	' Programs. Typically, this would be done by implementing the AddIn's API
	' interface in a class and returning that class object through this property.
	Public ReadOnly Property Automation As Object Implements ApplicationAddInServer.Automation
		Get
			Return Nothing
		End Get
	End Property

	Private Shared ReadOnly Item4 As String() = New String() {"id_PanelP_SheetMetalManageUnfold"}
	Private Shared ReadOnly Item4Array As String() = New String() {"id_PanelP_ToolsOptions", "Add-Ins"}
	Private Shared ReadOnly Item4Array0 As String() = New String() {"Add-Ins"}
	Private Shared ReadOnly Item4Array1 As String() = New String() {"Add-Ins"}
	Private Shared ReadOnly Item4Array2 As String() = New String() {"id_PanelP_ToolsOptions"}
	Private Shared ReadOnly Item4Array3 As String() = New String() {"id_PanelP_FlatPatternExit"}

	' Note:this method is now obsolete, you should use the 
	' ControlDefinition functionality for implementing commands.
	Public Sub ExecuteCommand(commandId As Integer) Implements ApplicationAddInServer.ExecuteCommand
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
			    Tuple.Create("id_TabSheetMetal", "dxfUpdate", _dxfUpdate,
			                 Item4),
			    Tuple.Create("id_TabFlatPattern", "dxfUpdate", _dxfUpdate, Item4Array3),
			    Tuple.Create("id_TabTools", "dxfUpdate", _dxfUpdate, Item4Array2)
			    }

		' Option button appears on all document types
		Dim optionsButtonConfigs = New List(Of Tuple(Of String, String, ButtonDefinition, String())) From {
			    Tuple.Create("id_TabPlaceViews", "printUpdate", _printUpdate, Item4Array1),
			    Tuple.Create("id_TabAnnotate", "printUpdate", _printUpdate, Item4Array0),
			    Tuple.Create("id_TabTools", "userOptions", _optionsButton,
			                 Item4Array)
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

	' Helper method to add a button to specific ribbon with fallback handling
	Private Shared Sub AddButtonToRibbon(ribbon As Ribbon, tabName As String, panelName As String,
	                                     buttonDef As ButtonDefinition, fallbackPanels As String())
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
							' Create a new Add-Ins panel if it doesn't exist
							panel = tab.RibbonPanels.Add("Add-Ins", panelName, AddInClientId())
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

			' Add the button if a panel was found
			panel?.CommandControls.AddButton(buttonDef, True)

		Catch ex As Exception
			' Log specific errors for debugging
			Debug.Print(
				$"Failed to add button '{buttonDef.DisplayName}' to tab '{tabName}' in ribbon '{ribbon.InternalName}': { _
				           ex.Message}")
		End Try
	End Sub

	Private Sub UiEvents_OnResetRibbonInterface(context As NameValueMap) Handles _uiEvents.OnResetRibbonInterface
		' The ribbon was reset, so add back the add-ins user-interface.
		AddToUserInterface()
	End Sub

	' Sample handler for the button.
	Private Shared Sub DXFUpdate_OnExecute(context As NameValueMap) Handles _dxfUpdate.OnExecute
		Call Sub() RunDxfUpdate(ThisApplication)
		'Call Sub() userName()
	End Sub

	Private Shared Sub PrintUpdate_OnExecute(context As NameValueMap) Handles _printUpdate.OnExecute
		Call Sub() RunPrintUpdate(ThisApplication)
	End Sub

	Private Shared Sub OptionsButton_OnExecute(context As NameValueMap) Handles _optionsButton.OnExecute
		Dim optionsForm As New UserOptionsForm()
		optionsForm.ShowDialog(New WindowWrapper(ThisApplication.MainFrameHWND))
	End Sub

#End Region
End Class


Public Module Globals

#Region "Function to get the add-in client ID."
	' This function uses reflection to get the GuidAttribute associated with the add-in.
	Public Function AddInClientId() As String
		Dim guid = ""
		Try
			Dim t As Type = GetType(StandardAddInServer)
			Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
			Dim guidAttribute = CType(customAttributes(0), GuidAttribute)
			guid = "{" + guidAttribute.Value.ToString() + "}"
		Catch
		End Try

		Return guid
	End Function

#End Region

#Region "hWnd Wrapper Class"
	' This class is used to wrap a Win32 hWnd as a .NET IWind32Window class.
	' This is primarily used for parenting a dialog to the Inventor window.
	Public Class WindowWrapper
		Implements IWin32Window

		Public Sub New(handle As IntPtr)
			Me.Handle = handle
		End Sub

		Public ReadOnly Property Handle As IntPtr _
			Implements IWin32Window.Handle
	End Class

#End Region
End Module
