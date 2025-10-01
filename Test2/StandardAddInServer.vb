Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text.Json
Imports System.Windows.Forms
Imports Doyle_Addin.Options
Imports Inventor
Imports IPictureDisp = stdole.IPictureDisp
Imports File = System.IO.File


<ProgId("Test2.StandardAddInServer"), GuidAttribute("513b9d7e-103e-4569-8eb5-ab3929cd33ad")>
Public Class StandardAddInServer
	Implements ApplicationAddInServer

	Private WithEvents uiEvents As UserInterfaceEvents
	Private WithEvents dxfUpdate As ButtonDefinition
	Private WithEvents printUpdate As ButtonDefinition
	Private WithEvents optionsButton As ButtonDefinition
	Private WithEvents obsoleteButton As ButtonDefinition

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
				Dim icons =
					    {New With {.Name = "PrintUpdate", .Icon = "Doyle_Addin.PrintUpdateIcon.svg", .InternalName = "printUpdate"},
					     New With {.Name = "DXFUpdate", .Icon = "Doyle_Addin.DXFUpdateIcon.svg", .InternalName = "dxfUpdate"},
					     New With {.Name = "Settings", .Icon = "Doyle_Addin.SettingsIcon.svg", .InternalName = "userOptions"},
					     New With {.Name = "ObsoletePrint", .Icon = "Doyle_Addin.ObsoletePrint.svg", .InternalName = "ObsoletePrint"}}

				For Each icon In icons
					Dim largeIcon As IPictureDisp = PictureConverter.SvgResourceToPictureDisp _
						    (icon.Icon, iconSizes(1), iconSizes(1), themeSuffix)
					Dim smallIcon As IPictureDisp = PictureConverter.SvgResourceToPictureDisp _
						    (icon.Icon, iconSizes(0), iconSizes(0), themeSuffix)

					' Try to remove the existing definition first if it exists
					Try
						Dim existingDef = controlDefs.Item(icon.InternalName)
						If existingDef IsNot Nothing Then
							existingDef.Delete()
						End If
					Catch
						' Definition doesn't exist, which is fine
					End Try

					Select Case icon.Name
						Case "PrintUpdate"
							printUpdate = controlDefs.AddButtonDefinition _
								("Print" & Chr(10) & "Update",
								 "printUpdate",
								 CommandTypesEnum.kShapeEditCmdType,
								 AddInClientId,
								 ,
								 ,
								 smallIcon,
								 largeIcon)
						Case "DXFUpdate"
							dxfUpdate = controlDefs.AddButtonDefinition _
								("DXF" & Chr(10) & "Update",
								 "dxfUpdate",
								 CommandTypesEnum.kShapeEditCmdType,
								 AddInClientId,
								 ,
								 ,
								 smallIcon,
								 largeIcon)
						Case "Settings"
							optionsButton = controlDefs.AddButtonDefinition _
								("Options", "userOptions", CommandTypesEnum.kNonShapeEditCmdType, AddInClientId, , , smallIcon, largeIcon)
						Case "ObsoletePrint"
							obsoleteButton = controlDefs.AddButtonDefinition _
								("Obsolete" & Chr(10) & "Print",
								 "ObsoletePrint",
								 CommandTypesEnum.kNonShapeEditCmdType,
								 AddInClientId,
								 ,
								 ,
								 smallIcon,
								 largeIcon)
					End Select
				Next
		End Select

		' Always add to the user interface (not just the first time)
		' This ensures buttons appear when add-in is reloaded
		AddToUserInterface()

		' Connect to the user-interface events to handle a ribbon reset.
		uiEvents = ThisApplication.UserInterfaceManager.UserInterfaceEvents

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
				Dim result = MessageBox.Show _
					    ($"A new version of the Doyle AddIn is available ({latestVersion}) . Update now?",
					     "Update Available",
					     MessageBoxButtons.YesNo,
					     MessageBoxIcon.Question)
				If result = DialogResult.Yes Then
					File.WriteAllText("C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update")
					ThisApplication.Quit()
				Else
					File.WriteAllText("C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin\pending_update.txt", "update")
					MessageBox.Show _
						("The update will be installed after you close Inventor.",
						 "Update Scheduled",
						 MessageBoxButtons.OK,
						 MessageBoxIcon.Information)
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

		' Clean up button definitions
		Try
			If printUpdate IsNot Nothing Then
				printUpdate.Delete()
				printUpdate = Nothing
			End If
		Catch
		End Try

		Try
			If dxfUpdate IsNot Nothing Then
				dxfUpdate.Delete()
				dxfUpdate = Nothing
			End If
		Catch
		End Try

		Try
			If optionsButton IsNot Nothing Then
				optionsButton.Delete()
				optionsButton = Nothing
			End If
		Catch
		End Try

		Try
			If obsoleteButton IsNot Nothing Then
				obsoleteButton.Delete()
				obsoleteButton = Nothing
			End If
		Catch
		End Try

		' Release objects.
		uiEvents = Nothing
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
	Private Shared ReadOnly Item4Array2 As String() = New String() {"id_PanelP_ToolsOptions"}
	Private Shared ReadOnly Item4Array3 As String() = New String() {"id_PanelP_FlatPatternExit"}
	Private Shared ReadOnly Item4Array4 As String() = New String() {"id_PanelD_AnnotateRevision"}

	Public Sub ExecuteCommand(commandId As Integer) Implements ApplicationAddInServer.ExecuteCommand
	End Sub

#End Region

#Region "User interface definition"
	' Sub where the user-interface creation is done.  This is called when
	' the add-in loaded and also if the user interface is reset.
	Private Sub AddToUserInterface()
		' Load user options to check feature flags
		Dim options = UserOptions.Load()

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
			    Tuple.Create("id_TabSheetMetal", "dxfUpdate", dxfUpdate, Item4),
			    Tuple.Create("id_TabFlatPattern", "dxfUpdate", dxfUpdate, Item4Array3),
			    Tuple.Create("id_TabTools", "dxfUpdate", dxfUpdate, Item4Array2)
			    }

		' Option button appears on all document types
		Dim optionsButtonConfigs = New List(Of Tuple(Of String, String, ButtonDefinition, String())) From {
			    Tuple.Create("id_TabPlaceViews", "printUpdate", printUpdate, Item4Array0),
			    Tuple.Create("id_TabAnnotate", "printUpdate", printUpdate, Item4Array0),
			    Tuple.Create("id_TabTools", "userOptions", optionsButton, Item4Array)
			    }

		' Obsolete Print button - only add if feature is enabled
		Dim obsoletePrintConfigs = New List(Of Tuple(Of String, String, ButtonDefinition, String()))()
		If options.EnableObsoletePrint Then
			obsoletePrintConfigs.Add(Tuple.Create("id_TabAnnotate", "ObsoletePrint", obsoleteButton, Item4Array4))
		End If

		' Add buttons to appropriate ribbons based on document context
		For Each kvp In ribbonMappings
			Dim ribbonName = kvp.Key
			Dim ribbon = kvp.Value

			' Add the DXF button only to Part ribbon
			If ribbonName = "Part" Then
				For Each config In dxfButtonConfigs
					Dim tabName = config.Item1
					Dim panelName = config.Item2
					Dim buttonDef = config.Item3
					Dim fallbackPanels = config.Item4

					AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels)
				Next
			End If

			' Add Options buttons to all ribbons
			For Each config In optionsButtonConfigs
				Dim tabName = config.Item1
				Dim panelName = config.Item2
				Dim buttonDef = config.Item3
				Dim fallbackPanels = config.Item4

				AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels)
			Next

			' Add the Obsolete Print button to Drawing ribbon only if enabled
			If ribbonName = "Drawing" Then
				For Each config In obsoletePrintConfigs
					Dim tabName = config.Item1
					Dim panelName = config.Item2
					Dim buttonDef = config.Item3
					Dim fallbackPanels = config.Item4

					AddButtonToRibbon(ribbon, tabName, panelName, buttonDef, fallbackPanels)
				Next
			End If
		Next
	End Sub

	' Helper method to add a button to specific ribbon with fallback handling
	Private Shared Sub AddButtonToRibbon _
		(ribbon As Ribbon, tabName As String, panelName As String, buttonDef As ButtonDefinition, fallbackPanels As String())
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

			' Check if the button already exists in this panel
			If panel IsNot Nothing Then
				Dim buttonExists As Boolean = False
				For Each ctrl As CommandControl In panel.CommandControls
					Try
						' Add null checks before accessing properties
						If _
							ctrl IsNot Nothing AndAlso ctrl.ControlDefinition IsNot Nothing AndAlso buttonDef IsNot Nothing AndAlso
							ctrl.ControlDefinition.InternalName = buttonDef.InternalName Then
							buttonExists = True
							Exit For
						End If
					Catch
						' Skip this control if there's any issue
						Continue For
					End Try
				Next

				' Add the button only if it doesn't already exist
				If Not buttonExists Then
					panel.CommandControls.AddButton(buttonDef, True)
				End If
			End If

		Catch ex As Exception
			' Log specific errors for debugging
			Debug.Print _
				($"Failed to add button '{If(buttonDef IsNot Nothing, buttonDef.DisplayName, "Unknown")}' to tab '{tabName _
					}' in ribbon '{ribbon.InternalName}': {ex.Message}")
		End Try
	End Sub

	Private Sub UiEvents_OnResetRibbonInterface(context As NameValueMap) Handles uiEvents.OnResetRibbonInterface
		AddToUserInterface()
	End Sub

	Private Shared Sub DXFUpdate_OnExecute(context As NameValueMap) Handles dxfUpdate.OnExecute
		Call Sub() RunDxfUpdate(ThisApplication)
	End Sub

	Private Shared Sub PrintUpdate_OnExecute(context As NameValueMap) Handles printUpdate.OnExecute
		Call Sub() RunPrintUpdate(ThisApplication)
	End Sub

	Private Sub OptionsButton_OnExecute(context As NameValueMap) Handles optionsButton.OnExecute
		Dim optionsForm As New UserOptionsForm()
		Dim result = optionsForm.ShowDialog(New WindowWrapper(ThisApplication.MainFrameHWND))

		' Refresh the ribbon after options are saved
		If result = DialogResult.OK Then
			RefreshRibbon()
		End If
	End Sub

	Private Sub ObsoleteButton_OnExecute(context As NameValueMap) Handles obsoleteButton.OnExecute
		Call Sub() ApplyObsoletePrint(ThisApplication)
	End Sub

	' Helper method to refresh the ribbon UI
	Private Sub RefreshRibbon()
		Try
			' Remove the obsolete print button from all ribbons first
			RemoveObsoletePrintButton()

			' Re-add buttons with updated settings
			AddToUserInterface()

			'			MessageBox.Show("Ribbon updated successfully. Changes are now active.",
			'	"Settings Applied",
			'MessageBoxButtons.OK,
			'MessageBoxIcon.Information)'
		Catch ex As Exception
			'MessageBox.Show($"Error refreshing ribbon: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End Try
	End Sub

	' Helper method to remove the obsolete print button from ribbons
	Private Sub RemoveObsoletePrintButton()
		Try
			Dim uiManager = ThisApplication.UserInterfaceManager
			Dim ribbon = uiManager.Ribbons.Item("Drawing")

			If ribbon IsNot Nothing Then
				Try
					Dim tab = ribbon.RibbonTabs.Item("id_TabAnnotate")
					If tab IsNot Nothing Then
						' Find all panels and remove the obsolete print button
						For Each panel As RibbonPanel In tab.RibbonPanels
							Dim controlsToRemove As New List(Of CommandControl)

							' Collect controls to remove (with null checks)
							For Each ctrl As CommandControl In panel.CommandControls
								Try
									If _
										ctrl IsNot Nothing AndAlso ctrl.ControlDefinition IsNot Nothing AndAlso
										ctrl.ControlDefinition.InternalName = "ObsoletePrint" Then
										controlsToRemove.Add(ctrl)
									End If
								Catch
									' Skip this control if there's any issue accessing it
									Continue For
								End Try
							Next

							' Remove collected controls
							For Each ctrl In controlsToRemove
								Try
									ctrl.Delete()
								Catch ex As Exception
									Debug.Print($"Failed to delete control: {ex.Message}")
								End Try
							Next
						Next
					End If
				Catch ex As Exception
					' Button might not exist, continue
					Debug.Print($"Could not remove button from Annotate tab: {ex.Message}")
				End Try
			End If
		Catch ex As Exception
			' Log error but don't show to the user during refresh
			Debug.Print($"Error removing obsolete print button: {ex.Message}")
		End Try
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

		Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
	End Class

#End Region
End Module
