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

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
            Dim PrintUpdateIconLarge As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.PrintUpdateIconLarge)
            Dim PrintUpdateIconSmall As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.PrintUpdateIconSmall)
            Dim DXFUpdateIconSmall As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.DXFUpdateIconSmall)
            Dim DXFUpdateIconLarge As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.DXFUpdateIconLarge)
            DXFUpdate = controlDefs.AddButtonDefinition("DXF Update", "dxfUpdate", CommandTypesEnum.kShapeEditCmdType, AddInClientID, , , DXFUpdateIconSmall, DXFUpdateIconLarge)
            PrintUpdate = controlDefs.AddButtonDefinition("Print Update", "printUpdate", CommandTypesEnum.kShapeEditCmdType, AddInClientID, , , PrintUpdateIconSmall, PrintUpdateIconLarge)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If

            ' Connect to the user-interface events to handle a ribbon reset.
            UiEvents = ThisApplication.UserInterfaceManager.UserInterfaceEvents

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
            Dim partRibbon = uiManager.Ribbons.Item("Part")
            Dim drawingRibbon = uiManager.Ribbons.Item("Drawing")

            ' Add DXF Update to the existing Flat Pattern panel on the Sheet Metal Tools tab
            Dim DXFTabs = New Dictionary(Of String, String) From {
                {"id_TabSheetMetal", "dxfUpdate"},
                {"id_TabFlatPattern", "dxfUpdate"},
                {"id_TabTools", "dxfUpdate"}
            }

            For Each tabInfo In DXFTabs
                Try
                    Dim tab As RibbonTab = Nothing
                    Try
                        tab = partRibbon.RibbonTabs.Item(tabInfo.Key)
                    Catch ex As Exception
                        Debug.Print("Tab not found: " & tabInfo.Key)
                        Continue For
                    End Try

                    Dim panel As RibbonPanel = Nothing
                    Try
                        panel = tab.RibbonPanels.Item(tabInfo.Value)
                    Catch
                        ' Try fallback panels
                        Try
                            Select Case tabInfo.Key
                                Case "id_TabSheetMetal"
                                    panel = tab.RibbonPanels.Item("id_PanelP_SheetMetalManageUnfold")
                                Case "id_TabFlatPattern"
                                    panel = tab.RibbonPanels.Item("id_PanelP_FlatPatternExit")
                                Case "id_TabTools"
                                    panel = partRibbon.RibbonTabs.Item("id_TabTools").RibbonPanels.Item("id_PanelP_ToolsOptions")
                            End Select
                        Catch ex As Exception
                            Debug.Print("Panel not found for tab: " & tabInfo.Key)
                            Continue For
                        End Try
                    End Try

                    panel?.CommandControls.AddButton(DXFUpdate, True)
                Catch ex As Exception
                    Debug.Print("Unexpected error: " & ex.Message)
                End Try
            Next


            ' Add Print Update to Drawing Place Views tab

            Dim PrintTabs = New Dictionary(Of String, String) From {
                    {"id_TabPlaceViews", "printUpdate"},
                    {"id_TabAnnotate", "printUpdateAnnotate"}
                }
            For Each tabInfo In PrintTabs
                Try
                    Dim tab = drawingRibbon.RibbonTabs.Item(tabInfo.Key)
                    Dim panel As RibbonPanel = Nothing
                    For Each p As RibbonPanel In tab.RibbonPanels
                        If p.InternalName = tabInfo.Value Then
                            panel = p
                            Exit For
                        End If
                    Next
                    If panel Is Nothing Then
                        panel = tab.RibbonPanels.Add("Add-Ins", tabInfo.Value, AddInClientID())
                    End If
                    panel.CommandControls.AddButton(PrintUpdate, True)
                Catch ex As Exception
                    ' Handle missing tab or other errors
                End Try

            Next
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
