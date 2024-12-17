Imports System.Runtime.InteropServices
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

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            UiEvents = Nothing
            ThisApplication = Nothing

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
            Try
                Dim flatPatternPanel = partRibbon.RibbonTabs.Item("id_TabSheetMetal").RibbonPanels.Item("id_PanelP_SheetMetalManageUnfold")
                flatPatternPanel.CommandControls.AddButton(DXFUpdate, True)
            Catch ex As Exception
                ' Panel or tab may not exist, handle gracefully
            End Try

            '' Add Genius to Part tab
            'Try
            '    Dim sheetMetalTab = partRibbon.RibbonTabs.Item("id_TabSheetMetal")
            '    Dim geniusPanel As RibbonPanel = Nothing
            '    ' Check if panel already exists to avoid duplicates
            '    For Each panel As RibbonPanel In sheetMetalTab.RibbonPanels
            '        If panel.InternalName = "GeniusUpdate" Then
            '            geniusPanel = panel
            '            Exit For
            '        End If
            '    Next
            '    If geniusPanel Is Nothing Then
            '        geniusPanel = sheetMetalTab.RibbonPanels.Add("Genius", "GeniusUpdate", AddInClientID())
            '    End If
            '    geniusPanel.CommandControls.AddButton(GeniusUpdate, True)
            'Catch ex As Exception
            '    ' Handle missing tab or other errors
            'End Try

            ' Add Print Update to Drawing Place Views tab
            Try
                Dim placeViewsTab = drawingRibbon.RibbonTabs.Item("id_TabPlaceViews")
                Dim placeViewsPanel As RibbonPanel = Nothing
                For Each panel As RibbonPanel In placeViewsTab.RibbonPanels
                    If panel.InternalName = "printUpdate" Then
                        placeViewsPanel = panel
                        Exit For
                    End If
                Next
                If placeViewsPanel Is Nothing Then
                    placeViewsPanel = placeViewsTab.RibbonPanels.Add("Add-Ins", "printUpdate", AddInClientID())
                End If
                placeViewsPanel.CommandControls.AddButton(PrintUpdate, True)
            Catch ex As Exception
                ' Handle missing tab or other errors
            End Try

            ' Add Print Update to Drawing Annotate tab
            Try
                Dim annotateTab = drawingRibbon.RibbonTabs.Item("id_TabAnnotate")
                Dim annotatePanel As RibbonPanel = Nothing
                For Each panel As RibbonPanel In annotateTab.RibbonPanels
                    If panel.InternalName = "printUpdateAnnotate" Then
                        annotatePanel = panel
                        Exit For
                    End If
                Next
                If annotatePanel Is Nothing Then
                    annotatePanel = annotateTab.RibbonPanels.Add("Add-Ins", "printUpdateAnnotate", AddInClientID())
                End If
                annotatePanel.CommandControls.AddButton(PrintUpdate, True)
            Catch ex As Exception
                ' Handle missing tab or other errors
            End Try
        End Sub

        Private Sub UiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles UiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        Private Sub DXFUpdate_OnExecute(Context As NameValueMap) Handles DXFUpdate.OnExecute
            Call Sub() runDxfUpdate(ThisApplication)
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

    '#Region "Image Converter"
    '    ' Class used to convert bitmaps and icons from their .Net native types into
    '    ' an IPictureDisp object which is what the Inventor API requires. A typical
    '    ' usage is shown below where MyIcon is a bitmap or icon that's available
    '    ' as a resource of the project.
    '    '
    '    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    '    Public NotInheritable Class PictureDispConverter
    '        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
    '        Private Shared Function OleCreatePictureIndirect(
    '    <MarshalAs(UnmanagedType.Struct)> ByVal picdesc As Object,
    '    ByRef iid As GUID,
    '    <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
    '        End Function

    '        Shared iPictureDispGuid As GUID = GetType(stdole.IPictureDisp).GUID

    '        Private NotInheritable Class PICTDESC
    '            Private Sub New()
    '            End Sub

    '            'Picture Types
    '            Public Const PICTYPE_BITMAP As Short = 1
    '            Public Const PICTYPE_ICON As Short = 3

    '            <StructLayout(LayoutKind.Sequential)>
    '            Public Class Icon
    '                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
    '                Friend picType As Integer = PICTDESC.PICTYPE_ICON
    '                Friend hicon As IntPtr = IntPtr.Zero
    '                Friend unused1 As Integer
    '                Friend unused2 As Integer

    '                Friend Sub New(ByVal icon As System.Drawing.Icon)
    '                    Me.hicon = icon.ToBitmap().GetHicon()
    '                End Sub
    '            End Class

    '            <StructLayout(LayoutKind.Sequential)>
    '            Public Class Bitmap
    '                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
    '                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
    '                Friend hbitmap As IntPtr = IntPtr.Zero
    '                Friend hpal As IntPtr = IntPtr.Zero
    '                Friend unused As Integer

    '                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
    '                    Me.hbitmap = bitmap.GetHbitmap()
    '                End Sub
    '            End Class
    '        End Class

    '        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
    '            Dim pictIcon As New PICTDESC.Icon(icon)
    '            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
    '        End Function

    '        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
    '            Dim pictBmp As New PICTDESC.Bitmap(bmp)
    '            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
    '        End Function
    '    End Class
    '#End Region

End Module
