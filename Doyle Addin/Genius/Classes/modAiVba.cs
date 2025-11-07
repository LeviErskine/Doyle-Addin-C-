
'' Will be using this one to work on VB Extensibility code

Public Function m2g0f0() As Long
    Dim pj As Inventor.InventorVBAProject
    
    With ThisApplication
        For Each pj In .VBAProjects
            Debug.Print pj.VBProject '.Name
        Next
        m2g0f0 = .VBAProjects.Count
    End With
End Function

Public Function m2g1f0() As Inventor.InventorVBAProject
    Set m2g1f0 = ThisApplication.VBAProjects.Item(1)
End Function

Public Function fnOfDefaultVBAproject() As String
    fnOfDefaultVBAproject = ThisApplication.FileOptions.DefaultVBAProjectFileFullFilename
End Function

Public Function m2g1f2(ob As Inventor.InventorVBAProject) As VBIDE.VBProject
    Set m2g1f2 = ob.VBProject
End Function
'Debug.Print m2g1f2(dcInVBAprojects(ThisApplication).Item(fnOfDefaultVBAproject)).BuildFileName


Public Function dcInVBAprojects(ap As Inventor.Application) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pj As Inventor.InventorVBAProject
    Dim mx As Long
    Dim dx As Long
    
    Set rt = New Scripting.Dictionary
    With ap.VBAProjects
        mx = .Count
        For dx = 1 To mx
            Set pj = .Item(dx)
            rt.Add m2g1f2(pj).Filename, pj
        Next
    End With
    Set dcInVBAprojects = rt
End Function

