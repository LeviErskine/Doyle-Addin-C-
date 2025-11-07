

Public Function nuFso() As Scripting.FileSystemObject
    Set nuFso = New Scripting.FileSystemObject
End Function

Public Function fdUserHome() As Scripting.Folder
    Set fdUserHome = nuFso.GetFolder(Environ("AppData")).ParentFolder
End Function

Public Function fnExt(fn As String) As String
    Dim ar As Variant
    Dim mx As Long
    
    ar = Split(fn, ".")
    mx = UBound(ar)
    If mx < 0 Then
        fnExt = ""
    Else
        fnExt = ar(mx)
    End If
End Function

Public Function folderIfPresent(Path As String, _
    Optional Base As Scripting.Folder = Nothing _
) As Scripting.Folder
    Dim rt As Scripting.Folder
    
    If Base Is Nothing Then
        With nuFso()
            If .FolderExists(Path) Then
                Set rt = .GetFolder(Path)
            Else
                Set rt = Nothing
            End If
        End With
    Else
        Set rt = folderIfPresent(Base.Path & "\" & Path)
    End If
    
    Set folderIfPresent = rt
End Function

Public Function fileIfPresent(Path As String, _
    Optional Base As Scripting.Folder = Nothing _
) As Scripting.File
    Dim rt As Scripting.File
    
    If Base Is Nothing Then
        With nuFso()
            If .FileExists(Path) Then
                Set rt = .GetFile(Path)
            Else
                Set rt = Nothing
            End If
        End With
    Else
        Set rt = fileIfPresent(Base.Path & "\" & Path)
    End If
    
    Set fileIfPresent = rt
End Function

Public Function dcFilesIn(fd As Scripting.Folder) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim fl As Scripting.File
    
    Set rt = New Scripting.Dictionary
    If Not fd Is Nothing Then
        For Each fl In fd.Files
            rt.Add fl.Name, fl
        Next
    End If
    Set dcFilesIn = rt
End Function
'send2clipBd dcFilesIn(nuFso.GetFolder(""))
'send2clipBd Join(dcFoldersIn(nuFso.GetFolder("C:\Doyle_Vault\Designs\doyle")).Keys, vbNewLine)
'send2clipBd Join(dcFilesIn(nuFso.GetFolder("W:\Parts Lists")).Keys, vbNewLine)

Public Function dcFoldersIn( _
    fd As Scripting.Folder _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim fl As Scripting.Folder
    
    Set rt = New Scripting.Dictionary
    If Not fd Is Nothing Then
        For Each fl In fd.SubFolders
            rt.Add fl.Name, fl
        Next
    End If
    Set dcFoldersIn = rt
End Function
