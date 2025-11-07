

Public Function fd0g1f1(ls As Variant) As Variant
    Dim fm As fmEmpty
    Dim lb As MSForms.ListBox
    Dim mg As Single
    Dim wk As Variant
    Dim dc As Scripting.Dictionary
    Dim dx As Long
    Dim rt As String
    
    If IsArray(ls) Then
        wk = ls
    ElseIf IsObject(ls) Then
        Set dc = dcOb(obOf(ls))
        If dc Is Nothing Then
            wk = Empty
        Else
            wk = dcOb(ls).Keys
        End If
    Else
        wk = Empty
    End If
    
    If IsEmpty(wk) Then wk = Array( _
        "*vvvvvvvvvvv*", _
        "*Unsupported*", _
        "*List Source*", _
        "*^^^^^^^^^^^*" _
    )
    
    Set fm = nuFmEmpty()
    mg = 10
    
    Set lb = nuMsFmCtListBox(fm, , "lbxA")
    With obMsFmControl(lb)
        .Top = mg
        .Left = mg
        .Height = fm.InsideHeight - mg - mg
        .Width = fm.InsideWidth - mg - mg
    End With
    
    With lb
        .MultiSelect = fmMultiSelectMulti ' fmMultiSelectExtended
        .ListStyle = fmListStyleOption
        
        .List = wk
        fm.Show vbModal
        
        'Stop
        rt = ""
        dx = 0
        Do Until dx = .ListCount
            If .Selected(dx) Then
            rt = rt & vbVerticalTab & .List(dx)
            End If
            dx = 1 + dx
        Loop
    End With
    
    fd0g1f1 = Split(Mid$(rt, 2), vbVerticalTab)
End Function

Public Function nuFmEmpty(Optional f As Variant) As fmEmpty
    With New fmEmpty
        '''
        Set nuFmEmpty = .Itself
    End With
End Function

Public Function obMsFmControl(it As Variant) As MSForms.Control
    Dim ob As Object
    
    Set ob = obOf(it)
    If TypeOf ob Is MSForms.Control Then
        Set obMsFmControl = ob
    Else
        Set obMsFmControl = Nothing
    End If
End Function

Public Function nuMsFmCtListBox(fm As fmEmpty, _
    Optional sp As Variant = Empty, _
    Optional nm As Variant = "", _
    Optional vs As Boolean = True _
) As MSForms.ListBox 'MSForms.UserForm
    '''
    ''' nuMsFmCtListBox -- add new ListBox
    '''     Control to supplied fmEmpty Object
    '''
    '''     accepts, but does not yet use,
    '''     a specification sp laying out
    '''     the parameters defining the
    '''     desired control
    '''
    '''     note that fm MUST be fmEmpty
    '''     a general MSForms UserForm is
    '''     NOT accepted, because it does
    '''     NOT support certain essential
    '''     properties; for example, there
    '''     are not properties to set size
    '''     or position, which are essential
    '''     to the goals of this system
    '''
    Dim rt As MSForms.ListBox
    'Dim ct As MSForms.Control
    
    'fm.Left
    
    Set rt = fm.Controls.Add( _
        "Forms.ListBox.1", , vs _
    )
    If Len(nm) > 0 Then
        obMsFmControl(rt).Name = nm
        'Set ct = rt
        'ct.Name = nm
    End If
    
    With rt
    End With
    
    Set nuMsFmCtListBox = rt
End Function
