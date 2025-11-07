

Implements ifcDatum

'Private ps As Inventor.PropertySet
Private pr As Inventor.Property
'Private nm As String
Private vlWas As Variant
Private vlNow As Variant

Public Function Connect( _
    ToProp As Inventor.Property _
) As ifcDatum
    Set pr = ToProp
    If ToProp Is Nothing Then
        'Set ps = Nothing
        'nm = ""
        vlWas = Empty
        ''  not sure what's best here
        ''  be prepared to change
        ''  as necessary
    Else
        With pr
        'Set ps = .Parent
        'nm = .Name
        vlWas = .Value
        End With
    End If
    vlNow = vlWas
    
    Set Connect = Me
End Function
''' replaces disabled function below
'
'Public Function AttachedTo(Name As String, _
'    Optional InPropSet As Inventor.PropertySet = Nothing _
') As ifcAiProperty
'    '''
'    '''
'    '''
'    nm = Name
'
'    If Not InPropSet Is Nothing Then
'        Set ps = InPropSet
'    End If
'
'    If Not ps Is Nothing Then
'        On Error Resume Next
'        Set pr = ps.Item(nm)
'        If Err.Number = 0 Then
'            vlWas = pr.Value
'        Else
'            Set pr = Nothing
'            vlWas = Empty
'        End If
'        On Error GoTo 0
'    End If
'
'    Set AttachedTo = Me
'End Function

Public Function MakeValue( _
    This As Variant _
) As ifcDatum
    If IsObject(This) Then
        ''' really should NOT support this
        ''' but will let stand for now
        'Set vlNow = This
        ''' no. will opt for
        ''' 'do nothing' instead
    Else
        vlNow = This
    End If
    
    Set MakeValue = Me
End Function
''' replaces disabled function below
'
'Public Function WithValue( _
'    NewVal As Variant _
') As ifcAiProperty
'    Me.Value = NewVal
'
'    Set WithValue = Me
'End Function

Public Function Commit() As ifcAiProperty
    'Dim ps As Inventor.PropertySet
    'Dim ck As Variant
    
    If IsEmpty(vlWas) Then
        Stop
        'don't know about clearing Property
    Else
        If pr Is Nothing Then
            'If ps Is Nothing Then
                'can't do anything
            'Else
                'Set pr = ps.Add(vlWas, nm)
                'If pr Is Nothing Then
                '    Set Commit = Me
                'Else
                '    Set Commit = Commit()
                'End If
            'End If
        Else
            vlWas = pr.Value 'ck
            
            If vlNow = vlWas Then 'ck
                'shouldn't need updated
            ElseIf CStr(vlNow) = CStr(vlWas) Then 'ck
                'PROBABLY shouldn't need updated
            Else
                On Error Resume Next
                
                pr.Value = vlNow 'vlWas
                If Err.Number = 0 Then 'should be okay
                    'don't worry about it
                Else 'something went wrong
                    Stop 'and see what we can do
                End If
                
                On Error GoTo 0
            End If
        End If
    End If
    
    Set Commit = Me
End Function

Private Function Itself() As ifcDatum
    Set Itself = Me
End Function
''' replaces disabled function below
'
'Public Function Obj() As ifcAiProperty
'    Set Obj = Me
'End Function

Private Function Connected( _
    Optional ToThis As Inventor.Property = Nothing _
) As Boolean
    If ToThis Is Nothing Then
        Connected = Not pr Is Nothing
    Else
        Connected = pr Is ToThis
    End If
End Function

Private Function Value() As Variant
    If IsObject(vlNow) Then
        ''' this should NOT ever happen
        ''' but just to be robust...
        Set Value = vlNow
    Else
        Value = vlNow
    End If
End Function

Public Function Status() As Long
    Status = -1
End Function

Public Function Name() As Variant
    If pr Is Nothing Then
        Name = pr.Name
    Else
        Name = ""
    End If
End Function

'Public Property Get Value() As Variant
'    Value = vlWas
'End Property
'
'Public Property Let Value(NewVal As Variant)
'    If IsEmpty(NewVal) Then
'        Stop
'    'ElseIf IsNull(NewVal) Then
'    'ElseIf IsMissing(NewVal) Then
'    ElseIf IsObject(NewVal) Then
'        Stop
'    Else
'        vlWas = NewVal
'    End If
'End Property

Private Sub Class_Initialize()
    'nm = ""
    vlWas = Empty
    'Set ps = Nothing
    Set pr = Nothing
End Sub

Private Sub Class_Terminate()
    'If ps Is Nothing Then 'nowhere to save
        'so nothing to do but drop it
    'Else
    If pr Is Nothing Then 'we need
        'to create Property, if desired
        '(and possible)
        Stop
    ElseIf vlWas = pr.Value Then 'no change
        'so nothing needs doing
    ElseIf CStr(vlWas) = CStr(pr.Value) Then
        'likely no REAL change
    Else 'value has changed
        'and MIGHT need to be committed
        Stop
    End If
End Sub

Private Function ifcDatum_Commit() As ifcDatum
    '''
End Function

Private Function ifcDatum_Connect( _
    ToThis As Object _
) As ifcDatum
    Set ifcDatum_Connect = _
    Connect(obAiProp(ToThis))
End Function

Private Function ifcDatum_Connected( _
    Optional ToThis As Object = Nothing _
) As Boolean
    ifcDatum_Connected = _
    Connected(obAiProp(ToThis))
End Function

Private Function ifcDatum_Itself() As ifcDatum
    Set ifcDatum_Itself = Me
End Function

Private Function ifcDatum_MakeValue( _
    This As Variant _
) As ifcDatum
    Set ifcDatum_MakeValue _
    = MakeValue(This)
End Function

Private Function ifcDatum_Value() As Variant
    If IsObject(vlNow) Then
        ''' this should NOT ever happen
        ''' but just to be robust...
        Set ifcDatum_Value = vlNow
    Else
        ifcDatum_Value = vlNow
    End If
End Function
