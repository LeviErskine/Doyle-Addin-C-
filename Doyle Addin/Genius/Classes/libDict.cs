

Public Function dcNewIfNone( _
    Dict As Scripting.Dictionary _
) As Scripting.Dictionary
    If Dict Is Nothing Then
        Set dcNewIfNone = _
        New Scripting.Dictionary
    Else
        Set dcNewIfNone = Dict
    End If
End Function

Public Function dcOfRsFields( _
    rs As ADODB.Recordset _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim fd As ADODB.Field
    
    Set rt = New Scripting.Dictionary
    With rs
        If .State = adStateOpen Then
        For Each fd In .Fields
            rt.Add fd.Name, fd
        Next
        End If
    End With
    Set dcOfRsFields = rt
End Function

Public Function dcDotted( _
    Optional Under As Scripting.Dictionary = Nothing, _
    Optional Using As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcDotted -- return Dictionary with
    '''     links to itself, under key ".",
    '''     and under "..", either itself,
    '''     or, if supplied, an optional
    '''     "parent" Dictionary
    '''
    ''' this mimics the traditional linkage within
    '''     POSIX-compliant and other file systems,
    '''     where the "." and ".." names in each
    '''     directory are assigned to itself
    '''     and its parent, respecrively
    '''
    ''' !!WARNING!!
    '''     this self- and back-linkage WILL cause
    '''     endless loops in Dictionary traversal
    '''     routines not prepared to deal with them!
    '''     Be sure to review any procedure BEFORE
    '''     calling against a Dictionary using
    '''     this linkage!
    '''
    Dim rt As Scripting.Dictionary
    
    Set rt = dcNewIfNone(Using)
    With rt
        If .Exists(".") Then
            If .Item(".") Is rt Then
            Else
                Stop
            End If
        Else
            .Add ".", rt
        End If
        
        If .Exists("..") Then
            If TypeOf .Item("..") Is Scripting.Dictionary Then
                If Using Is Nothing Then
                Else
                    Stop
                    Set .Item("..") = Using
                End If
            Else
                Stop
            End If
        Else
            .Add "..", IIf( _
                Under Is Nothing, _
                rt, Under _
            )
        End If
    End With
    
    Set dcDotted = rt
End Function

Public Function dcUnDotted( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcUnDotted -- remove Keys "." and ".."
    '''     from supplied Dictionary dc
    '''
    '''     no checks are made of the Items under these Keys.
    '''     the Dictionary is assumed to have originated from
    '''     or passed through a prior call to dcDotted, and
    '''     thus include self- and back-linkage thereunder.
    '''
    '''     (a check system was considered and attempted,
    '''      but deemed too unweildy, and so abandoned)
    '''
    With dc
        If .Exists(".") Then .Remove "."
        If .Exists("..") Then .Remove ".."
    End With
    Set dcUnDotted = dc
End Function

Public Function dcFrom2Fields( _
    rs As ADODB.Recordset, _
    fnKey As String, fnVal As String, _
    Optional flt As String = "" _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim fdKey As ADODB.Field
    Dim fdVal As ADODB.Field
    
    Set rt = New Scripting.Dictionary
    With rs
        With .Fields
            Set fdKey = .Item(fnKey)
            Set fdVal = .Item(fnVal)
        End With
        
        .Filter = flt
        Do Until .BOF Or .EOF
            With rt
                If .Exists(fdKey.Value) Then
                    Stop
                Else
                    .Add fdKey.Value, fdVal.Value
                End If
            End With
            .MoveNext
        Loop
    End With
    Set dcFrom2Fields = rt
End Function

Public Function dcFromAdoRS( _
    rs As ADODB.Recordset, _
    Optional flt As String = "" _
) As Scripting.Dictionary
', fnKey As String, fnVal As String
', Optional ovr As Long = -1
    '''
    ''' dcFromAdoRS - return a Dictionary
    '''     of tuples (rows) from an ADODB
    '''     Recordset, keyed on order of
    '''     encounter and processing.
    '''
    ''' NOTE that this Dictionary is NOT
    '''     keyed on any particular Field.
    '''     The wide range of situations which
    '''     might be encountered suggests that
    '''     indexing and keying on field values
    '''     is best addressed in a separate,
    '''     dedicated process.
    '''
    Dim rt As Scripting.Dictionary
    Dim tp As Scripting.Dictionary
    'Dim fdKey As ADODB.Field
    Dim fdVal As ADODB.Field
    Dim ky As Variant
    Dim vl As Variant
    Dim nm As String
    
    Set rt = New Scripting.Dictionary
    With rs
        'With .Fields
        '    Set fdKey = .Item(fnKey)
        'End With
        
        .Filter = flt
        Do Until .BOF Or .EOF
            ky = rt.Count 'fdKey.Value
            With rt
            '    If .Exists(ky) Then 'we have a collision!
            '        Stop 'and figure out what to do!
            '    Else
                    '.Add ky, dcFromAdoRSrow(rs)
                    .Add ky, New Scripting.Dictionary
            '    End If
                Set tp = dcOb(.Item(ky))
                'tp.Add "*ROW*", ky
            End With
            
            For Each fdVal In .Fields
                With fdVal
                    nm = .Name
                    vl = .Value
                End With
                
                With tp
                'If .Exists(nm) Then
                '    If ovr Then 'change if needed
                '        If .Item(nm) <> vl Then
                '            .Item(nm) = vl
                '        End If
                '    Else 'fuhgeddaboudit!
                '    End If
                'Else
                    .Add nm, vl
                'End If
                End With
            Next
            
            .MoveNext
        Loop
    End With
    Set dcFromAdoRS = rt
End Function

Public Function dcFromAdoRSrow( _
    rs As ADODB.Recordset, _
    Optional nullVal As Variant = Null _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim fd As ADODB.Field
    Dim nm As String
    Dim ck As Boolean
    
    Set rt = New Scripting.Dictionary
    With rs
        ck = .BOF Or .EOF
        For Each fd In .Fields
            With fd
                nm = .Name
                If ck Then
                    rt.Add nm, nullVal
                Else
                    rt.Add nm, .Value
                End If
            End With
        Next
    End With
    Set dcFromAdoRSrow = rt
End Function

Public Function dcDxFromRecSetDc( _
    Dict As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcDxFromRecSetDc -- Generate Dictionary
    '''     of Indices from "RecordSet" Dictionary
    '''     as returned by dcFromAdoRS
    '''
    Dim tp As Scripting.Dictionary
    Dim dcDx As Scripting.Dictionary
    Dim dcVl As Scripting.Dictionary
    Dim dcTp As Scripting.Dictionary
    ''
    Dim k0 As Variant
    Dim k1 As Variant
    Dim k2 As Variant
    ''
    Dim vl As Variant
    
    'Set rt = New Scripting.Dictionary
    
    Set dcDx = New Scripting.Dictionary
    ''  the Dictionary of Indices
    
    With Dict
        ''  Start scanning primary Keys
        ''  to begin overall process
        For Each k0 In .Keys
            ''  Retrieve "record" Dictionary
            ''  for next/current Key
            Set tp = dcOb(.Item(k0))
            If tp Is Nothing Then
                'Stop
                Debug.Print ; 'Breakpoint Landing
            Else
                With tp
                    ''  Scan "field" Keys of current "record"
                    ''  to identify and populate Index Dictionaries
                    For Each k1 In .Keys
                        ''  Retrieve "index" Dictionary for current
                        ''  "field". Generate new one, if not present.
                        ''
                        ''  (might want to support Key filtering
                        ''  to either exclude some "fields",
                        ''  or limit indexing to a list)
                        With dcDx
                            If .Exists(k1) Then
                            Else
                                .Add k1, New Scripting.Dictionary
                            End If
                            Set dcVl = dcOb(.Item(k1))
                        End With
                        
                        ''  Retrieve current "field" value, and return its
                        ''  Dictionary from the "field index" Dictionary.
                        ''
                        ''  Again, generate a new one, if needed.
                        vl = .Item(k1)
                        With dcVl
                            If .Exists(vl) Then
                            Else
                                .Add vl, New Scripting.Dictionary
                            End If
                            Set dcTp = dcOb(.Item(vl))
                        End With
                        
                        ''  Add the current "record" to the recovered
                        ''  "field value" Dictionary. This SHOULD only
                        ''  add a link to the same "record" Dictionary,
                        ''  rather than duplicate the whole thing.
                        ''
                        ''  However, converting to JSON generates
                        ''  a new dump of the Dictionary structure
                        ''  wherever it appears, thus replicating it
                        ''  multiple times in the output.
                        ''
                        With dcTp
                            If .Exists(k0) Then 'might be a problem
                                Stop 'for now. might still be okay
                            Else
                                .Add k0, tp
                            End If
                            
                            DoEvents
                        End With
                        ''  Initial plans were to replace this segment
                        ''  and the "key value" Dictionary with a simple
                        ''  String of matching "record" Keys to be used
                        ''  for lookup within a single Dictionary in
                        ''  the JSON document.
                        ''
                        ''  However, current thinking suggests that
                        ''  keeping these Dictionaries facilitates
                        ''  more efficient access and use in live
                        ''  operation.
                        ''
                        ''  Conversion to more efficient JSON would
                        ''  be better supported by replacement of
                        ''  these Dictionaries with their Key lists
                        ''  immediately prior to export. Naturually,
                        ''  this replacement would occur in a COPY
                        ''  of the live Dictionary tree.
                    Next
                End With
            End If
        Next
    End With
    
    With dcDx
        If .Exists("") Then
            Stop 'because we have
            'an "anonymous" field index
        Else
            .Add "", Dict 'to provide access
            'to original Dictionary, and ensure
            'easy lookup by original keys
        End If
    End With
    
    Set dcDxFromRecSetDc = dcDx
End Function

Public Function dcRecSetDcDx4json( _
    Dict As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcRecSetDcDx4json -- Prep RecordSet
    '''     Index Dictionary for JSON export.
    '''
    ''' Replaces each field/value index
    ''' Dictionary with its Keys for export
    ''' to JSON, to avoid replicating each
    ''' original "record" Item in its entirety
    ''' wherever it's referenced in the indices.
    '''
    Dim rt As Scripting.Dictionary
    Dim dcFdIn As Scripting.Dictionary
    Dim dcFdOut As Scripting.Dictionary
    Dim dcVl As Scripting.Dictionary
    ''
    Dim k0 As Variant
    Dim vl As Variant
    
    Set rt = New Scripting.Dictionary
    ''  the Dictionary of Indices
    
    With Dict
        ''  Start scanning field
        ''  names (top level Keys)
        ''  to begin transformation
        For Each k0 In .Keys
            ''  Retrieve next "field index" Dictionary
            Set dcFdIn = dcOb(.Item(k0))
            
            ''  Check for original RecordSet Dictionary
            If k0 = "" Then 'copy it over straight
                rt.Add "", dcFdIn
            Else
                ''  Generate corresponding "field index"
                ''  output Dictionary
                With rt
                    If .Exists(k0) Then
                        Stop 'because it should NOT
                             'be there already!
                    Else
                        .Add k0, New Scripting.Dictionary
                    End If
                    
                    Set dcFdOut = dcOb(.Item(k0))
                End With
                
                ''  Scan value Keys of current "field"
                ''  to retrieve index Dictionaries
                With dcFdIn: For Each vl In .Keys
                    ''  Retrieve Dictionary for current value
                    Set dcVl = dcOb(.Item(vl))
                    
                    With dcFdOut
                    If .Exists(vl) Then
                        Stop 'because it should
                             'NOT already be there!
                    Else
                        ''  Dump record Keys to output
                        ''  field value index Dictionary
                        .Add vl, dcVl.Keys
                    End If: End With
                Next: End With
            End If
        Next
    End With
    
    Set dcRecSetDcDx4json = rt
'send2clipBdWin10 ConvertToJson(dcRecSetDcDx4json( _
    dcDxFromRecSetDc(dcFromAdoRS( _
        cnGnsDoyle().Execute( _
            q1g1x2(aiDocActive()) _
        ) _
    )) _
), vbTab)
End Function

Public Function dcOfSubDict(dc As Scripting.Dictionary, _
    Optional rt As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcOfSubDict -- intended to return
    '''     a "flat" Dictionary containing
    '''     the supplied Dictionary and all
    '''     Dictionary objects within it.
    '''
    '''     DO NOT ATTEMPT TO USE AT THIS TIME!!!
    '''     Need to work out a way to tell
    '''     if the supplied Dictionary is already in the returned
    '''
    Dim ky As Variant
    
    If rt Is Nothing Then
        Set dcOfSubDict = dcOfSubDict( _
        dc, New Scripting.Dictionary)
    ElseIf dc Is Nothing Then
        Set dcOfSubDict = rt
    Else
        '''
    End If
End Function
