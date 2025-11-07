

Public Function dcCopy( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcCopy -- return a new Dictionary
    '''     copying the contents of the
    '''     one supplied, including COPIES
    '''     of any Dictionary Objects within.
    '''
    Dim rt As Scripting.Dictionary
    Dim ck As Scripting.Dictionary
    Dim bx As Variant
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    If dc Is Nothing Then 'skip transfer
        'cuz there ain't Nothing in Nothing!
    Else 'we MIGHT have stuff to move
        With dc: For Each ky In .Keys
            bx = Array(.Item(ky))
            Set ck = dcOb(obOf(bx(0)))
            
            If ck Is Nothing Then
                rt.Add ky, bx(0)
            Else
                rt.Add ky, dcCopy(ck)
            End If
        Next: End With
    End If
    
    Set dcCopy = rt
End Function

Public Function dcWith(ky As Variant, it As Variant, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    If dc Is Nothing Then
        Set dcWith = dcWith(ky, it, _
        New Scripting.Dictionary)
    Else
        With dc
            If .Exists(ky) Then .Remove ky
            .Add ky, it
        End With
        Set dcWith = dc
    End If
End Function

Public Function nuDcPopulator( _
    Optional Dict As Scripting.Dictionary = Nothing, _
    Optional Opts As Long = 0 _
) As dcPopulator
    With New dcPopulator
        Set nuDcPopulator = .Using(Dict, Opts)
    End With
'Debug.Print dumpLsKeyVal(nuDcPopulator().Setting("A", "B").Setting("C", "D").Dictionary)
End Function

Public Function dcItemIfPresent( _
    dc As Scripting.Dictionary, ky As Variant, _
    Optional vtMissing As VbVarType _
) As Variant
    If dc Is Nothing Then
        dcItemIfPresent = noVal(vtMissing)
    Else
        With dc
            If .Exists(ky) Then
                If IsObject(.Item(ky)) Then
                    Set dcItemIfPresent = .Item(ky)
                Else
                    dcItemIfPresent = .Item(ky)
                End If
            Else
                If vtMissing = vbObject Then
                    Set dcItemIfPresent = noVal(vtMissing)
                Else
                    dcItemIfPresent = noVal(vtMissing)
                End If
            End If
        End With
    End If
End Function

Public Function dcInDc(dcKey As String, _
    inDc As Scripting.Dictionary _
) As Scripting.Dictionary
    Set dcInDc = dcOb(obOf( _
        dcItemIfPresent( _
        inDc, dcKey, vbObject _
    )))
End Function

Public Function dcInDcMk( _
    ky As Variant, dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcInDcMk --
    '''
    Dim rt As Scripting.Dictionary
    
    With dc
    If .Exists(ky) Then
        Set rt = dcOb(.Item(ky))
    Else
        Set rt = New Scripting.Dictionary
        .Add ky, rt
    End If
    End With
    
    Set dcInDcMk = rt
End Function

Public Function dcOfKeys2match( _
    ls As Variant _
) As Scripting.Dictionary
    '''
    ''' dcOfKeys2match -- generate a Dictionary
    '''     mapping a supplied Key, or Array
    '''     of Keys to itself or themselves
    '''     '
    '''     primary purpose is to provide
    '''     a 'reference' Dictionary of Keys
    '''     to be sought in other Dictionaries
    '''     '
    '''     (formerly d4g4f2)
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    If IsArray(ls) Then
        For Each ky In ls
        rt.Add ky, ky
        Next
    Else
        rt.Add ls, ls
    End If
    
    Set dcOfKeys2match = rt
End Function

Public Function dcKeysInCommon( _
    d0 As Scripting.Dictionary, _
    d1 As Scripting.Dictionary, _
    Optional pk As Long = 0 _
) As Scripting.Dictionary
    '''
    ''' dcKeysInCommon -- return intersection
    '''     of two Dictionary Objects based on
    '''     matching keys. Use optional pk value
    '''     to select which Dictionary's Items
    '''     to return in result:
    '''
    '''     0 returns an array of Items from both
    '''       this is the default
    '''     1 returns only Items from the first
    '''     2 returns only Items from the second
    '''
    ''' NOTE that if either Dictionary Object
    '''     is not supplied (is Nothing), then
    '''     an empty Dictionary is returned,
    '''     just as if an empty Dictionary
    '''     had been supplied.
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ls As Variant
    
    If d0 Is Nothing Then
        Set rt = dcKeysInCommon( _
            New Scripting.Dictionary, d1 _
        )
    ElseIf d1 Is Nothing Then
        Set rt = dcKeysInCommon( _
            d0, New Scripting.Dictionary _
        )
    Else
        Set rt = New Scripting.Dictionary
        With d0: For Each ky In .Keys
        If d1.Exists(ky) Then
            ls = Array( _
                .Item(ky), _
                d1.Item(ky) _
            )
            rt.Add ky, Array(ls, ls(0), ls(1))(pk)
        End If: Next: End With
    End If
    Set dcKeysInCommon = rt
End Function

Public Function dcKeysMissing( _
    dcWith As Scripting.Dictionary, _
    dcWout As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcKeysMissing -- return difference
    ''' of first Dictionary Object minus
    ''' those keys found in the second.
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dcWith
        For Each ky In .Keys
            If dcWout.Exists(ky) Then
                'don't include this one
            Else
                rt.Add ky, .Item(ky)
            End If
        Next
    End With
    Set dcKeysMissing = rt
End Function

Public Function dcKeysCombined( _
    d0 As Scripting.Dictionary, _
    d1 As Scripting.Dictionary, _
    Optional pk As Long = 0 _
) As Scripting.Dictionary
    '''
    ''' dcKeysCombined -- return union
    ''' of two Dictionary Objects. For
    ''' keys in both, use optional pk
    ''' value to select which Dictionary's
    ''' Items to return in result:
    '''
    '''     0 returns an array of Items from both
    '''       this is the default
    '''     1 returns only Items from the first
    '''     2 returns only Items from the second
    '''
    Dim rt As Scripting.Dictionary
    Dim ob As Scripting.Dictionary
    Dim ky As Variant
    Dim ls As Variant
    
    If pk > 1 Then
        Set rt = dcKeysCombined(d1, d0, 1)
    Else
        Set rt = dcKeysInCommon(d0, d1, pk)
        
        For Each ls In Array(d0, d1)
        Set ob = ls
        With dcKeysMissing(ob, rt) 'd0
        For Each ky In .Keys
            rt.Add ky, .Item(ky)
        Next: End With: Next
        
        'With dcKeysMissing(d1, rt)
        'For Each ky In .Keys
        '    rt.Add ky, .Item(ky)
        'Next: End With
    End If
    
    'Set rt = dcKeysInCommon()
    
    If d0 Is Nothing Then
        Stop
    ElseIf d1 Is Nothing Then
        Stop
    Else
        ''' NOTE[2023.04.10.1449]
        ''' need to review what's going on here
        ''' this LOOKS like an effor to include items
        ''' left out of earlier operation, however,
        ''' it's not clear this stage isn't redundant.
        With d0: For Each ky In .Keys
        If d1.Exists(ky) Then
            ls = Array( _
                .Item(ky), d1.Item(ky) _
            )
            If rt.Exists(ky) Then
                If ConvertToJson(Array(ls, ls(0), ls(1))(pk)) <> ConvertToJson(rt.Item(ky)) Then
                Stop
                End If
            Else
                rt.Add ky, Array(ls, ls(0), ls(1))(pk)
            End If
        End If
        Next: End With
    End If
    Set dcKeysCombined = rt
End Function

Public Function dcOfIdent( _
    src As Variant _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    If IsArray(src) Then
        For Each ky In src
            rt.Add ky, ky
        Next
    Else
        rt.Add src, src
    End If
    
    Set dcOfIdent = rt
End Function

Public Function dcTransposed( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' Transpose Key values of supplied
    ''' Dictionary with matching Item values.
    '''
    ''' As written, will ONLY work against
    ''' a Dictionary whose Item values,
    ''' like its Keys, are unique.
    '''
    Dim rt As Scripting.Dictionary
    Dim fm As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each fm In .Keys
            If rt.Exists(.Item(fm)) Then
                Stop
            Else
                rt.Add .Item(fm), fm
            End If
        Next
    End With
    Set dcTransposed = rt
'Debug.Print dumpLsKeyVal(dcTransposed(dcOfRgAddresses(dcByFormulaOnly(dcOfRgByFormulaR1C1(wsNamed(chosenWorkbook(), "StandardItems"))))), "|")
End Function

Public Function dcTransGrouped( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcTransGrouped
    '''     (derived from dcTransposed)
    '''
    ''' generate new Dictionary "tramsposing"
    ''' Key values with matching Item values
    ''' in supplied Dictionary.
    '''
    ''' because more than one Key might map to
    ''' the same Item, the returned Dictionary
    ''' maps each (Item) Key to a Dictionary of
    ''' the Keys which originally mapped to it.
    '''
    ''' each of these Dictionaries in turn
    ''' maps the original Key back to its
    ''' corresponding Item once more,
    ''' since, HEY, it might as well!
    '''
    ''' obviously, an effor to work around
    ''' the main limitation of the original
    ''' dcTransposed, which can only work
    ''' against a Dictionary whose Items,
    ''' like its Keys, are unique.
    '''
    Dim rt As Scripting.Dictionary
    Dim it As Scripting.Dictionary
    Dim ky As Variant
    Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        ar = Array(.Item(ky))
        If Not rt.Exists(ar(0)) Then
            rt.Add ar(0), New Scripting.Dictionary
        End If
        dcOb(rt.Item(ar(0))).Add ky, ar(0)
    Next: End With
    Set dcTransGrouped = rt
End Function

Public Function dcDepth( _
    dc As Scripting.Dictionary _
) As Long
    '''
    ''' this function extracts the "depth"
    ''' of the supplied Dictionary object,
    ''' that is, how many "levels" of
    ''' Dictionary objects it contains,
    ''' counting the supplied Dictionary
    ''' itself. When an actual Dictionary
    ''' is supplied, the value returned
    ''' will be at least 1. It will only
    ''' be zero when Nothing is supplied.
    '''
    Dim rt As Long
    Dim ck As Long
    Dim ky As Variant
    
    If dc Is Nothing Then
        dcDepth = 0
    Else
        rt = 0
        
        With dc: For Each ky In .Keys
            ck = dcDepth(dcOb(obOf(.Item(ky))))
            If ck > rt Then rt = ck
        Next: End With
        
        dcDepth = 1 + rt
    End If
End Function

Public Function dcFlattenDown( _
    Dict As Scripting.Dictionary, _
    Optional DownTo As Long = 1 _
) As Scripting.Dictionary
    '''
    ''' this function partially "flattens"
    ''' a hierarchy of Dictionary objects
    ''' (a Dictionary of Dictionaries,
    '''  potentially of more Dictionaries)
    ''' starting from the top, working
    ''' down to 'DownTo' levels
    '''
    Dim rt As Scripting.Dictionary
    Dim sd As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    Dim sk As Variant
    
    If Dict Is Nothing Then
        Set dcFlattenDown = Nothing
    Else
        Set rt = New Scripting.Dictionary
        
        With Dict
        For Each ky In .Keys
            it = Array(.Item(ky))
            
            If DownTo > 0 Then
                Set sd = dcFlattenDown(dcOb(obOf(it(0))), DownTo - 1)
            Else
                Set sd = Nothing
            End If
            
            If sd Is Nothing Then
                rt.Add ky, it(0)
            Else
                With sd
                For Each sk In .Keys
                    rt.Add ky & "." & sk, .Item(sk)
                Next
                End With
            End If
        Next
        End With
        
        Set dcFlattenDown = rt
    End If
End Function

Public Function dcFlattenUp( _
    Dict As Scripting.Dictionary, _
    Optional DownFrom As Long = 0 _
) As Scripting.Dictionary
    '''
    ''' this function partially "flattens"
    ''' a hierarchy of Dictionary objects
    ''' (a Dictionary of Dictionaries,
    '''  potentially of more Dictionaries)
    ''' starting BELOW the top, skipping
    ''' DownFrom levels before "flattening"
    ''' Dictionaries below and at that level
    '''
    Dim rt As Scripting.Dictionary
    Dim sd As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    Dim sk As Variant
    
    If Dict Is Nothing Then
        Set dcFlattenUp = Nothing
    Else
        Set rt = New Scripting.Dictionary
        
        With Dict
        For Each ky In .Keys
            it = Array(.Item(ky))
            Set sd = dcOb(obOf(it(0))) 'dcFlattenUp(, DownFrom - 1)
            
            If sd Is Nothing Then
                rt.Add ky, it(0)
            Else
                If DownFrom > 0 Then
                    rt.Add ky, dcFlattenUp(sd, DownFrom - 1)
                Else
                    With dcFlattenUp(sd, 0)
                    For Each sk In .Keys
                        rt.Add ky & "." & sk, .Item(sk)
                    Next
                    End With
                End If
            End If
        Next
        End With
        
        Set dcFlattenUp = rt
    End If
End Function

Public Function dcOfDcRekeyedSecToPri( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcOfDcRekeyedSecToPri - take Dictionary of Dictionaries
    '''     and return Dictionary of Dictionaries
    '''     with Secondary Keys promoted to Primary,
    '''     and Primary demoted to Secondary
    '''
    Dim rt As Scripting.Dictionary
    Dim sb As Scripting.Dictionary
    Dim kp As Variant
    Dim ks As Variant
    Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each kp In .Keys
        Set sb = dcOb(.Item(kp))
        If sb Is Nothing Then 'we got a problem
            'Stop
            Debug.Print ; 'Breakpoint Landing
        Else
        With sb: For Each ks In .Keys
            ar = Array(.Item(ks))
            With rt
                If Not .Exists(ks) Then
                .Add ks, New Scripting.Dictionary
                End If
                
                With dcOb(.Item(ks))
                If .Exists(kp) Then 'another problem
                    Stop
                Else
                    .Add kp, ar(0)
                End If: End With
            End With
        Next: End With: End If
    Next: End With
    
    Set dcOfDcRekeyedSecToPri = rt
End Function

Public Function dcCmpTextOf2items( _
    id0 As String, id1 As String, _
    it0 As Variant, it1 As Variant _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim dc0 As Scripting.Dictionary
    Dim dc1 As Scripting.Dictionary
    Dim ob0 As Object
    Dim ob1 As Object
    Dim tx0 As String
    Dim tx1 As String
    Dim ck As Long
    Dim c2 As Long
    
    ck = IIf(IsObject(it0), 1, 0) _
       + IIf(IsObject(it1), 2, 0) _
       + IIf(IsEmpty(it0), 4, 0) _
       + IIf(IsEmpty(it1), 8, 0)
    If ck And 1 Then ck = ck + IIf(it0 Is Nothing, 4, 0)
    If ck And 2 Then ck = ck + IIf(it1 Is Nothing, 8, 0)
    ''' UPDATE[2021.12.09]:
    '''     Added trap for Nothing Objects
    '''     to treat them as Empty as well
    
    c2 = ck And 12 '$11XX
    If c2 Then 'at least one side
               'is Empty/Nothing
        With nuDcPopulator()
            If c2 = 12 Then 'both
                'are equally Empty
                Set rt = .Setting( _
                    "==", "" _
                ).Dictionary()
            ElseIf c2 = 8 Then 'it's just it1
                'and id0 needs processed
                'If ck And 1 Then
                '    Set rt = dcCmpTextOf2dc( _
                '        dcOb(it0), .Dictionary() _
                '    )
                'Else
                    Set rt = .Setting( _
                        id0, it0 _
                    ).Dictionary()
                'End If
            Else 'it0 is the empty one
                'and id1 needs processed
                'If ck And 2 Then
                '    Set rt = dcCmpTextOf2dc( _
                '        .Dictionary(), dcOb(it1) _
                '    )
                'Else
                    Set rt = .Setting( _
                        id1, it1 _
                    ).Dictionary()
                'End If
            End If
        End With
    Else 'both sides have data
        If ck = 3 Then 'PROBABLY a couple
            'of comparable Dictionaries.
            'compare these recursively
            Set rt = dcCmpTextOf2dc( _
                dcOb(it0), dcOb(it1) _
            )
        ElseIf ck = 0 Then 'NO Dictionaries
            'just a couple of String
            'or (hopefully) String-
            'compatible values.
            'compare them directly.
            tx0 = CStr(it0)
            tx1 = CStr(it1)
            
            'Set rt = New Scripting.Dictionary
            'With rt
            With nuDcPopulator()
                If tx0 = tx1 Then 'match found
                    Set rt = .Setting( _
                        "==", tx0 _
                    ).Dictionary()
                    '.Add "==", tx0
                Else 'mismatched
                    Set rt = .Setting( _
                        id0, tx0 _
                    ).Setting( _
                        id1, tx1 _
                    ).Dictionary()
                    '.Add id0, tx0
                    '.Add id1, tx1
                End If
            End With
        Else 'we've got a Dictionary AND a String!
            'can't compare them in any way
            'just add each on its own side
            With nuDcPopulator()
                Set rt = .Setting( _
                    id0, it0 _
                ).Setting( _
                    id1, it1 _
                ).Dictionary()
            End With
        'ElseIf (ck And 14) = 6 Then '%011X
        '    'covers cases 6 and 7
        '    'excludes 4 and 5
        '    'along with 8~F
        '    'Stop
        '    Set rt = dcCmpTextOf2dc( _
        '        New Scripting.Dictionary, dcOb(it1) _
        '    )
        'ElseIf (ck And 13) = 9 Then '%10X1
        '    'covers cases 9 and B
        '    'excludes 8 and A
        '    'along with 0~7
        '    Set rt = dcCmpTextOf2dc( _
        '        dcOb(it0), New Scripting.Dictionary _
        '    )
        '    'ElseIf ck And 4 Then 'it0 is missing
        '    '    If ck And 2 Then 'it1 is a Dictionary
        '    '    Else
        '    '        .Add id1, CStr(it1)
        '    '    End If
        '    'ElseIf ck = 8 Then
        '    '    .Add id0, CStr(it0)
        '    Else 'either one or both members are Empty,
        '        ' Nothing, or of incompatible form.
        '        ' They cannot be compared directly,
        '        ' but must, if present, be separately
        '        ' included as they are.
        '
        '        ''' NOTE
        '        'Debug.Print "First Item ";
        '        If (ck And 4) = 0 Then 'it0 is present
        '            'Debug.Print "is " & TypeName(it0)
        '            If ck And 1 Then
        '                .Add id0, it0
        '            Else
        '                .Add id0, CStr(it0)
        '            End If
        '            Debug.Print ; 'Breakpoint Landing
        '        Else
        '            'Debug.Print "NOT present!"
        '            Debug.Print ; 'Breakpoint Landing
        '        End If
        '        'Stop
        '
        '        'Debug.Print "Second Item ";
        '        If (ck And 8) = 0 Then 'it1 is present
        '            'Debug.Print "is " & TypeName(it1)
        '            If ck And 2 Then
        '                .Add id1, it1
        '            Else
        '                .Add id1, CStr(it1)
        '            End If
        '            Debug.Print ; 'Breakpoint Landing
        '        Else
        '            'Debug.Print "NOT present!"
        '            Debug.Print ; 'Breakpoint Landing
        '        End If
        '        'Stop
        '    End If
        End If
    End If
    
    Set dcCmpTextOf2items = rt
End Function

Public Function dcCmpTextOf2dc( _
    dc0 As Scripting.Dictionary, _
    dc1 As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim qi As Scripting.Dictionary
    Dim nm0 As String
    Dim nm1 As String
    Dim tx0 As String
    Dim tx1 As String
    Dim ky As Variant
    Dim nm As String
    
    nm0 = "src0"
    nm1 = "src1"
    
    If dc0 Is Nothing Then
        Set rt = dcCmpTextOf2dc( _
            New Scripting.Dictionary, dc1 _
        )
    ElseIf dc1 Is Nothing Then
        Set rt = dcCmpTextOf2dc(dc0, _
            New Scripting.Dictionary _
        )
    Else
        Set rt = New Scripting.Dictionary
        
        With dc0 'add all from wb0
            'and matching from wb1
            For Each ky In .Keys
                'tx0 = CStr(.Item(ky))
                
                'Set qi = New Scripting.Dictionary
                'rt.Add ky, qi
                
                If dc1.Exists(ky) Then 'check for match
                    rt.Add ky, dcCmpTextOf2items( _
                        nm0, nm1, .Item(ky), dc1.Item(ky) _
                    )
                    'tx1 = CStr(dc1.Item(ky))
                    
                    'If tx0 = tx1 Then 'match found
                    '    qi.Add "==", tx0
                    'Else 'mismatched
                    '    qi.Add nm0, tx0
                    '    qi.Add nm1, tx1
                    'End If
                Else 'no match
                    'qi.Add nm0, tx0
                    rt.Add ky, dcCmpTextOf2items( _
                        nm0, nm1, .Item(ky), Empty _
                    )
                End If
            Next
        End With
        
        With dc1 'add any missed from wb1
            For Each ky In .Keys
                If rt.Exists(ky) Then 'skip it
                    'picked up first round
                Else 'missed it before
                    'so add it now
                    
                    'tx1 = CStr(.Item(ky))
                    '
                    'Set qi = New Scripting.Dictionary
                    'qi.Add nm1, tx1
                    'rt.Add ky, qi
                    rt.Add ky, dcCmpTextOf2items( _
                        nm0, nm1, Empty, .Item(ky) _
                    )
                End If
            Next
        End With
    End If
    
    Set dcCmpTextOf2dc = rt
End Function

Public Function dcCmpTextOf2subDc( _
    dc As Scripting.Dictionary, _
    k0 As String, k1 As String _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    Set rt = dcCmpTextOf2dc( _
        dcInDc(k0, dc), _
        dcInDc(k1, dc) _
    )
    Debug.Print ;
    Set dcCmpTextOf2subDc = rt
End Function

Public Function dcWBQbyCmpResult( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ck As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim ky As Variant
    Dim gk As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dc
        For Each ky In .Keys
            Set ck = .Item(ky)
            
            If ck.Count > 1 Then
                gk = "!="
            ElseIf ck.Count < 1 Then
                Stop 'because SOMETHING went wrong
                gk = "**"
            Else
                gk = ck.Keys(0)
            End If
            
            With rt
                If .Exists(gk) Then
                    Set gp = .Item(gk)
                Else
                    Set gp = New Scripting.Dictionary
                    rt.Add gk, gp
                End If
                
                gp.Add ky, ck
            End With
        Next
    End With
    
    Set dcWBQbyCmpResult = rt
End Function

