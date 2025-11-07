

Public Function xprtModText() As Long
    Dim vp As VBIDE.VBProject
    Dim vc As VBIDE.VBComponent
    Dim dc As Scripting.Dictionary
    Dim d1 As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim n1 As String
    Dim n2 As String
    
    Set dc = dcOfVbProjects(ThisApplication.VBAProjects) 'ThisWorkbook.Application.VBE.VBProjects
    
    n1 = "C:\Users\athompson\Documents\dvl\libExt.xlsm"
    Set d1 = dc(n1)
    send2clipBdWin10 n1 & vbNewLine & dumpKeyedText(d1, d1)
    Stop
    
    n2 = "C:\Users\athompson\Documents\dvl\libExt-rcvr-2017-0608.xlsm"
    Set d2 = dc(n2)
    send2clipBdWin10 n2 & vbNewLine & dumpKeyedText(d1, d2)
    Stop
End Function
'C:\Users\athompson\Documents\dvl\libExt-rcvr-2017-0608.xlsm
'C:\Users\athompson\Documents\dvl\libExt.xlsm

Public Function txOfVbModule(cm As VBIDE.CodeModule) As String
    If cm Is Nothing Then
        txOfVbModule = ""
    Else
    With cm
        If .CountOfLines > 0 Then
            txOfVbModule = .Lines(1, .CountOfLines)
        Else
            txOfVbModule = ""
        End If
    End With
    End If
End Function

Public Function lVCXg1f1( _
    pj As VBIDE.VBProject _
) As Scripting.Dictionary
    '''
    ''' lVCXg1f1 - generate Dictionary of
    '''     Dictionaries of collected text
    '''     of all procedures in each module
    '''     of given VBProject, keyed first
    '''     by module, and then by procedure
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dcOfVbModules(pj)
    For Each ky In .Keys
        rt.Add ky, dcOfVbProcs( _
        obVbCodeMod(.Item(ky)))
    Next: End With
    
    Set lVCXg1f1 = rt
End Function

Public Function lVCXg1f2( _
    pj As VBIDE.VBProject _
) As Scripting.Dictionary
    '''
    ''' lVCXg1f2 - generate Dictionary of
    '''     procedures in given VBProject
    '''     keyed first by procedure name
    '''     and then by module name
    '''
    '''     this is accomplished using
    '''     function dcOfDcRekeyedSecToPri
    '''     to promote function names over
    '''     module names, with the expected
    '''     result being a Dictionary of
    '''     mostly single-entry Dictionaries.
    '''
    '''     each of these can then be replaced
    '''     with the text of its one entry
    '''     in a subsequent function
    '''
    '''     multi-entry Dictionaries might
    '''     have to be left as is
    '''
    '''     note that with all headers filed
    '''     under a blank key, at least one
    '''     multi-entry Dictionary is guaranteed
    '''
    Set lVCXg1f2 = dcOfDcRekeyedSecToPri(lVCXg1f1(pj))
End Function

Public Function lVCXg1f3( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' lVCXg1f3 - return transformation
    '''     of supplied Dictionary of sort
    '''     returned by lVCXg1f2 as described
    '''     in that procedure's comments
    '''
    '''     each single-entry Dictionary is
    '''     replaced with the text of its entry
    '''     while multi-entry Dictionaries
    '''     are simply copied over
    '''
    '''     note that this function accepts
    '''     a Dictionary and NOT a VBProject
    '''     as lVCXg1f2 does. this permits other
    '''     functions to be applied to the result
    '''     of a single call to lVCXg1f2
    '''     and so reduce redundancy
    '''
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dc 'lVCXg1f3(pj)
    For Each ky In .Keys
        'If Len(ky) Then 'to filter out header
        
        Set wk = dcOb(.Item(ky))
        If wk Is Nothing Then
            If VarType(.Item(ky)) = vbString Then
                'just add to Dictionary as itself
                rt.Add ky, .Item(ky)
                
                'this is hoped to address
                'possible "rerun" against
                'result of a prior call
                'to this function
            Else
                Stop 'problem!
            End If
        Else
        With wk
            If .Count > 1 Then
                rt.Add ky, wk
            ElseIf .Count < 1 Then
                Stop 'another problem!
            Else
                rt.Add ky, .Items(0)
            End If
        End With: End If
        
        'Else 'to process header separately
        'End If
    Next: End With
    
    Set lVCXg1f3 = rt
End Function

Public Function lVCXg2f1(tx As String) As String
    '''
    ''' lVCXg2f1 - "decontinuate" VB text
    '''     replace " _" and newline
    '''     at end of each continued
    '''     line with a vertical tab
    '''
    '''     thus reducing each continued
    '''     line sequence to a single line
    '''     while retaining a clear marker
    '''     for rebreaking, if necessary
    '''
    'lVCXg2f1 = Join(Split(tx, _
    " _" & vbNewLine), _
    " _" & vbVerticalTab)
    lVCXg2f1 = Replace(tx, _
        " _" & vbNewLine, _
        " _" & vbVerticalTab _
    )
End Function

Public Function lVCXg2f2(tx As String) As String
    '''
    ''' lVCXg2f2 - "decomment" and "dequote" VB text
    '''     locate and remove any and all
    '''     remarks and string constants
    '''     from a line of VB text
    '''     '
    '''     a tricky operation, requiring detection
    '''     of the FIRST of either single or double
    '''     quotes (' or ") on a line.
    '''     '
    '''     subsequent procedure depends on WHICH
    '''     is discovered first. for a single quote,
    '''     the entire remainder of the line should
    '''     be dropped.
    '''     '
    '''     for a double quote, only the text prior
    '''     to the NEXT double quote (which MUST be
    '''     present) is dropped. the remainder must
    '''     then be searched for further reductions
    '''     '
    '''     REV[2023.04.20.1001]: recalling that TWO
    '''     double quotes INSIDE a string constant
    '''     form an "escape" sequence representing
    '''     ONE double quote, it is necessary to
    '''     replace all such instances with another
    '''     placeholder, in order to ensure the
    '''     correct closing quote is found. active
    '''     implementation has been modified
    '''     to achieve this
    '''
    lVCXg2f2 = lVCXg2f2b(tx)
'Debug.Print lVCXg2f2("dl = ""NO_RAW_STOCK"" & vbTab & ""<No Raw Stock Declared>""")
'Debug.Print lVCXg2f2("rt.Add ky & vbTab & dl, sd 'rt.Add ky & vbTab & ""(RAW STOCK NOT YET IMPLMENTED)"", sd")
'Debug.Print lVCXg2f2("Debug.Print ""  gnMatlNum = """""" & gnMatlNum & """"""""")
End Function

Public Function lVCXg2f2a(tx As String) As String
    '''
    ''' lVCXg2f2a - "decomment" and "dequote" VB text
    '''     initial implementation deactivated
    '''     to be held in reserve pending
    '''     verification of rewritten
    '''     version lVCXg2f2b
    '''
    Dim rt As String
    Dim ar() As String
    Dim qt1 As Long 'location of first single quote (rem)
    Dim qt2 As Long 'location of first double quote (string)
    Dim rmk As Long 'location of first 'Rem'
    
    rmk = InStr(1, tx, "rem", vbTextCompare)
    If rmk Then
        Debug.Print tx
        Stop 'to review
    End If
    
    qt1 = InStr(1, tx, "'")
    qt2 = InStr(1, tx, """")
    
    Debug.Print tx 'while debugging only
    
    If qt1 * qt2 Then
        Stop
        If qt1 < qt2 Then
            rt = Left$(tx, qt1)
        Else
            Stop
        End If
    ElseIf qt1 Then
        rt = Left$(tx, qt1)
        Debug.Print rt
        Stop
    ElseIf qt2 Then
        ar = Split(tx, """", 3)
        rt = lVCXg2f2a(ar(0) & "$$" & ar(2))
        Debug.Print rt
        Stop
    Else
        rt = tx
    End If
    
    lVCXg2f2a = rt
End Function

Public Function lVCXg2f2b(tx As String) As String
    '''
    ''' lVCXg2f2b - "decomment" and "dequote" VB text
    '''     currently active implementation @[2023.04.20.0957]
    '''
    Dim rt As String
    Dim rf As String
    Dim ar() As String
    Dim qt(2) As Long 'locations of first single and double quote
    Dim mx As Long
    Dim dx As Long
    Dim rmk As Long 'location of first 'Rem'
    
    'Debug.Print "IN: "; tx 'while debugging only
    
    rf = "'"""
    mx = Len(tx)
    For dx = 1 To 2
        qt(dx) = InStr(1, tx, Mid(rf, dx, 1))
        If qt(dx) = 0 Then qt(dx) = 1 + mx
    Next
    
    rmk = InStr(1, " " & tx, " rem ", vbTextCompare)
    If rmk = 0 Then rmk = Len(tx) + 2
    If rmk < qt(1) Then
    If rmk < qt(2) Then
    'If Mid$(tx, rmk + 3, 1) <= " " Then
        Debug.Print tx
        Debug.Print Left$(tx, rmk - 1)
        Debug.Print Mid$(tx, rmk)
        Stop 'to review
    'End If
    End If
    End If
    
    
    If qt(1) < qt(2) Then
        rt = Left$(tx, qt(1))
        'Debug.Print rt
        'Stop
    ElseIf qt(2) > mx Then
        rt = tx
        'Debug.Print rt
        'Stop
    Else
        ar = Split(tx, """", 2) 'was 3
        
        'rt = lVCXg2f2b(ar(0) & "$$" & ar(2))
        'rt = ar(0) & """""" & lVCXg2f2b(ar(2))
        rt = ar(0) & """"""
        
        'ar = Split(Join( _
            Split(ar(1), """""" _
            ), vbFormFeed _
        ), """", 2)
        ''' REF: Replace("expr","find","rplc")
        ar = Split(Replace( _
            ar(1), """""", vbFormFeed _
        ), """", 2)
        
        If UBound(ar) > 0 Then
            rt = rt & lVCXg2f2b(Replace( _
                ar(1), vbFormFeed, """""" _
            ))
        Else
            Stop 'problem!
        End If
        'Debug.Print rt
        'Stop
    End If
    
    lVCXg2f2b = rt
End Function

Public Function lVCXg2f3(tx As String) As Scripting.Dictionary
    '''
    ''' lVCXg2f3 - decompose VB text to a "keyword" Dictionary
    '''     mapping each "keyword" to a count of instances
    '''     note that "keyword" includes not only words
    '''     reserved by VB, but any unbroken set of non-space
    '''     characters: variables, procedure names, etc.
    '''     '
    '''     the Keys returned in the resulting Dictionary
    '''     can then be matched against another Dictionary
    '''     listing all entities defined in a VB project,
    '''     as returned by other functions in this module,
    '''     thereby producting a first-level dependency map
    '''     '
    '''     note that the text supplied IS assumed to be
    '''     Visual Basic code, which should already be
    '''     "cleaned" of any "inactive" elements: comments
    '''     and the content of string literals. this should
    '''     limit the rate of "false positives," that is,
    '''     detection of entity names not actually required
    '''     by a procedure, but mentioned either in string
    '''     literals or comments the compiler does not parse.
    '''     '
    Dim rt As Scripting.Dictionary
    Dim wk As String
    Dim ls As String
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    wk = tx
    ls = Join(Array(vbCrLf, vbTab, "():&.!,[]"), "")
    Do
        wk = Replace(wk, Left$(ls, 1), " ")
        ls = Mid$(ls, 2)
    Loop While Len(ls)
    
    With rt: For Each ky In Split(wk, " ")
    If Len(ky) Then
    If .Exists(ky) Then
        .Item(ky) = 1 + .Item(ky)
    Else
        .Add ky, 1
    End If: End If: Next: End With
    
    Set lVCXg2f3 = rt
'send2clipBdWin10 ConvertToJson(lVCXg2f3(lVCXg3f2(lVCXg1f3(lVCXg1f2(ThisDocument.VBAProject.VBProject))).Item("dcAiDocsByCompList")), vbTab)
End Function

Public Function lVCXg3f1(tx As String) As String
    '''
    ''' lVCXg3f1 - "clean" supplied VB text
    '''     removing all comments and content
    '''     of string constants, and leaving
    '''     only comment markers and empty
    '''     strings in their place
    '''
    '''     goal of this function is to remove
    '''     any "inactive" content from text
    '''     of VB procedure definitions, and
    '''     thereby limit the number of "false
    '''     positives" returned in a search
    '''     for procedural dependencies
    '''
    Dim wk() As String
    Dim mx As Long
    Dim dx As Long
    
    wk = Split(Replace(tx, _
        " _" & vbNewLine, _
        " _" & vbVerticalTab _
    ), vbNewLine)
    
    mx = UBound(wk)
    For dx = LBound(wk) To mx
        wk(dx) = lVCXg2f2(wk(dx))
    Next
    
    lVCXg3f1 = Replace( _
        Join(wk, vbNewLine), _
        " _" & vbVerticalTab, _
        " _" & vbNewLine _
    )
'send2clipBdWin10 lVCXg3f1(lVCXg1f3(lVCXg1f2(ThisDocument.VBAProject.VBProject)).Item("dcAiDocsByCompList"))
End Function

Public Function lVCXg3f2( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' lVCXg3f2 - "clean" all String Items
    '''     in supplied Dictionary, including
    '''     any sub Dictionaries, as VB text
    '''
    '''     any Items not recognized as String
    '''     or Dictionary Items are passed
    '''     through as is, at present
    '''
    '''     might want to reconsider this,
    '''     and probably make it optional
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ar() As Variant
    'tx As String
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each ky In .Keys
        ar = Array(.Item(ky))
        If VarType(ar(0)) = vbString Then
            rt.Add ky, lVCXg3f1(CStr(ar(0)))
        ElseIf TypeOf ar(0) Is Scripting.Dictionary Then
            rt.Add ky, lVCXg3f2(dcOb(ar(0)))
        Else
            rt.Add ky, ar(0)
        End If
    Next: End With
    
    Set lVCXg3f2 = rt
'send2clipBdWin10 lVCXg3f2(lVCXg1f3(lVCXg1f2(ThisDocument.VBAProject.VBProject))).Item("dcAiDocsByCompList")
End Function

Public Function lVCXg3f3( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' lVCXg3f3 - "clean" all String Items
    '''     in supplied Dictionary, including
    '''     any sub Dictionaries, as VB text
    '''
    '''     any Items not recognized as String
    '''     or Dictionary Items are passed
    '''     through as is, at present
    '''
    '''     might want to reconsider this,
    '''     and probably make it optional
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ar() As Variant
    'tx As String
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each ky In .Keys
        ar = Array(.Item(ky))
        If VarType(ar(0)) = vbString Then
            rt.Add ky, lVCXg3f1(CStr(ar(0)))
        ElseIf TypeOf ar(0) Is Scripting.Dictionary Then
            rt.Add ky, lVCXg3f3(dcOb(ar(0)))
        Else
            rt.Add ky, ar(0)
        End If
    Next: End With
    
    Set lVCXg3f3 = rt
'send2clipBdWin10 lVCXg3f3(lVCXg1f3(lVCXg1f2(ThisDocument.VBAProject.VBProject))).Item("dcAiDocsByCompList")
End Function

Public Function lVCXg3f4( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' lVCXg3f4 - generate "keyword" lists for all
    '''     String Items in supplied Dictionary,
    '''     including any sub Dictionaries
    '''     '
    '''     derived from lVCXg3f3, this function just
    '''     adds one level of processing, calling
    '''     lVCXg2f3 against the results of lVCXg3f1
    '''     '
    '''     it does NOT require input from lVCXg3f3
    '''     and would likely fail against such a source
    '''     '
    '''     any Items not recognized as String
    '''     or Dictionary Items are passed
    '''     through as is, at present
    '''     '
    '''     might want to reconsider this,
    '''     and probably make it optional
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ar() As Variant
    'tx As String
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each ky In .Keys
        ar = Array(.Item(ky))
        If VarType(ar(0)) = vbString Then
            rt.Add ky, lVCXg2f3(lVCXg3f1(CStr(ar(0))))
        ElseIf TypeOf ar(0) Is Scripting.Dictionary Then
            rt.Add ky, lVCXg3f4(dcOb(ar(0)))
        Else
            rt.Add ky, ar(0)
        End If
    Next: End With
    
    Set lVCXg3f4 = rt
'send2clipBdWin10 ConvertToJson(lVCXg3f4(lVCXg1f3(lVCXg1f2(ThisDocument.VBAProject.VBProject))).Item("dcAiDocsByCompList"), vbTab)
'send2clipBdWin10 ConvertToJson(lVCXg3f4(lVCXg1f3(lVCXg1f2(ThisDocument.VBAProject.VBProject))), vbTab)
End Function

Public Function lVCXg4f1(dc As Scripting.Dictionary, _
    Optional rf As Scripting.Dictionary = Nothing, _
    Optional rt As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' lVCXg4f1 - generate basic dependency list from
    '''     supplied Dictionary of "keyword" Dictionaries
    '''     keyed either to VB procedure names, or for
    '''     a subset of identically named procedures,
    '''     the module names of each implementation
    '''     '
    '''     note that this function does NOT disambiguate
    '''     dependencies on multiply defined names.
    '''     that is a task left to the client supplying
    '''     the initial Dictionary, on the assumption
    '''     that said client will have retained any
    '''     prior source used to generate it
    '''     '
    '''     note also that optional Dictionary parameters
    '''     rf and rt are NOT expected to be provided by
    '''     an outside client, but passed to a recursive
    '''     invocation when processing a multiply defined
    '''     procedure name. for this purpose, the original
    '''     source Dictionary dc is passed through rf to
    '''     ensure its availability to all recursive calls
    '''     '
    Dim ky As Variant
    Dim ar() As Variant
    Dim wk As Scripting.Dictionary
    
    If rf Is Nothing Then
        Set rt = lVCXg4f1(dc, dc, rt)
    ElseIf rt Is Nothing Then
        Set rt = lVCXg4f1(dc, rf, _
        New Scripting.Dictionary)
    Else
    With dc: For Each ky In .Keys
        Set wk = .Item(ky)
        With wk
        If .Count > 0 Then
            ar = Array(.Item(.Keys(0)))
            
            'wk, NOT ar(0)
            If IsNumeric(ar(0)) Then 'this is a base list
                'intersect with reference Dictionary rf
                'taking only the usage counts from wf
                rt.Add ky, dcKeysInCommon(wk, rf, 1)
            ElseIf TypeOf ar(0) Is Scripting.Dictionary Then
                'we have a subset of implementations
                'keyed to module locations
                'so need to go down a level
                rt.Add ky, lVCXg4f1(wk, rf, _
                New Scripting.Dictionary)
            ElseIf VarType(ar(0)) = vbString Then
                Debug.Print ; 'Breakpoint Landing
                Stop 'because this does NOT normally happen
                'rt.Add ky, dcKeysInCommon(rf, wk, 0)
            Else
                Stop 'because this shouldn't happen either
                Debug.Print ; 'Breakpoint Landing
            End If
        End If: End With
        
        Debug.Print ; 'Breakpoint Landing
        'Stop
    Next: End With: End If
    
    Set lVCXg4f1 = rt
'send2clipBdWin10 ConvertToJson(lVCXg4f1(lVCXg3f4(lVCXg1f3(lVCXg1f2(ThisDocument.VBAProject.VBProject)))), vbTab)
End Function

Public Function dxOfVbProcLocsInMod( _
    cm As VBIDE.CodeModule _
) As Scripting.Dictionary
    '''
    ''' dxOfVbProcLocsInMod -- Return Dictionary
    '''     of Procedure Locations in CodeModule.
    '''     derived from dcOfVbProcs as a "light"
    '''     alternative that only returns an "index"
    '''     of procedures, leaving the client to
    '''     extract their text as needed
    '''
    ''' REV[2023.05.05.1307]: copied from dcOfVbProcs
    '''     see that function for prior REVs
    '''     if and where applicable
    '''
    Dim rt As Scripting.Dictionary
    Dim mx As Long
    Dim dx As Long
    Dim fw As Long
    Dim ck As String
    'Dim tx As String
    Dim ar As Variant
    Dim tp As Long
    
    Set rt = New Scripting.Dictionary
    ar = Array( _
        Array(vbext_pk_Proc, ""), _
        Array(vbext_pk_Get, ""), _
        Array(vbext_pk_Let, "=#"), _
        Array(vbext_pk_Set, "=@") _
    )
    
    With cm
        mx = .CountOfLines
        dx = .CountOfDeclarationLines
        ''' REV[2023.05.05.1329]: modified
        ''' dx assignment for reuse below
        ''' and call .CountOfDeclarationLines
        ''' only once.
        rt.Add "", Array(1, dx)  '.Lines
        ''' REV[2023.05.05.1310]: replaced
        ''' .Lines with Array to capture start
        ''' line and line count of header
        ''' (AKA DeclarationLines)
        
        dx = 1 + dx
        Do While dx < mx
            fw = dx
            Do While fw < mx
                ck = .ProcOfLine(fw, vbext_pk_Proc)
                If Len(ck) = 0 Then
                    fw = fw + 1
                Else
                    tp = 0
                    On Error Resume Next
                    Do
                        Err.Clear
                        dx = .ProcStartLine(ck, ar(tp)(0))
                        If Err.Number Then dx = fw + 1
                        If dx <> fw Then tp = tp + 1
                        If tp > 3 Then Stop 'should NOT happen...
                        'Debug.Print fw, dx
                    Loop Until Err.Number = 0 And dx = fw
                    Err.Clear
                    On Error GoTo 0
                    
                    fw = .ProcCountLines(ck, ar(tp)(0))
                    'tx = .Lines(dx, fw)
                    rt.Add ck & ar(tp)(1), Array(dx, fw) 'tx
                    ''' REV[2023.05.05.1337]: replaced tx
                    ''' .Lines with Array(dx, fw) to capture
                    ''' start line and line count of procedure
                    dx = dx + fw
                    fw = mx
                End If
            Loop
        Loop
    End With
    Set dxOfVbProcLocsInMod = rt
End Function

Public Function vbProcTextFromPrj(nm As String, _
    Optional pj As VBIDE.VBProject = Nothing _
) As String
    '''
    ''' vbProcTextFromPrj
    '''     derived from vbTextOfProcInProject (sort of)
    '''
    ''' NOTE: In order to use this Function
    '''         from an external library,
    '''         the option has been removed
    '''         to call itself recursively
    '''         against ThisWorkbook. Since
    '''         ThisWorkbook would be the
    '''         library itself, a call against
    '''         it could result in a breach
    '''         of security.
    '''
    Dim dc As Scripting.Dictionary
    Dim ky As String
    Dim ar As Variant
    Dim cm As VBIDE.CodeModule
    
    If pj Is Nothing Then
        ky = ""
    Else
        With dxOfVbProcLocsInPrj(pj)
        If .Exists(nm) Then
            ar = Array(Nothing)
            
            Set dc = .Item(nm)
            With dc
            If .Count > 0 Then
                ky = .Keys(0)
                
                If .Count > 1 Then
                ky = userChoiceFromDc(dc, ky)
                End If
                
                If .Exists(ky) Then
                ar = .Item(ky)
                End If
            End If: End With
            
            ky = ""
            If UBound(ar) >= 2 Then
                Set cm = obVbCodeMod(obOf(ar(0)))
                If Not cm Is Nothing Then
                    ky = cm.Lines( _
                    ar(1), ar(2))
                End If
            End If
        Else
            ky = ""
        End If: End With
    End If
    
    vbProcTextFromPrj = ky
End Function

Public Function dxOfVbProcLocsInPrj( _
    pj As VBIDE.VBProject _
) As Scripting.Dictionary
    '''
    ''' dxOfVbProcLocsInPrj - generate Dictionary of
    '''     Dictionaries of collected text
    '''     of all procedures in each module
    '''     of given VBProject, keyed first
    '''     by module, and then by procedure
    '''
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim cm As VBIDE.CodeModule
    Dim kMd As Variant
    Dim kPr As Variant
    Dim ar As Variant
    'Dim nm As String
    
    Set rt = New Scripting.Dictionary
    
    With dcOfVbModules(pj)
    For Each kMd In .Keys
        Set cm = obVbCodeMod(.Item(kMd))
        
        With dxOfVbProcLocsInMod(cm)
        For Each kPr In .Keys
            With rt
                If Not .Exists(kPr) Then
                .Add kPr, New Scripting.Dictionary
                End If
                
                Set wk = .Item(kPr)
            End With
            
            ar = .Item(kPr)
            With wk
            If .Exists(kMd) Then
                Stop 'for problem
            Else
                .Add kMd, Array(cm, ar(0), ar(1))
            End If: End With
        Next: End With
    Next: End With
    
    Set dxOfVbProcLocsInPrj = rt
End Function

Public Function dcOfVbProcs( _
    cm As VBIDE.CodeModule _
) As Scripting.Dictionary
    '''
    ''' dcOfVbProcs -- Return Dictionary of Procedures
    '''             -- from supplied CodeModule
    '''
    ''' REV[2023.04.19.1146]: added code to capture
    '''     declaration text preceding all proc defs
    ''' REV[2023.02.15.0904]: modified to accommodate Let,
    '''     Set, and Get Procedures. Let and Set Procedures
    '''     are stored under keys modified to indicate their
    '''     role: "=#" for Let indicates assignment to a value,
    '''     while "=@" for Set indicates an Object assignment.
    ''' NOTE: this new version, while now able to accommodate
    '''     Class Modules, is likely not the most efficient
    '''     in addressing the problem. Further development
    '''     might be warranted, should this prove an issue.
    '''
    Dim rt As Scripting.Dictionary
    Dim mx As Long
    Dim dx As Long
    Dim fw As Long
    Dim ck As String
    Dim tx As String
    Dim ar As Variant
    Dim tp As Long
    
    Set rt = New Scripting.Dictionary
    ar = Array( _
        Array(vbext_pk_Proc, ""), _
        Array(vbext_pk_Get, ""), _
        Array(vbext_pk_Let, "=#"), _
        Array(vbext_pk_Set, "=@") _
    )
    
    With cm
        mx = .CountOfLines
        dx = 1 + .CountOfDeclarationLines
        ''' REV[2023.04.19.1146]: added following
        ''' to capture header, AKA declaration lines
        rt.Add "", .Lines(1, .CountOfDeclarationLines)
        
        Do While dx < mx
            fw = dx
            Do While fw < mx
                ck = .ProcOfLine(fw, vbext_pk_Proc)
                If Len(ck) > 0 Then
                    tp = 0
                    On Error Resume Next
                    Do
                        Err.Clear
                        dx = .ProcStartLine(ck, ar(tp)(0))
                        If Err.Number Then dx = fw + 1
                        If dx <> fw Then tp = tp + 1
                        If tp > 3 Then Stop 'should NOT happen...
                        'Debug.Print fw, dx
                    Loop Until Err.Number = 0 And dx = fw
                    Err.Clear
                    On Error GoTo 0
                    
                    fw = .ProcCountLines(ck, ar(tp)(0))
                    tx = .Lines(dx, fw)
                    rt.Add ck & ar(tp)(1), tx
                    dx = dx + fw
                    fw = mx
                Else
                    fw = fw + 1
                End If
            Loop
        Loop
    End With
    Set dcOfVbProcs = rt
End Function

Public Function dcOfVbProcs_obs2023_0419(cm As VBIDE.CodeModule) As Scripting.Dictionary
    '''
    ''' dcOfVbProcs_obs2023_0419     -- Return Dictionary of Procedures
    '''                 -- from supplied CodeModule
    '''
    ''' NOTE: This function ONLY looks for general Procedures.
    '''     It does NOT look for Get, Let, or Set Procedures.
    '''     It MIGHT NOT WORK properly against Class Modules!
    '''
    Dim rt As Scripting.Dictionary
    Dim mx As Long
    Dim dx As Long
    Dim fw As Long
    Dim ck As String
    Dim tx As String
    
    Set rt = New Scripting.Dictionary
    With cm
        If .Parent.Type = vbext_ct_StdModule Then
            mx = .CountOfLines
            dx = 1 + .CountOfDeclarationLines
            'Debug.Print .Lines(1, .CountOfDeclarationLines) & "'''"
            
            Do While dx < mx
                fw = dx
                Do While fw < mx
                    ck = .ProcOfLine(fw, vbext_pk_Proc)
                    If Len(ck) > 0 Then
                        dx = .ProcStartLine(ck, vbext_pk_Proc)
                        fw = .ProcCountLines(ck, vbext_pk_Proc)
                        tx = .Lines(dx, fw)
                        rt.Add ck, tx
                        dx = dx + fw
                        fw = mx
                    Else
                        fw = fw + 1
                    End If
                Loop
            Loop
        Else
            'Debug.Print "!!!WARNING!!! Module " & .Parent.Name & " is NOT a standard module!"
            'Stop
        End If
        
        'Stop
    End With
    Set dcOfVbProcs_obs2023_0419 = rt
End Function

Public Function dcOfVbModules( _
    vb As VBIDE.VBProject _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim vc As VBIDE.VBComponent
    
    Set rt = New Scripting.Dictionary
    With vb
        If .Protection = vbext_pp_none Then
            For Each vc In .VBComponents
                With vc
                    rt.Add .Name, .CodeModule
                    'rt.Add .Name, dcOfVbProcs(.CodeModule)
                    'rt.Add .Name, txOfVbModule(.CodeModule)
                End With
            Next
        Else
            rt.Add "<PROTECTED>", New Scripting.Dictionary
            'Join(Array( _
                "''' ", _
                "''' VB Project " & .Name & " Locked", _
                "''' Component Definitions and Data Unavailable", _
                "''' " _
            ), vbNewLine)
        End If
    End With
    Set dcOfVbModules = rt
End Function

Public Function dcOfVbProcsFlat( _
    pj As VBIDE.VBProject _
) As Scripting.Dictionary
    '''
    ''' dcOfVbProcsFlat - generate Dictionary
    '''     of collected text of all procedures
    '''     in each module of given VBProject,
    '''     keyed by procedure name, or by
    '''     combination of module and procedure
    '''     name when more than one procedure
    '''     of same name is found
    '''
    ''' NOTE[2023.04.19.1256] the compromise
    '''     noted above is NOT ideal.
    '''     As the purpose of this function
    '''     is to produce a FLAT list
    '''     of procedure names for quick
    '''     searching purposes, the need
    '''     to modify duplicate names
    '''     for is likely to make it
    '''     difficult or impractical
    '''     to find all possible matches
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim dc As Scripting.Dictionary
    Dim kyMd As Variant
    Dim kyPr As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dcOfVbModules(pj)
    For Each kyMd In .Keys
        With dcOfVbProcs(obOf(.Item(kyMd))) 'dcOb
        For Each kyPr In .Keys
            If rt.Exists(kyPr) Then
                Debug.Print ; 'breakpoint landing
                'Stop
                'going to need a better way
                'to handle this situation
                'but for now...
                rt.Add kyMd & "." & kyPr, .Item(kyPr)
            Else
                rt.Add kyPr, .Item(kyPr)
            End If
        Next
        End With
    Next
    End With
    
    Set dcOfVbProcsFlat = rt
'send2clipBdWin10 dumpLsKeyVal(dcOfVbProcsFlat(ThisWorkbook.VBProject))
End Function

Public Function dcOfVbProjects( _
    pjs As VBIDE.VBProjects _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim vb As VBIDE.VBProject
    
    Set rt = New Scripting.Dictionary
    With pjs
        For Each vb In pjs
            rt.Add vb.Filename, _
            dcOfVbModules(vb)
        Next
    End With
    Set dcOfVbProjects = rt
End Function

Public Function dcTxOfVbModule( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    If dc Is Nothing Then
    Else
        With dc
        If .Count > 0 Then
            For Each ky In .Keys
                rt.Add ky, txOfVbModule( _
                    obVbCodeMod(.Item(ky)) _
                )
            Next
        End If
    End With
    End If
    
    Set dcTxOfVbModule = rt
End Function

Public Function dcTxOfVbProjMods( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dc
    If .Count > 0 Then
        For Each ky In .Keys
            rt.Add ky, dcTxOfVbModule( _
                dcOb(.Item(ky)) _
            ) 'obVbProject
        Next
    End If
    End With
    
    Set dcTxOfVbProjMods = rt
End Function

Public Function vbTextOfProcInDict( _
    nm As String, dc As Scripting.Dictionary _
) As String
    '''
    ''' vbTextOfProcInDict -- Retrieve text from Dictionary
    '''
    ''' This Function's name is probably
    ''' WAY unnecessarily specific.
    '''
    ''' The Function itself simply returns the String
    ''' found under the supplied key variable 'nm',
    ''' or an empty String if none is found. This is
    ''' a fairly general type of Function, one which
    ''' could be named far more generically.
    '''
    ''' dcItemIfPresent looks a likely candidate,
    ''' although it might be a bit TOO general...
    '''
    If dc Is Nothing Then
        'Recursive call option removed.
        'See text of vbTextOfProcInProject
        'for details on security issue.
        '
        vbTextOfProcInDict = "" ' _
        vbTextOfProcInDict( _
            nm, dcOfVbProcsFlat( _
            ThisWorkbook.VBProject _
        ))
    Else
        vbTextOfProcInDict = CStr( _
            dcItemIfPresent( _
            dc, nm, vbString _
        ))
        'With dc
        'If .Exists(nm) Then
        '    vbTextOfProcInDict = .Item(nm)
        'Else
        '    vbTextOfProcInDict = ""
        'End If
        'End With
    End If
End Function

Public Function vbTextOfProcIn( _
    nm As String, cm As VBIDE.CodeModule _
) As String
    vbTextOfProcIn = _
    vbTextOfProcInDict( _
        nm, dcOfVbProcs(cm) _
    )
End Function

Public Function vbTextOfProcInProject( _
    nm As String, pj As VBIDE.VBProject _
) As String
    '''
    ''' vbTextOfProcInProject
    '''
    ''' NOTE: In order to use this Function
    '''         from an external library,
    '''         the option has been removed
    '''         to call itself recursively
    '''         against ThisWorkbook. Since
    '''         ThisWorkbook would be the
    '''         library itself, a call against
    '''         it could result in a breach
    '''         of security.
    '''
    If pj Is Nothing Then
        vbTextOfProcInProject = ""
    Else
        vbTextOfProcInProject = _
        vbTextOfProcInDict( _
            nm, dcOfVbProcsFlat(pj) _
        )
    End If
End Function

Public Function send2clipBd(src As Variant) As Variant
    Dim ck As String
    
    ck = send2clipBdWin10(src)
    
    'With New MSForms.DataObject
    '    .SetText src
    '    .PutInClipboard
    '
    '    .GetFromClipboard
    '    On Error Resume Next
    '    Do
    '    Err.Clear
    '    ck = .GetText
    '    If Err.Number Then
    '        If MsgBox( _
    '            Join(Array( _
    '                "Error Getting Text from DataObject.", _
    '                "A simple retry will usually succeed.", _
    '                "", "Go ahead and retry?" _
    '            ), vbNewLine), _
    '            vbYesNo, "Retry GetText?" _
    '        ) = vbNo Then
    '            Err.Clear
    '        End If
    '    End If
    '    Loop Until Err.Number = 0
    '    On Error GoTo 0
    'End With
    
    If ck = src Then
    Else
        Stop
    End If
    
    send2clipBd = src
End Function

Public Function getFromClipBd(Optional fmt As Variant = 1) As Variant
    ''  1 is the value of CF_TEXT, one of the clipboard format
    ''  enums which SHOULD be defined, but apparently aren't.
    ''  That is the effective default format used by GetText,
    ''  if none is given
    Dim rt As Variant
    With New MSForms.DataObject
        .GetFromClipboard
        rt = .GetText(fmt)
    End With
    getFromClipBd = rt
End Function

Public Function dumpKeyedText( _
    d1 As Scripting.Dictionary, _
    d2 As Scripting.Dictionary _
) As String
    ''  Extract values from second dictionary
    ''  filed under keys from FIRST dictionary.
    ''  Theory is, the keys will always be
    ''  retrieved in the same order, as long as
    ''  no changes have been made between runs.
    ''
    ''  By supplying the same dictionary for
    ''  both d1 and d2, that dictionary's
    ''  content can be extracted, and then
    ''  a different d2's content can be
    ''  extracted in the same order.
    Dim ky As Variant
    Dim rt As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    For Each ky In d1.Keys
        rt.Add "{" & ky & "}" & vbNewLine & d2(ky), 1
    Next
    dumpKeyedText = Join(rt.Keys, vbNewLine)
End Function

