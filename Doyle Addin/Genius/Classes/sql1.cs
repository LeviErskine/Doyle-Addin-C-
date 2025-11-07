

Public Function sqlListValues( _
    dc As Scripting.Dictionary, _
    Optional sep As String = "', '", _
    Optional pfx As String = "('", _
    Optional sfx As String = "')" _
) As String
    '''
    '''
    '''
    sqlListValues = pfx & Join(dc.Keys, sep) & sfx
    '"')," & vbNewLine & vbTab & "('"
End Function

Public Function sqlValSelFromDict( _
    dc As Scripting.Dictionary _
) As String
    '''
    '''
    '''
    sqlValSelFromDict = sqlListValues(dc, _
        "')," & vbNewLine & vbTab & "('", _
        "(select pn from (VALUES ('", _
        "')" & vbNewLine & ") as a(pn))" _
    )
    'With dc
        'sqlValSelFromDict = "(select pn from (VALUES " _
            & vbNewLine & vbTab & "('" _
            & Join(.Keys, "')," & vbNewLine & vbTab & "('") _
            & "')" & vbNewLine & ") as a(pn))" _
            & ""
    'End With
End Function

Public Function sqlValsFromDict( _
    dc As Scripting.Dictionary, _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    '''
    '''
    '''
    sqlValsFromDict = sqlListValues(dc, _
        "')," & vbNewLine & vbTab & "('", _
        "(values ('", "')" & vbNewLine & ") as " _
        & lsName & "(" & fdName & ")" _
    )
End Function

Public Function sqlValsFromAssy( _
    AiDoc As Inventor.Document, _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    '''
    '''
    '''
    Dim dc As Scripting.Dictionary
    Dim ck As String
    
    Set dc = dcRemapByPtNum(dcAiDocComponents(AiDoc))
    
    ck = sqlListValues(dc, _
        "')," & vbNewLine & vbTab & "('", _
        "(values ('", "')" & vbNewLine & ") as " _
        & lsName & "(" & fdName & ")" _
    )
    If sqlValsFromDict(dc) = ck Then
        Stop
    End If
    sqlValsFromAssy = ck
End Function

Public Function q1g0x0(AiDoc As Inventor.Document) As String
    '''
    ''' SQL text function naming convention
    '''     q1 - "q" for "query", with module number
    '''     g1 - "g" for "group" (typical usage)
    '''     x1 - "x" for "text" (stands out better than "t")
    '''
    q1g0x0 = "-- SQL text begins here" _
    & vbNewLine & "" _
    & vbNewLine & "-- SQL text ends here"
    '''
End Function

Public Function q1g1x1( _
    AiDoc As Inventor.Document, _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    '''
    ''' SQL text function naming convention
    '''     q1 - "q" for "query", with module number
    '''     g1 - "g" for "group" (typical usage)
    '''     x1 - "x" for "text" (stands out better than "t")
    '''
    q1g1x1 = "from vgMfiItems i inner join" _
        & vbNewLine & sqlValsFromAssy( _
            AiDoc, lsName, fdName _
        ) & vbNewLine & "on i.Item = " _
        & lsName & "." & fdName
End Function

Public Function q1g1x1v2( _
    AiDoc As Inventor.Document, _
    Optional gnsTbl As String = "vgMfiItems", _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    '''
    ''' SQL text function naming convention
    '''     q1 - "q" for "query", with module number
    '''     g1 - "g" for "group" (typical usage)
    '''     x1 - "x" for "text" (stands out better than "t")
    '''
    
    'REV[2021.08.18] (REVERSED)
    '   changed inner join to right (outer) join
    '   to pick up Inventor Items not (yet) in Genius
    '   REVERSED -- since all returned fields are null,
    '   no information is returned for missing Items.
    q1g1x1v2 = "from " & gnsTbl & " i inner join" _
        & vbNewLine & sqlValsFromAssy( _
            AiDoc, lsName, fdName _
        ) & vbNewLine & "on i.Item = " _
        & lsName & "." & fdName
End Function

Public Function q1g1x1v3( _
    dc As Scripting.Dictionary, _
    Optional gnsTbl As String = "vgMfiItems", _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    '''
    ''' SQL text function naming convention
    '''     q1 - "q" for "query", with module number
    '''     g1 - "g" for "group" (typical usage)
    '''     x1 - "x" for "text" (stands out better than "t")
    '''
    
    'REV[2021.08.18] (REVERSED)
    '   changed inner join to right (outer) join
    '   to pick up Inventor Items not (yet) in Genius
    '   REVERSED -- since all returned fields are null,
    '   no information is returned for missing Items.
    q1g1x1v3 = "from " & gnsTbl & " i inner join" _
        & vbNewLine & sqlValsFromDict( _
            dc, lsName, fdName _
        ) & vbNewLine & "on i.Item = " _
        & lsName & "." & fdName
End Function

Public Function q1g1x2( _
    AiDoc As Inventor.Document, _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    q1g1x2 = "select" & " i." & _
    IIf(False, "*", Join(Array("Item" _
        , "Family", "Description1", "Description3", "Unit" _
        , "Thickness", "Width", "Length" _
        , "Height", "Diameter", "Weight" _
        , "Specification1", "Specification2", "Specification3" _
        , "Specification4", "Specification5", "Specification6" _
        , "Specification7", "Specification8", "Specification9" _
    ), ", i.")) & vbNewLine _
    & q1g1x1(AiDoc, lsName, fdName)
    
    '& ", " & lsName & "." & join(Array( _
        "", "" _
    ), ", " & lsName & ".") _
    '''
    '''
'send2clipBd ConvertToJson(dcDxFromRecSetDc( _
    dcFromAdoRS(cnGnsDoyle().Execute( _
        q1g1x2(ThisApplication.ActiveDocument) _
    )) _
), vbTab)
End Function

Public Function q1g1x2v2( _
    dc As Scripting.Dictionary, _
    Optional gnsTbl As String = "vgMfiItems", _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    q1g1x2v2 = "select" & " i." & _
    IIf(False, "*", Join(Array("Item" _
        , "Family", "Description1", "Description3", "Unit" _
        , "Thickness", "Width", "Length" _
        , "Height", "Diameter", "Weight" _
        , "Specification1", "Specification2", "Specification3" _
        , "Specification4", "Specification5", "Specification6" _
        , "Specification7", "Specification8", "Specification9" _
    ), ", i.")) & vbNewLine _
    & q1g1x1v3(dc, gnsTbl, lsName, fdName)
    'was q1g1x1v2(aiDoc,...
    
    '& ", " & lsName & "." & join(Array( _
        "", "" _
    ), ", " & lsName & ".") _
    '''
    '''
'send2clipBdWin10 ConvertToJson(dcRecSetDcDx4json( _
    dcDxFromRecSetDc(dcFromAdoRS( _
        cnGnsDoyle().Execute( _
            q1g1x2v2(aiDocActive()) _
        )) _
    ) _
), vbTab)
'Debug.Print txDumpLs(dcKeysMissing( _
    dcRemapByPtNum( _
        dcAiDocComponents(aiDocActive()) _
    ), _
    dcOb(dcDxFromRecSetDc(dcFromAdoRS( _
        cnGnsDoyle().Execute( _
            q1g1x2v2(aiDocActive()) _
        ) _
    )).Item("Item")) _
).Keys)
End Function

Public Function q1g2x1( _
    AiDoc As Inventor.Document, _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    '''
    ''' SQL text function naming convention
    '''     q1 - "q" for "query", with module number
    '''     g1 - "g" for "group" (typical usage)
    '''     x1 - "x" for "text" (stands out better than "t")
    '''
    q1g2x1 = "from vgIcoBillOfMaterials b inner join" _
        & vbNewLine & sqlValsFromAssy( _
            AiDoc, lsName, fdName _
        ) & vbNewLine & "on i.Item = " _
        & lsName & "." & fdName
End Function

Public Function q1g2x2( _
    AiDoc As Inventor.Document, _
    Optional lsName As String = "ls", _
    Optional fdName As String = "it" _
) As String
    q1g2x2 = "select" & " b." & _
    IIf(False, "*", Join(Array( _
        "Product", "ItemOrder", "Item", _
        "QuantityInConversionUnit", _
        "ConversionUnit", _
        "ItemType", "Reserved" _
    ), ", b.")) & vbNewLine _
    & q1g2x1(AiDoc, lsName, fdName)
    ''' _
    '''
    
    '& ", " & lsName & "." & join(Array( _
        "", "" _
    ), ", " & lsName & ".") _
    '''
    '''
'send2clipBd ConvertToJson(dcDxFromRecSetDc( _
    dcFromAdoRS(cnGnsDoyle().Execute( _
        q1g2x2(ThisApplication.ActiveDocument) _
    )) _
), vbTab)
End Function

Public Function sqlSelAiPurch01fromTextV01(txList As String) As String
    '''
    '''
    '''
    sqlSelAiPurch01fromTextV01 = "-- " _
    & vbNewLine & "with" _
        & vbNewLine & vbTab & "ls as " & txList _
    & vbNewLine & "select" _
        & vbNewLine & vbTab & "ls.pn, it.Type, it.Family" _
    & vbNewLine & "from" _
        & vbNewLine & vbTab & "ls inner join vgMfiItems as it" _
        & vbNewLine & vbTab & "on ls.pn = it.Item" _
    & vbNewLine & "where" _
        & vbNewLine & vbTab & "it.Type = 'R'" _
        & vbNewLine & vbTab & "or it.Family in (" _
        & "'D-HDWR', 'D-PTO', 'D-PTS', 'R-PTO', 'R-PTS'" _
        & ")" _
    & "" _
    & ""
End Function

Public Function sqlSelAiPurch01fromTextV02( _
    txList As String _
) As String
    '''
    '''
    '''
    Dim n0 As Long
    Dim n1 As Long
    Dim s0 As String
    Dim s1 As String
    Dim a0 As Variant
    
    n0 = InStr(1, txList, "'")
    s0 = Mid$(txList, 1 + n0)
    n0 = InStr(1, s0, "'")
    n1 = InStr(1 + n0, s0, "'") '- n0 + 1
    If n1 > 0 Then
        n1 = n1 - n0 + 1
        s1 = Mid$(s0, n0, n1)
        a0 = Split(s0, s1)
        n0 = UBound(a0)
        a0(n0) = Split(a0(n0), "'")(0)
    Else
        a0 = Array(Left$(s0, n0 - 1))
    End If
    's0 = Join(a0, "', '")
    
    Debug.Print ;
    
    sqlSelAiPurch01fromTextV02 = "-- " _
    & vbNewLine & "select" _
        & vbNewLine & vbTab & "it.Item, it.Type, it.Family" _
    & vbNewLine & "from" _
        & vbNewLine & vbTab & "vgMfiItems as it" _
    & vbNewLine & "where" _
        & vbNewLine & vbTab & "it.Item in ('" _
            & Join(a0, "', '") & "')" _
        & vbNewLine & vbTab & "and (it.Type = 'R'" _
        & vbNewLine & vbTab & "or it.Family in ('" _
            & Join(Array("D-HDWR", _
                "D-PTO", "D-PTS", _
                "R-PTO", "R-PTS" _
            ), "', '") _
        & "'))" _
    & "" _
    & ""
    '
    '& vbNewLine & "with" _
        & vbNewLine & vbTab & "ls as " & txList _
    '
End Function

Public Function sqlSelAiPdParts01fromTextV01( _
    txList As String _
) As String
    '''
    '''
    '''
    Dim n0 As Long
    Dim n1 As Long
    Dim s0 As String
    Dim s1 As String
    Dim a0 As Variant
    
    n0 = InStr(1, txList, "'")
    s0 = Mid$(txList, 1 + n0)
    n0 = InStr(1, s0, "'")
    n1 = InStr(1 + n0, s0, "'") - n0 + 1
    s1 = Mid$(s0, n0, n1)
    a0 = Split(s0, s1)
    n0 = UBound(a0)
    a0(n0) = Split(a0(n0), "'")(0)
    's0 = Join(a0, "', '")
    
    Debug.Print ;
    
    sqlSelAiPdParts01fromTextV01 = "-- " _
    & vbNewLine & "select" _
        & vbNewLine & vbTab & Join(Array( _
            "bm.Product", "bm.Item", _
            "pd.Type pType", "pd.Family pFam", _
            "it.Type iType", "it.Family iFam" _
        ), ", ") _
    & vbNewLine & "from" _
        & vbNewLine & vbTab & "vgIcoBillOfMaterials bm Inner Join" _
        & vbNewLine & vbTab & "vgMfiItems pd On bm.Product = pd.Item" _
        & vbNewLine & vbTab & "Left Join vgMfiItems it On bm.Item = it.Item" _
    & vbNewLine & "where" _
        & vbNewLine & vbTab & "bm.Product in " _
            & txList _
        & "" _
    & "" _
    & ""
    '
    '(' Join(a0, "', '") ')
    '& vbNewLine & "with" _
        & vbNewLine & vbTab & "ls as " & txList _
    '
End Function

Public Function sqlSelTestFromText( _
    txList As String _
) As String
    '''
    '''
    '''
    
    sqlSelTestFromText _
    = sqlSelAiPdParts01fromTextV01(txList)
End Function

Public Function sqlSelAiPurch01fromText( _
    txList As String _
) As String
    '''
    '''
    '''
    
    sqlSelAiPurch01fromText _
    = sqlSelAiPurch01fromTextV02(txList)
End Function

Public Function sqlSelAiPdParts01fromText( _
    txList As String _
) As String
    '''
    '''
    '''
    
    sqlSelAiPdParts01fromText _
    = sqlSelAiPdParts01fromTextV01(txList)
End Function

Public Function sqlSelAiPurch01fromDict( _
    dc As Scripting.Dictionary _
) As String
    '''
    '''
    '''
    sqlSelAiPurch01fromDict _
    = sqlSelAiPurch01fromText(Join(Split( _
        sqlValSelFromDict(dc), _
        vbNewLine), vbNewLine & vbTab _
    ))
End Function

Public Function sqlSelAiPdParts01fromDict( _
    dc As Scripting.Dictionary _
) As String
    '''
    '''
    '''
    sqlSelAiPdParts01fromDict _
    = sqlSelAiPdParts01fromText( _
        sqlListValues(dc) _
    )
End Function

Public Function sqlSelAiPurch01fromAssy( _
    AiDoc As Inventor.Document _
) As String
    '''
    '''
    '''
    sqlSelAiPurch01fromAssy _
    = sqlSelAiPurch01fromDict( _
        dcRemapByPtNum( _
        dcAiDocComponents(AiDoc) _
        ) _
    )
End Function

Public Function sqlSelAiPdParts01fromAssy( _
    AiDoc As Inventor.Document _
) As String
    '''
    '''
    '''
    sqlSelAiPdParts01fromAssy _
    = sqlSelAiPdParts01fromDict( _
        dcRemapByPtNum( _
        dcAiDocComponents(AiDoc) _
        ) _
    )
End Function

