

Private cn          As ADODB.Connection
Private rs          As ADODB.Recordset
Private rsFamBuy    As ADODB.Recordset
Private fdItem      As ADODB.Field
Private fdFamily    As ADODB.Field
Private fd          As ADODB.Field
'Private fd          As ADODB.Field

'''
''' NOTE: the following SQL text constants
''' are left over from the initial design
''' of this class (under a different name).
'''
''' Their code will remain in place until
''' such time as their value may be better
''' ascertained. Assuming they remain useful,
''' they should be exported to a separate
''' library module for storage of SQL source.
'''
Private Const sql01 As String = "" _
    & "select Item, Family " _
    & "from vgMfiItems " _
    & "" _
    & ""

Private Const sql02 As String = "" _
    & "Select F.Family, F.Description1, " _
    & "F.DefaultPlanningId As pln, " _
    & "F.ProductCategory As cat " _
    & "From vgMfiFamilies F " _
    & "Where F.Type = 'R' " _
    & "And F.FamilyGroup = 'PARTS' " _
    & ""

Public Function Props(AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    If dc Is Nothing Then
        Set Props = Props(AiDoc, _
        New Scripting.Dictionary)
    'ElseIf AiDoc Is Nothing Then
    ''' REV[2022.03.18.1111] <-(seriously!)
    '''     disable branch for void AiDoc
    '''     should not be necessary since
    '''     fall-through branch should
    '''     return dc regardless
    '''     '
    '    Set Props = dc
    Else
        Set Props = _
        assyProps(aiDocAssy(AiDoc), _
        partProps(aiDocPart(AiDoc), _
        gnrlProps(AiDoc, dc)))
    End If
End Function

Private Function gnrlProps( _
    AiDoc As Inventor.Document, _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' gnrlProps -- derived 2023.01.13 from partProps
    '''     to collect (and set) Properties applicable
    '''     to both Parts and Assemblies
    '''     '
    '''     original intent to remove Family assignment
    '''     from Part Property section to a more general
    '''     applicable to both Parts and Assemblies,
    '''     however, that's proving a more challenging
    '''     task than anticipated
    '''     '
    '''     presently just a stub, pending further review
    '''
    If AiDoc Is Nothing Then
        Set gnrlProps = dc
    Else
        'With AiDoc
        '    With .PropertySets
        '        With .Item(gnDesign)
        '        End With
        '    End With
        'End With
        Set gnrlProps = dc
    End If
End Function

Private Function partProps( _
    AiDoc As Inventor.PartDocument, _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    If AiDoc Is Nothing Then
        Set partProps = dc
    Else
        Debug.Print ; 'Breakpoint Landing
        Set partProps = _
        dcGeniusPropsPart( _
        AiDoc, dc)
        Debug.Print ; 'Breakpoint Landing
    End If
End Function

Private Function assyProps( _
    AiDoc As Inventor.AssemblyDocument, _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    If AiDoc Is Nothing Then
        Set assyProps = dc
    Else
        Debug.Print ; 'Breakpoint Landing
        Set assyProps = _
        dcGeniusPropsAssy( _
        AiDoc, dc)
        Debug.Print ; 'Breakpoint Landing
    End If
End Function

Private Sub Class_Initialize()
    'Set cn = cnGnsDoyle()
    'Set rs = cn.Execute(sql01)
    'With rs
    '    With .Fields
    '        Set fdItem = .Item("Item")
    '        Set fdFamily = .Item("Family")
    '    End With
    'End With
    
    'Set rsFamBuy = cn.Execute(sql02)
    'With rsFamBuy
    '    With .Fields
    '    End With
    'End With
End Sub

'''
''' OBSOLETE SECTION
'''
''' All functions below this comment
''' are left over from the initial
''' effort to create some form of
''' Genius-oriented interface to
''' Autodesk Inventor Documents,
''' and particularly Part Documents.
'''
''' The TestXX functions, originally Public,
''' have been rendered Private to hide them
''' from any client processes. Their code
''' remains in place pending possible use
''' in some future process(es).
'''
''' They should at some point be removed,
''' once their value is better established,
''' and any useful portions incorporated
''' into appropriate procedures.
'''

Private Function Test01(invDoc As Inventor.PartDocument) As Inventor.BOMStructureEnum
    ''' Present Role: Categorize Part Document
    '''
    '''
    Dim nmFamily As String
    Dim bomStruct As Inventor.BOMStructureEnum
    '
    '''

    With invDoc
        nmFamily = .PropertySets(gnDesign).Item(pnFamily).Value
        
        If .ComponentDefinition.IsContentMember Then
            If .ComponentDefinition.BOMStructure _
            = kPurchasedBOMStructure Then
                bomStruct = kPurchasedBOMStructure
                nmFamily = "D-HDWR"
            Else
                Stop
            End If
        Else 'ry to identify other purchased part
            If InStr(1, invDoc.FullFileName, _
                "\Doyle_Vault\Designs\purchased\" _
            ) > 0 Then 'this is PROBABLY a purchased part
                bomStruct = kPurchasedBOMStructure
            ElseIf g0f0(nmFamily) = kPurchasedBOMStructure Then
                'this is almost certainly a purchased part
                bomStruct = kPurchasedBOMStructure
            Else 'we'll assume it's NOT purchased.
                bomStruct = kDefaultBOMStructure
                'Use Default to indicate
                'NON purchased parts.
                'We can determine ACTUAL
                'BOM structure elsewhere.
            End If
        End If
        ''
        '''''
        
    End With
    Test01 = bomStruct
    '''
    '
    Debug.Print ; 'Landing Line for Debug use. Do not disable.
    '''
    '''
    '''
End Function

Private Function Test02(invDoc As Inventor.PartDocument) As Inventor.BOMStructureEnum
    ''' Present Role: Categorize Part Document
    '''
    '''
    Dim rt As Scripting.Dictionary
'    ''
    Dim aiPropsUser As Inventor.PropertySet
    Dim aiPropsDesign As Inventor.PropertySet
'    ''
    Dim prFamily    As Inventor.Property
    'Dim prPartNum   As Inventor.Property
'    Dim prRawMatl   As Inventor.Property 'pnRawMaterial
'    Dim prRmUnit    As Inventor.Property 'pnRmUnit
'    Dim prRmQty     As Inventor.Property 'pnRmQty
'    ''
    Dim nmFamily As String
'    Dim mtFamily As String
'    ''' UPDATE[2018.05.30]:
'    '''     Rename variable Family to nmFamily
'    '''     to minimize confusion between code
'    '''     and comment text in searches.
'    '''     Also add variable mtFamily
'    '''     for raw material Family name
'    Dim pnStock As String
'    Dim qtUnit As String
    Dim bomStruct As Inventor.BOMStructureEnum
'    Dim ck As VbMsgBoxResult
'    '
'    '''
'
    With invDoc
        ' Get Property Sets
        'With .PropertySets
'            Set aiPropsUser = .Item(gnCustom)
            'Set aiPropsDesign = .Item(gnDesign)
        'End With
        
        ' Get Custom Properties
'        Set prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
'        Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
'        Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
    
        ' Family property is from Design, NOT Custom set
        'Set prFamily = aiGetProp(aiPropsDesign, pnFamily)
        'Set prPartNum = aiGetProp(aiPropsDesign, pnPartNum)
        'nmFamily = prFamily.Value
        nmFamily = .PropertySets(gnDesign).Item(pnFamily).Value
        
        If .ComponentDefinition.IsContentMember Then
            If .ComponentDefinition.BOMStructure _
            = kPurchasedBOMStructure Then
                bomStruct = kPurchasedBOMStructure
                nmFamily = "D-HDWR"
            Else
                Stop
            End If
        Else 'Try to identify other purchased part
            If InStr(1, invDoc.FullFileName, _
                "\Doyle_Vault\Designs\purchased\" _
            ) > 0 Then
                'this is PROBABLY a purchased part
                bomStruct = kPurchasedBOMStructure
            ElseIf g0f0(nmFamily) = kPurchasedBOMStructure Then
                'this is almost certainly a purchased part
                bomStruct = kPurchasedBOMStructure
            Else 'we'll assume it's NOT purchased.
                bomStruct = kNormalBOMStructure
                'If .SubType = guidSheetMetal Then
                    'It's PROBABLY sheet metal, BUT
                    'it might be something else
                'Else
                'End If
            End If
        End If
        ''
        '''''
        
'        With .ComponentDefinition
'            '''''
'            'Get GeniusMass
'            With .MassProperties
'                Set rt = dcWithProp( _
'                    aiPropsUser, pnMass, _
'                    cvMassKg2LbM * .Mass, rt _
'                )
'            End With
'            'Will want to move this elsewhere
'            'Should focus here on categorization
        'End With
    End With
    Test02 = bomStruct
    '''
    '
    Debug.Print ; 'Landing Line for Debug use. Do not disable.
    '''
    '''
    '''
End Function

Private Function g0f0(f As String) As Inventor.BOMStructureEnum
    With rsFamBuy
        .Filter = "Family = '" & f & "'"
        If .BOF Or .EOF Then
            g0f0 = kDefaultBOMStructure
        Else
            g0f0 = kPurchasedBOMStructure
        End If
    End With
End Function

