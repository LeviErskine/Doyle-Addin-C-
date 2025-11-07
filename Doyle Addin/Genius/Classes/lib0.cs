

' Measurement Unit Conversion Factors
Public Const cvArSqCm2SqFt  As Double = 0.00107639
                            ' 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
                            '
                            ' /  1ft | 1in    \2     2                2
                            '( ------+-------- ) * cm  = 0.00107639 ft
                            ' \ 12in | 2.54cm /
Public Const cvMassKg2LbM   As Double = 2.20462
Public Const cvLenIn2cm     As Double = 2.54

'''
Public Const guidRegPart    As String = "{4D29B490-49B2-11D0-93C3-7E0706000000}"
Public Const guidSheetMetal As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
Public Const guidDesignAccl As String = "{BB8FE430-83BF-418D-8DF9-9B323D3DB9B9}"
Public Const guidPipingSgmt As String = "{4D39D5F1-0985-4783-AA5A-FC16C288418C}"
Public Const guidILogicAdIn As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"
'''
Public Const guidRegAssy    As String = "{E60F81E1-49B3-11D0-93C3-7E0706000000}"
Public Const guidWeldment   As String = "{28EC8354-9024-440F-A8A2-0E0E55D635B0}"

'''
Public Const guidPrSetSumm  As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}" 'Summary Information (Inventor Summary Information)
Public Const guidPrSetDocu  As String = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}" 'Document Summary Information (Inventor Document Summary Information)
Public Const guidPrSetTrkg  As String = "{32853F0F-3444-11D1-9E93-0060B03C1CA6}" 'Design Tracking Properties (Design Tracking Properties)
Public Const guidPrSetUser  As String = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" 'User Defined Properties (Inventor User Defined Properties)
Public Const guidPrSetCLib  As String = "{B9600981-DEE8-4547-8D7C-E525B3A1727A}" 'Content Library Component Properties (Content Library Component Properties)
Public Const guidPrSetCCtr  As String = "{CEAAEE65-91D8-444E-ACBA-BE54A5FB9D4D}" 'ContentCenter (ContentCenter)
'Public Const guidPrSet____  As String = "{00000000-0000-0000-0000-000000000000}" 'Display Name (Name)
'''

Public Const gnDesign       As String = "Design Tracking Properties"
Public Const pnMaterial     As String = "Material"          '
Public Const pnPartNum      As String = "Part Number"       '
Public Const pnStockNum     As String = "Stock Number"      '
Public Const pnFamily       As String = "Cost Center"       '
Public Const pnDesc         As String = "Description"       '
Public Const pnCatWebLink   As String = "Catalog Web Link"  '

Public Const gnCustom       As String = "Inventor User Defined Properties"
Public Const pnMass         As String = "GeniusMass"    '
Public Const pnRawMaterial  As String = "RM"            '
Public Const pnRmQty        As String = "RMQTY"         '
Public Const pnRmUnit       As String = "RMUNIT"        '(replaces "RMUOM")
'                                                       '
Public Const pnArea         As String = "Extent_Area"   '
Public Const pnLength       As String = "Extent_Length" '
Public Const pnWidth        As String = "Extent_Width"  '
'
Public Const pnThickness    As String = "Thickness"     '
'

Public Function vbProjectLocal() As VBIDE.VBProject
    Set vbProjectLocal = ThisDocument.VBAProject.VBProject
End Function

Public Function cnGnsDoyle() As ADODB.Connection
    Dim rt As ADODB.Connection
    ''' NOTE[2021.12.08]:
    '''     Might consider make rt a Static Object.
    '''     If it can be created and opened just once
    '''     during a run, this could potentially save
    '''     a LOT of overhead from repeated open/close
    '''     operations, and might save a little load
    '''     on the server, as well.
    
    Set rt = New ADODB.Connection
    With rt
        .Provider = "SQLOLEDB" '"SQLNCLI11"
        .CursorLocation = adUseClient
        .Open "Data Source=DOYLE-ERP02", "GeniusReporting", "geniusreporting"
        .DefaultDatabase = "DoyleDB"
        '.Close
    End With
    Set cnGnsDoyle = rt
End Function

Public Function dcIvObjTypeEnum() As Scripting.Dictionary
    Dim dc As Scripting.Dictionary
    Dim en As Inventor.ObjectTypeEnum
    
    Set dc = New Scripting.Dictionary
    
    With dc
        en = k3dAViewObject
        en = kAliasFreeformFeatureObject
        en = kAliasFreeformFeatureProxyObject
        en = kAliasFreeformFeaturesObject
        en = kAnalysisManagerObject
        en = kAnalyticEdgeWorkAxisDefObject
        en = kAngleConstraintObject
        en = kAngleConstraintProxyObject
        en = kAngleExtentObject
        en = kAngleiMateDefinitionObject
        '.Add kUnknownDocumentObject, "kUnknownDocumentObject"
        '.Add kSATFileDocumentObject, "kSATFileDocumentObject"
        '.Add kPresentationDocumentObject, "kPresentationDocumentObject"
        '.Add kPartDocumentObject, "kPartDocumentObject"
        '.Add kNoDocument, "kNoDocument"
        '.Add kForeignModelDocumentObject, "kForeignModelDocumentObject"
        '.Add kDrawingDocumentObject, "kDrawingDocumentObject"
        '.Add kDesignElementDocumentObject, "kDesignElementDocumentObject"
        '.Add kAssemblyDocumentObject, "kAssemblyDocumentObject"
    End With
    
    Set dcIvObjTypeEnum = dc
End Function
'
'=====

Public Function dcIvDocTypeEnum() As Scripting.Dictionary
    Dim dc As Scripting.Dictionary
    Dim en As Inventor.DocumentTypeEnum
    
    Set dc = New Scripting.Dictionary
    
    With dc
        .Add kUnknownDocumentObject, "kUnknownDocumentObject"
        .Add kSATFileDocumentObject, "kSATFileDocumentObject"
        .Add kPresentationDocumentObject, "kPresentationDocumentObject"
        .Add kPartDocumentObject, "kPartDocumentObject"
        .Add kNoDocument, "kNoDocument"
        .Add kForeignModelDocumentObject, "kForeignModelDocumentObject"
        .Add kDrawingDocumentObject, "kDrawingDocumentObject"
        .Add kDesignElementDocumentObject, "kDesignElementDocumentObject"
        .Add kAssemblyDocumentObject, "kAssemblyDocumentObject"
    End With
    
    Set dcIvDocTypeEnum = dc
End Function
'
'=====

Public Function txDumpLs(ls As Variant, _
    Optional bk As String = vbNewLine _
) As String
    Dim rt As Variant
    Dim tx As Variant
    Dim mx As Long
    Dim bs As Long
    Dim dx As Long
    
    If IsArray(ls) Then
        bs = LBound(ls)
        mx = UBound(ls)
        If bs > mx Then
            txDumpLs = ""
        Else
        ReDim rt(bs To mx)
        For dx = LBound(ls) To mx
            rt(dx) = txDumpLs(ls(dx))
        Next
        txDumpLs = Join(rt, bk)
        End If
    ElseIf IsObject(ls) Then
        If TypeOf ls Is Scripting.Dictionary Then
            txDumpLs = txDumpLs(ls.Keys)
        Else
        End If
    Else
        txDumpLs = CStr(ls)
    End If
End Function

Public Sub lsDump(ls As Variant, _
    Optional bk As String = vbNewLine _
)
    Debug.Print txDumpLs(ls, bk)
End Sub

'''
''' The following is copied over from the Excel project file libExt.xlsm
''' to provide a means of dumping Key-Value pairs from a Dictionary.
'''

Public Function dumpLsKeyVal(dc As Scripting.Dictionary _
    , Optional dlmField As String = "," _
    , Optional dlmLine As String = vbNewLine _
    , Optional nullTxt As String = "<null>" _
    , Optional emptyTx As String = "<empty>" _
) As String
    Dim d2 As Scripting.Dictionary
    Dim ky As Variant
    Dim vl As Variant
    'Dim rt As String
    
    'rt = ""
    If dc Is Nothing Then
        dumpLsKeyVal = ""
    Else
        Set d2 = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                'rt = rt & ky & "," & .Item(ky) & vbNewLine
                
                vl = Array(.Item(ky))
                ''  Any values which have
                ''  no direct String conversion
                ''  are replaced with String defaults
                ''  (   user supplied, or see
                ''      Function declaration above )
                If IsNull(vl(0)) Then vl = nullTxt
                If IsEmpty(vl(0)) Then vl = emptyTx
                'If IsMissing(vl) Then vl = ""
                'If IsError(vl) Then vl = ""
                If IsObject(vl(0)) Then
                    If vl(0) Is Nothing Then
                        vl = "<ob:Nothing>"
                    Else
                        vl = "<ob:" & TypeName(vl(0)) & ">"
                    End If
                End If
                If IsArray(vl) Then
                    If IsArray(vl(0)) Then
                        vl = "<array>"
                    Else
                        vl = vl(0)
                    End If
                End If
                
                d2.Add Join(Array(ky, vl), dlmField), ky
                ''  Note that Key value might also
                ''  require String default replacement,
                ''  as well. Won't address this unless
                ''  and until it becomes an issue.
                ''
                ''  If it DOES, the defaulting process
                ''  will probably be broken out
                ''  into its own function
            Next
        End With
        dumpLsKeyVal = Join(d2.Keys, dlmLine)
    End If
End Function

