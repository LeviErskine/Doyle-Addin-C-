

Implements kyPick

Private pk As kyPick
'''
'''
''' kyPick Implementation code follows
'''

Private Function kyPick_Itself() As kyPick
    Set kyPick_Itself = Me.Itself
End Function


Private Function kyPick_WithInDc( _
    Dict As Scripting.IDictionary _
) As kyPick
    Set kyPick_WithInDc = WithInDc(Dict)
End Function

Private Function kyPick_WithOutDc( _
    Dict As Scripting.IDictionary _
) As kyPick
    Set kyPick_WithOutDc = WithOutDc(Dict)
End Function


Private Function kyPick_AfterScanning( _
    dSrc As Scripting.IDictionary _
) As kyPick
    Dim ky As Variant
    
    With dSrc: For Each ky In .Keys
        With dcFor(.Item(ky))
        If .Exists(ky) Then
            Stop
        Else
            .Add ky, dSrc.Item(ky)
        End If: End With
    Next: End With
    
    Set kyPick_AfterScanning = Me
End Function


Private Function kyPick_DcIn() As Scripting.IDictionary
    Set kyPick_DcIn = dcIn()
End Function

Private Function kyPick_DcOut() As Scripting.IDictionary
    Set kyPick_DcOut = dcOut()
End Function


Private Function kyPick_DcFor( _
    Item As Variant _
) As Scripting.IDictionary
    Set kyPick_DcFor = dcFor(Item)
End Function
'''
'''
''' General Class handling code follows
'''

Private Sub Class_Initialize()
    Set pk = New kyPick
End Sub
'''
'''
''' kyPickAiDocPurchased Class
''' implementation code follows
'''

Public Function Itself() As kyPick
    Set Itself = Me
End Function


Public Function WithInDc( _
    Dict As Scripting.Dictionary _
) As kyPick
    Set pk = pk.WithInDc(Dict)
    Set WithInDc = Me
End Function

Public Function WithOutDc( _
    Dict As Scripting.Dictionary _
) As kyPick
    Set pk = pk.WithOutDc(Dict)
    Set WithOutDc = Me
End Function


Public Function AfterScanning( _
    dSrc As Scripting.Dictionary _
) As kyPick
    Set AfterScanning = kyPick_AfterScanning(dSrc)
End Function


Public Function dcIn() As Scripting.Dictionary
    Set dcIn = pk.dcIn
End Function

Public Function dcOut() As Scripting.Dictionary
    Set dcOut = pk.dcOut
End Function


Public Function dcFor(Item As Variant) As Scripting.IDictionary
    Dim ck As Inventor.BOMStructureEnum
    Dim ob As Inventor.Document
    Dim pr As Inventor.Property
    ''' REV[2022.03.08.1021]
    '''     Added BOMStructureEnum variable ck
    '''     to collect BOMStructureEnum for each
    '''     relevant Document type, and consolidate
    '''     BOMStructureEnum check to one block
    '''     following Doc type accommodation.
    
    Set ob = aiDocument(obOf(Item))
    
    If ob Is Nothing Then
        ck = kDefaultBOMStructure
    ElseIf ob.DocumentType = kPartDocumentObject Then
        ck = aiDocPart(ob).ComponentDefinition.BOMStructure
    ElseIf ob.DocumentType = kAssemblyDocumentObject Then
        ck = aiDocAssy(ob).ComponentDefinition.BOMStructure
    Else
        ck = kDefaultBOMStructure
    End If
    
    If ck = kPurchasedBOMStructure Then
        Set dcFor = pk.dcFor(ob)
    Else
        ''' REV[2022.03.08.1038]
        '''     Additional checks on Item
        '''     Family and File Location
        '''     NOTE that this is more of
        '''     a "soft" identification
        '''     of likely purchased parts,
        '''     and might or might not be
        '''     appropriate to apply.
        With ob
            Set pr = .PropertySets.Item( _
                gnDesign).Item(pnFamily _
            )
            If InStr(1, ob.FullFileName, _
                "\Doyle_Vault\Designs\purchased\" _
            ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                "|" & pr.Value & "|" _
            ) > 0 Then
                Set dcFor = pk.dcFor(ob)
            Else
                Set dcFor = pk.dcFor(0)
            End If
        End With
    End If
End Function

