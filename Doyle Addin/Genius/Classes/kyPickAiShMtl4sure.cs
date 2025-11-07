

Implements kyPick

Private pk As kyPick
Private Const txVersion As String = "kyPickAiShMtl4sure v0.0.0.1 [2022.03.08.1336]"
''' prior Versions
'''     ""
'''
''' kyPick Implementation code follows
'''

Private Function kyPick_Itself() As kyPick
    Set kyPick_Itself = Me
End Function


Private Function kyPick_WithInDc( _
    Dict As Scripting.IDictionary _
) As kyPick
    Set pk = pk.WithInDc(Dict)
    Set kyPick_WithInDc = Me
End Function

Private Function kyPick_WithOutDc( _
    Dict As Scripting.IDictionary _
) As kyPick
    Set pk = pk.WithOutDc(Dict)
    Set kyPick_WithOutDc = Me
End Function


Private Function kyPick_AfterScanning( _
    dSrc As Scripting.IDictionary _
) As kyPick
    Dim dcI As Scripting.Dictionary
    Dim dcO As Scripting.Dictionary
    Dim dCk As Scripting.Dictionary
    Dim ky As Variant
    
    With pk.AfterScanning(dSrc)
        Set dcI = .dcIn
        Set dcO = .dcOut
    End With
    
    With dcI: For Each ky In .Keys
        Set dCk = kyPick_DcFor(.Item(ky))
        If dCk Is dcI Then 'should be okay
            'don't need to do anything here
            Debug.Print ; 'Breakpoint Landing
        Else
            With dCk
            If .Exists(ky) Then
                Stop
            Else
                .Add ky, dcI.Item(ky)
            End If: End With
        End If
    Next: End With
    Set pk = pk.WithInDc(dcKeysMissing(dcI, dcO))
    
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
    Dim ob As Inventor.PartDocument '.Document
    
    Set ob = aiDocPart(aiDocument(obOf(Item)))
    If ob Is Nothing Then
        Set kyPick_DcFor = pk.dcFor(0)
    Else
        Set kyPick_DcFor = g0f1( _
            ob.ComponentDefinition _
        )
    End If
End Function
'''
'''
''' General Class handling code follows
'''

Private Sub Class_Initialize()
    Set pk = New kyPickAiSheetMetal 'kyPick
End Sub
'''
'''
''' kyPickAiShMtl4sure Class
''' implementation code follows
'''

Public Function Itself() As kyPick
    Set Itself = Me
End Function


Public Function WithInDc( _
    Dict As Scripting.Dictionary _
) As kyPick
    Set WithInDc = kyPick_WithInDc(Dict)
End Function

Public Function WithOutDc( _
    Dict As Scripting.Dictionary _
) As kyPick
    Set WithOutDc = kyPick_WithOutDc(Dict)
End Function


Public Function dcIn() As Scripting.Dictionary
    Set dcIn = pk.dcIn
End Function

Public Function dcOut() As Scripting.Dictionary
    Set dcOut = pk.dcOut
End Function


Public Function AfterScanning( _
    dSrc As Scripting.Dictionary _
) As kyPick
    Set AfterScanning = kyPick_AfterScanning(dSrc)
End Function


Public Function dcFor(Item As Variant) As Scripting.IDictionary
    Set dcFor = kyPick_DcFor(Item)
End Function
'''
'''
''' Internal support code follows
'''

Private Function g0f0( _
    ob As Inventor.PartDocument _
) As Scripting.Dictionary
    If ob Is Nothing Then
        Set g0f0 = pk.dcFor(0)
    Else
        Set g0f0 = g0f1(ob.ComponentDefinition)
    End If
End Function

Private Function g0f1( _
    ob As Inventor.PartComponentDefinition _
) As Scripting.Dictionary
    If TypeOf ob Is Inventor.SheetMetalComponentDefinition Then
        Set g0f1 = g0f2(ob)
    Else
        Set g0f1 = pk.dcFor(0)
    End If
End Function

Private Function g0f2( _
    ob As Inventor.SheetMetalComponentDefinition _
) As Scripting.Dictionary
    Dim smThk As Double
    Dim fpHgt As Double
    Dim dfRns As Double
    
    If ob Is Nothing Then
        Set g0f2 = pk.dcFor(0)
    Else
        With ob
        If .HasFlatPattern Then 'MIGHT be
            ''  check stated thickness...
            smThk = .Thickness.Value
            'Debug.Print "Thickness: " & CStr(smThk)
            
            ''  against flat pattern height
            With nuAiBoxData().UsingBox( _
                .FlatPattern.RangeBox _
            )
                'Debug.Print .Dump()
                Debug.Print ; 'Breakpoint Landing
                fpHgt = .SpanZ
            End With
            
            dfRns = fpHgt - smThk
            If Abs(dfRns) < 0.001 Then 'close enough
                ''  assume it's valid
                Set g0f2 = pk.dcFor(ob.Document)
            Else 'something's screwy here
                ''  assume likely not
                'Stop
                Set g0f2 = pk.dcFor(0)
            End If
        Else 'likely not
            Set g0f2 = pk.dcFor(0)
        End If: End With
    End If
End Function
'''
'''
''' Version code follows
'''

Public Function Version() As String
    Version = txVersion
End Function
'''
''' End of Module
'''
