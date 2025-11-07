

Implements kyPick

Private pk As kyPick

Private Sub Class_Initialize()
    Set pk = New kyPick
End Sub

Public Function dcFor(Item As Variant) As Scripting.IDictionary
    Dim ob As Inventor.Document
    
    Set ob = aiDocument(obOf(Item))
    If ob Is Nothing Then
        Set dcFor = pk.dcFor(0)
    Else
        If ob.DocumentType = kPartDocumentObject Then
            With dcAiPropsInSet(ob.PropertySets.Item(gnCustom))
                If .Exists("RM") Then
                    Set dcFor = pk.dcFor(ob)
                Else
                    Set dcFor = pk.dcFor(0)
                End If
            End With
        Else
            Set dcFor = pk.dcFor(0)
        End If
    End If
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

Private Function kyPick_AfterScanning(dSrc As Scripting.IDictionary) As kyPick
    Dim ky As Variant
    
    With dSrc
        For Each ky In .Keys
            With dcFor(.Item(ky))
                If .Exists(ky) Then
                    Stop
                Else
                    .Add ky, dSrc.Item(ky)
                End If
            End With
        Next
    End With
    
    Set kyPick_AfterScanning = Me
End Function

'''
''' kyPick Implementation code follows
'''
Private Function kyPick_DcFor(Item As Variant) As Scripting.IDictionary
    Set kyPick_DcFor = dcFor(Item)
End Function

Private Function kyPick_DcIn() As Scripting.IDictionary
    Set kyPick_DcIn = dcIn()
End Function

Private Function kyPick_DcOut() As Scripting.IDictionary
    Set kyPick_DcOut = dcOut()
End Function

Public Function Itself() As kyPick
    Set Itself = Me
End Function

Private Function kyPick_Itself() As kyPick
    Set kyPick_Itself = Me.Itself
End Function

Private Function kyPick_WithInDc(Dict As Scripting.IDictionary) As kyPick
    Set kyPick_WithInDc = WithInDc(Dict)
End Function

Private Function kyPick_WithOutDc(Dict As Scripting.IDictionary) As kyPick
    Set kyPick_WithOutDc = WithOutDc(Dict)
End Function
