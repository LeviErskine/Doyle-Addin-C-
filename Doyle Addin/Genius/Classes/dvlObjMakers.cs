

Private Const txVersion As String = "dvlObjMakers REV[2022.03.16.0930]"

Public Function nu_wkgCls0( _
    Optional AiDoc As Inventor.Document = Nothing _
) As wkgCls0
    With New wkgCls0
        Set nu_wkgCls0 = .Using(AiDoc)
    End With
End Function

Public Function nu_gnsIfcAiDoc() As gnsIfcAiDoc
    Set nu_gnsIfcAiDoc = New gnsIfcAiDoc
End Function

Public Function nuILogicIfc( _
    Optional Using As Inventor.Document = Nothing _
) As iLogicIfc
    With New iLogicIfc 'ifcVault
    If .RuleSource Is Using Then
        If Using Is Nothing Then
            Set nuILogicIfc = _
            .WithRulesIn(Using)
        Else
        Set nuILogicIfc = .Itself
        End If
    ElseIf Using Is Nothing Then
        Set nuILogicIfc = .Itself
    Else
        Set nuILogicIfc = _
        .WithRulesIn(Using)
    End If
    End With
'Debug.Print txDumpLs(dcOb(nuILogicIfc().Apply("ilRuleText", nuAiNameValMap()).Item("OUT")).Keys)
'Debug.Print nuSelectorV2().WithList(nuILogicIfc().Apply("ilRuleText", nuAiNameValMap).Item("OUT")).GetReply()
'{mod from Debug.Print nuSelectorV2().WithList(nuIfcVault().iLogCall("ilRuleText", New Scripting.Dictionary).Item("OUT")).GetReply()}
End Function

Public Function nu_fmIfcTest04A( _
    Optional About As Scripting.Dictionary = Nothing _
) As fmIfcTest04A
    With New fmIfcTest04A
        Set nu_fmIfcTest04A _
        = .Using(About)
    End With
'Debug.Print nu_fmIfcTest04A().SeeUser().Version()
End Function

Public Function nu_fmIfcMatlQty01() As fmIfcMatlQty01
    Set nu_fmIfcMatlQty01 = New fmIfcMatlQty01
'Debug.Print nu_fmIfcMatlQty01().SeeUser().Version()
End Function

Public Function nu_FmGetList() As fmGetList
    Set nu_FmGetList = New fmGetList
End Function

Public Function newFmTest0() As fmTest0
    Set newFmTest0 = New fmTest0
End Function
'Debug.Print newFmTest0().ft0g0f0(aiDocument(ThisApplication.ActiveDocument).Thumbnail)

Public Function newFmTest1() As fmTest1
    Set newFmTest1 = New fmTest1
End Function

Public Function newFmTest2() As fmTest2
    Set newFmTest2 = New fmTest2
End Function

Public Function nuAiBoxData() As aiBoxData
    ''  Using "blank" version at this point
    Set nuAiBoxData = nuAiBoxDataRC0()
End Function
'Debug.Print nuAiBoxData().UsingInches().Sorted(aiDocPart(aiDocActive()).ComponentDefinition.RangeBox).Dump(0)

'''
''' TESTING SECTION
'''

Public Function tstFmTest1()
    Dim ky As Variant
    Dim nm As String
    
    'nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-047.ipt"
    nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-072.ipt"
    'nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-050.ipt"
    'nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-051.ipt"
    
    With newFmTest1()
        If .AskAbout( _
            ThisApplication.Documents.ItemByName(nm) _
        ) = vbYes Then
            With .ItemData
                For Each ky In .Keys
                    Debug.Print ky, .Item(ky)
                Next
                Stop
            End With
        Else
        End If
    End With
End Function

'''
''' VERSION / VARIANT SECTION
'''

Public Function nuAiBoxDataRC1(arg1 As Variant, _
    Optional UseInches As Long = -1 _
) As aiBoxData
    Dim ob As Object
    Dim rt As aiBoxData
    
    If UseInches < 0 Then
        If IsMissing(arg1) Then
        ElseIf IsObject(arg1) Then
        Else
        End If
    Else
    End If
    
    With New aiBoxData
        Set rt = .UsingInches(UseInches)
    End With
    
    Set nuAiBoxDataRC1 = rt
End Function

Public Function nuAiBoxDataRC0() As aiBoxData
    Set nuAiBoxDataRC0 = New aiBoxData
End Function

'''
''' END of MODULE dvlObjMakers
'''
Public Function dvlObjMakers() As String
    dvlObjMakers = txVersion
End Function
