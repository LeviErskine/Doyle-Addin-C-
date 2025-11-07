

Private Const f01 As String = "#,##0.000"
'Private Const f02 As String = "#,##0.0000 '"

Private bx As Inventor.Box
Private mn As Inventor.Point
Private mx As Inventor.Point

Private sc As Double

Private Sub Class_Initialize()
    sc = 1#
End Sub

Public Property Set Box(ThisBox As Inventor.Box)
    Set bx = ThisBox
    Set mn = bx.MinPoint
    Set mx = bx.MaxPoint
End Property

Public Property Get Box() As Inventor.Box
    Set Box = bx
End Property

Public Function UsingBox( _
    ThisOne As Inventor.Box _
) As aiBoxData
    Set Me.Box = ThisOne
    Set UsingBox = Me
End Function

Public Function UsingOrBox( _
    ThisOne As Inventor.OrientedBox _
) As aiBoxData
    Set bx = ThisApplication.TransientGeometry.CreateBox()
    
    With ThisOne
    bx.Extend ThisApplication.TransientGeometry.CreatePoint( _
        .DirectionOne.length, _
        .DirectionTwo.length, _
        .DirectionThree.length _
    )
    End With
    
    'Set Me.Box = ThisOne
    Set UsingOrBox = Me.UsingBox(bx)
'Debug.Print nuAiBoxData().UsingInches.UsingOrBox(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences.Item(1).OrientedMinimumRangeBox).Dump()
End Function

Public Function UsingBoxOb( _
    ThisOne As Object _
) As aiBoxData
    If ThisOne Is Nothing Then
        Set UsingBoxOb = Me
    ElseIf TypeOf ThisOne Is Inventor.Box Then
        Set UsingBoxOb = UsingBox(ThisOne)
    ElseIf TypeOf ThisOne Is Inventor.OrientedBox Then
        Set UsingBoxOb = UsingOrBox(ThisOne)
    Else
        Set UsingBoxOb = Me
    End If
End Function

Public Function UsingModel( _
   ThisOne As Inventor.Document, _
   Optional Oriented As Long = 0 _
) As aiBoxData
    Set UsingModel _
        = UsingPart(aiDocPart(ThisOne), Oriented _
        ).UsingAssy(aiDocAssy(ThisOne), Oriented _
    )
End Function

Public Function UsingPart( _
   ThisOne As Inventor.PartDocument, _
   Optional Oriented As Long = 0 _
) As aiBoxData
    If ThisOne Is Nothing Then
        Set UsingPart = Me
    Else
        With ThisOne.ComponentDefinition
        Set UsingPart = UsingBoxOb(IIf(Oriented = 0, _
            .RangeBox, .OrientedMinimumRangeBox _
        ))
        End With
    End If
End Function

Public Function UsingAssy( _
   ThisOne As Inventor.AssemblyDocument, _
   Optional Oriented As Long = 0 _
) As aiBoxData
    If ThisOne Is Nothing Then
        Set UsingAssy = Me
    Else
        With ThisOne.ComponentDefinition
        Set UsingAssy = UsingBoxOb(IIf(Oriented = 0, _
            .RangeBox, .OrientedMinimumRangeBox _
        ))
        End With
        'Set UsingAssy = UsingBox( _
        ThisOne.ComponentDefinition.RangeBox _
        )
    End If
End Function

Public Function SortingDims( _
    Optional ThisBox As Inventor.Box = Nothing _
) As aiBoxData
    If ThisBox Is Nothing Then
        If bx Is Nothing Then
            Set SortingDims = Me
        Else
            Set SortingDims = SortingDims(bx)
        End If
    Else
        Set Me.Box = aiBoxSortDown(ThisBox)
        Set SortingDims = Me
    End If
End Function

Private Function Span( _
    ptMin As Double, _
    ptMax As Double _
) As Double
    Span = sc * (ptMax - ptMin)
End Function

Public Function SpanX() As Double
    SpanX = Span(mn.X, mx.X)
End Function

Public Function SpanY() As Double
    SpanY = Span(mn.Y, mx.Y)
End Function

Public Function SpanZ() As Double
    SpanZ = Span(mn.Z, mx.Z)
End Function

Public Function SpansXYZ() As Double()
    Dim rt(2) As Double
    
    rt(0) = SpanX
    rt(1) = SpanY
    rt(2) = SpanZ
    
    SpansXYZ = rt
End Function

Public Function SpansOrdered() As Double()
    SpansOrdered = sort3dimsUp(SpanX, SpanY, SpanZ)
End Function

Public Function UsingInches(Optional Yes As Long = 1) As aiBoxData
    If Yes Then sc = 1 / 2.54 Else sc = 1
    Set UsingInches = Me
End Function

Public Function Dump(Optional Form As Long = 0) As String
    Dump = ""
    'ConvertToJson(nuDcPopulator().Setting("X SPAN", Format$(me.SpanX, "#,##0.0000 '")).Setting("Y SPAN", Format$(me.SpanY, "#,##0.0000 '")).Setting("Z SPAN", Format$(me.SpanZ, "#,##0.0000 '")).Dictionary,vbTab)
    'ConvertToJson(nuDcPopulator().Setting("X SPAN", Round(me.SpanX,4)).Setting("Y SPAN", Round(me.SpanY,4)).Setting("Z SPAN", Round(me.SpanZ,4)).Dictionary,vbTab)
    Select Case Form
    Case 67518582
        With nuDcPopulator().Setting("X SPAN", Format$(Me.SpanX, "#,##0.0000 '")).Setting("Y SPAN", Format$(Me.SpanY, "#,##0.0000 '")).Setting("Z SPAN", Format$(Me.SpanZ, "#,##0.0000 '")).Dictionary
            '''
        End With
    Case Else
        Dump = "X SPAN" & vbTab & "Y SPAN" & vbTab & "Z SPAN" _
            & vbNewLine & Format(Me.SpanX, f01) _
            & vbTab & Format(Me.SpanY, f01) _
            & vbTab & Format(Me.SpanZ, f01)
    End Select
End Function

Public Function Dictionary( _
    Optional Form As Long = 3 _
) As Scripting.Dictionary
    '''
    ''' Dictionary -- return Dictionary of dimensions
    '''     keyed according to Form, a sum of:
    '''     1 - "X", "Y", "Z", per Model
    '''     2 - magnitudes "Min", "Mid", "Max"
    '''         (note that sorting keys in descending order
    '''          produces values sorted in ascending order)
    '''     3 - BOTH sets of keys (1 + 2)
    '''
    ''' REV[2022.08.31.1444] Method Dictionary
    ''' added to Class to support extraction
    ''' of Dictionary Object for data export
    ''' (see dcGnsPtProps_Rev20220830_inProg)
    '''
    Dim rt As Scripting.Dictionary
    Dim dm() As Double
    
    If (Form And 3) = 0 Then
        Set rt = Dictionary(3)
    Else
        Set rt = New Scripting.Dictionary
        
        With rt
            If Form And 1 Then 'add XYZ entries
                .Add "X", SpanX()
                .Add "Y", SpanY()
                .Add "Z", SpanZ()
            End If
            
            If Form And 2 Then 'add Min, Mid, Max entries
                dm = SpansOrdered()
                .Add "Min", dm(0)
                .Add "Mid", dm(1)
                .Add "Max", dm(2)
            End If
        End With
    End If
    
    Set Dictionary = rt
End Function

