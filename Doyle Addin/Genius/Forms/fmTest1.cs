VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmTest1 
   Caption         =   "Please Review"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7950
   OleObjectBlob   =   "fmTest1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmTest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cn As ADODB.Connection
Private rsFam As ADODB.Recordset
Private rsPrt As ADODB.Recordset
Private rsItm As ADODB.Recordset

Private dc As Scripting.Dictionary
Private ad As Inventor.Document

Private psDsn As Inventor.PropertySet
Private psUsr As Inventor.PropertySet

Private dcDsn As Scripting.Dictionary
Private dcUsr As Scripting.Dictionary

Private prFam As Inventor.Property
Private prStk As Inventor.Property
Private prThk As Inventor.Property

Public Function AskAbout( _
    AiDoc As Inventor.Document, _
    Optional txMsg As String = "" _
) As VbMsgBoxResult
    Dim pc As stdole.IPictureDisp
    Dim ck As VbMsgBoxResult
    Dim pn As String    'part number
    Dim sn As String    'material (stock) number
    Dim sf As String    'material (stock) family
    Dim pd As String    'part description
    Dim df As Single

    ad = AiDoc
    With ad
        On Error Resume Next
        Err.Clear()
        pc = .Thumbnail
        If Err.Number = 0 Then 'we're good
        Else 'no image
        End If
        On Error GoTo 0

        psDsn = .PropertySets(gnDesign)
        psUsr = .PropertySets(gnCustom)

        dcDsn = dcAiPropsInSet(psDsn)
        dcUsr = dcAiPropsInSet(psUsr)

        prFam = psDsn.Item(pnFamily)

        ''  Get Sheet Metal Thickness Property
        prThk = aiPropShtMetalThickness(ad)
        ''  NOTE: Function returns Nothing
        ''      if Part is NOT Sheet Metal!

        With dcUsr
            If .Exists(pnRawMaterial) Then
                prStk = psUsr.Item(pnRawMaterial)
            Else
                On Error Resume Next
                Err.Clear()
                prStk = psUsr.Add("", pnRawMaterial)
                If Err.Number Then
                    Stop
                Else
                    .Add pnRawMaterial, prStk
                End If
                On Error GoTo 0
            End If
        End With

        ''' REV[2022.04.28.1615]
        ''' added initializtion of Dictionary dc
        ''' with initial raw material setting.
        ''' sn now assigned from the Dictionary.
        ''' NOTE: probably want to  initial
        ''' values in a separate "recovery"
        ''' Dictionary to be restored if
        ''' the User chooses to cancel.
        ''' Also, see function/method dcUpd.
        ''' looks like it gets called when
        ''' something changes. Easy to miss!
        dc.Item(pnRawMaterial) = prStk.Value
        sn = dc.Item(pnRawMaterial)
        pn = psDsn.Item(pnPartNum).Value
        pd = psDsn.Item(pnDesc).Value
    End With

    With Me
        .Caption = "Please Review Part Number: " & pn

        If pc Is Nothing Then
        Else
            .imThmNail.Picture = pc
        End If

        With .lbMsg
            .Caption = pn & ": " & pd _
                & vbNewLine & txMsg & IIf(Len(txMsg) > 0, vbNewLine, "") _
                & ft1g0f0(pnCatWebLink, psDsn.Item(pnCatWebLink)) & vbNewLine _
                & ft1g0f0(pnMaterial, psDsn.Item(pnMaterial)) & vbNewLine _
                & ft1g0f0(pnThickness, prThk) & vbNewLine _
                & ""
                '& vbNewLine _
                & vbNewLine & pnThickness & ": " & psUsr.Item(pnThickness).Value _
                '
        End With
        df = mdl1g1f2(.lbMsg)
        If df > 0 Then
            mdl1g1f3.lbMtFamily, 0, df
            mdl1g1f3.lbxFamily, 0, df
        End If

        .dbFamily.Value = prFam.Value

        If Len(sn) > 0 Then
            With cn.Execute(
                "select Family from vgMfiItems where Item = '" _
                & Replace(sn, "'", "''") & "'"
            )
                ''' REV[2022.08.19.1359]
                ''' temporarily replacing direct use of sn
                ''' with call to Replace single quotes
                ''' in string with doubled single quotes
                ''' (NOT double quotes!) to "escape" the
                ''' character in a string value.
                ''' '
                ''' will ultimately want to produce some
                ''' sort of 'handler' to preprocess values
                ''' for use in SQL commands to avoid errors
                ''' that arise from this sort of thing.
                If .BOF Or .EOF Then
                    sf = ""
                Else
                    sf = .Fields(0).Value
                End If
                ''' NOTE[2022.04.28.1625]
                ''' THOUGHT Material Family was also
                ''' added to Dictionary dc, but believe
                ''' that's actually the PART Family.
            End With

            If Len(sf) = 0 Then 'selected Material
                'EITHER doesn't have a Family,
                'OR is not (yet) in Genius.
                'SO, let's just ...
                sf = "DSHEET" 'as a default!
            Else 'no need to do anything
                'comments below are from when
                'this was the NO family block.
                'retain for now, until we're sure
                'this are working all right.

                'this SHOULDN'T happen, so hopefully
                'things won't come to this...
                'Stop
                'worry about the handler later
            End If

            On Error Resume Next
            Err.Clear()
            .lbxFamily.Value = sf
            If Err.Number Then
                Debug.Print "FAILED TO  MATERIAL FAMILY " & sf
                ck = MsgBox(Join(Array(
                    "Part Number " & pn, "uses Material " & sn _
                    , "which is a" & IIf(
                        InStr(1, "AEIOU", UCase$(Left$(sf, 1))),
                        "n ", " "
                    ) & sf & " Item." _
                    , "" _
                    , "This interface does not presently" _
                    , "support Materials from this Family." _
                    , "" _
                    , "You might not be able to find the correct" _
                    , "Material for this Part, and might wish" _
                    , "to avoid changing it here." _
                    , "" _
                    , "Do you wish to proceed anyway?"
                ), vbNewLine),
                    vbYesNoCancel + vbExclamation + vbDefaultButton2,
                    "Material Family not Supported"
                )
                If ck = vbCancel Then
                    Stop
                End If
            Else
                Err.Clear()
                .lbxItem.Value = sn

                ''' REV[2022.05.06.1329]
                ''' added intermediate error handler
                ''' to capture failure in Material
                ''' Family selector to adopt new Value.
                ''' it re-implements process of Event
                ''' handler Sub lbxFamily_Change
                ''' against variable 'sf' directly
                ''' in an effort to force population
                ''' of Material list.
                If Err.Number Then
                    Debug.Print() ; 'Breakpoint Landing
                    Err.Clear()
                    rsItm.Filter = "Family = '" & sf & "'"
                    .lbxItem.List = m0g3f1(rsItm)
                    .lbxItem.Value = sn
                End If
                ''' something MIGHT have happened
                ''' to prevent normal Value update
                ''' when lbxFamily is  above.
                ''' further investigation may be
                ''' warranted.

                If Err.Number Then
                    Debug.Print "FAILED TO  MATERIAL " & sn
                    ck = MsgBox(Join(Array("!!WARNING!!", "" _
                        , "Active Material " & sn _
                        , "for Part Number " & pn _
                        , "could NOT be selected," _
                        , "and might be unavailable." _
                        , "" _
                        , "You might wish to avoid" _
                        , "making Material changes" _
                        , "to this Part here." _
                        , "" _
                        , "Do you wish to proceed anyway?"
                    ), vbNewLine),
                        vbYesNoCancel + vbExclamation,
                        "Active Material Not Found!"
                    )
                    If ck = vbCancel Then
                        Stop
                    End If
                Else
                    ck = vbYes
                    lbxItem_Change()
                    'lbxFamily_Change
                    rsItm.Filter = "Family = '" & sf & "'"
                    .lbxItem.List = m0g3f1(rsItm)
                End If
            End If
            On Error GoTo 0
        Else
            ck = vbYes
        End If

        If ck = vbYes Then
            .Show 1
        End If
    End With
    AskAbout = ck 'vbYes ' = 1
End Function

Private Function ft1g0f0(
    pn As String, pr As Inventor.Property
) As String
    If pr Is Nothing Then
        ft1g0f0 = ""
    Else
        ft1g0f0 = vbNewLine & pn & ": " & pr.Value
    End If
End Function

Private Sub dbFamily_Change()
    Debug.Print dcUpd(pnFamily, dbFamily.Value)
End Sub
'Me.lbxItem.ColumnWidths = "84 pt;6 pt;180 pt"
'Me.lbxItem.ColumnWidths = "84 pt;48 pt;216 pt"

Private Sub lbMsg_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Stop
End Sub

Private Sub lbxFamily_Change()
    With Me
        rsItm.Filter = "Family = '" & .lbxFamily.Value & "'"
        .lbxItem.List = m0g3f1(rsItm)
    End With
End Sub

Public Function ItemData() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant

    rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            rt.Add ky, .Item(ky)
        Next
    End With
    ItemData = rt
End Function

Public Function Synch() As Scripting.Dictionary
    With dc
        If .Exists(pnFamily) Then prFam.Value = dc.Item(pnFamily)
        If .Exists(pnRawMaterial) Then prStk.Value = dc.Item(pnRawMaterial)
    End With

    Synch = Me.ItemData
End Function

Private Function dcUpd(ky As String, vl As Variant) As String
    Dim rt As String

    If IsNull(vl) Then
        dcUpd = dcUpd(ky, "")
    Else
        With dc
            If .Exists(ky) Then
                rt = CStr(.Item(ky))
                .Item(ky) = vl
                dcUpd = "CHANGE[" & ky & "] FROM '" & rt _
                & "' TO '" & CStr(.Item(ky)) & "'"
            Else
                .Add ky, vl
            dcUpd = "[" & ky & "] TO '" _
                & CStr(.Item(ky)) & "'"
            End If
        End With
    End If
End Function

Private Sub lbxItem_Change()
    Debug.Print dcUpd(pnRawMaterial, lbxItem.Value)
End Sub

Private Sub UserForm_Initialize()
    dc = New Scripting.Dictionary
    cn = cnGnsDoyle()

    With cn
        ' rsFam = .Execute(Join(Array( _
            "select Family, Description1", _
            "from vgMfiFamilies", _
            "order by Family" _
        ), " ")) ', _
            "where FamilyGroup = 'RAW'"
         rsPrt = .Execute(Join(Array(
            "select Family, FamilyGroup, Description1",
            "from vgMfiFamilies",
            "order by Family"
        ), vbNewLine)) ', _
            "where FamilyGroup = 'PARTS'"
         rsItm = .Execute(Join(Array(
            "Select I.Item, I.Family, I.Description1, I.Specification1",
            "From vgMfiItems as I",
            "Inner Join vgMfiFamilies as F",
            "On I.Family = F.Family",
            "Where F.FamilyGroup = 'RAW'",
            "order by Family, Item"
        ), " "))
    End With

    With Me
        rsPrt.Filter = "FamilyGroup = 'RAW'"
        .lbxFamily.List = m0g3f1(rsPrt) 'rsFam

        rsPrt.Filter = "FamilyGroup = 'PARTS'"
        .dbFamily.List = m0g3f1(rsPrt)
    End With
End Sub

Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
    Cancel = 1
    Me.Hide()
End Sub

Private Sub UserForm_Terminate()
    cn.Close
    cn = Nothing
End Sub

