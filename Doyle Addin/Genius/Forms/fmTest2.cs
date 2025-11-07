VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmTest2 
   Caption         =   "Please Review Item #"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3990
   OleObjectBlob   =   "fmTest2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ad As Inventor.Document

Private psDsn As Inventor.PropertySet
Private psUsr As Inventor.PropertySet

Private dcDsn As Scripting.Dictionary
Private dcUsr As Scripting.Dictionary

Private prFam As Inventor.Property
Private prStk As Inventor.Property

'Private dmFmHt As Long
'Private dmFmWd As Long
''Private dmLbMsHt As Long
Private dmLbMsWd As Long
''Private dmDfFmMsHt As Long
''Private dmDfFmMsWd As Long
Private dmFmHt2cmdTop As Long

Private rtAnswer As VbMsgBoxResult

Public Function AskAbout( _
    Optional AiDoc As Inventor.Document = Nothing, _
    Optional txPre As String = "", _
    Optional txPost As String = "" _
) As VbMsgBoxResult
    '''
    ''' AskAbout -- prompt User for action
    '''     to take on supplied Document
    ''' UPDATE[2021.12.13]
    '''     Document parameter now Optional.
    '''     will attempt to use previously
    '''     registered Document when none
    '''     supplied. Warning/error message
    '''     will be presented if no Document
    '''     is registered OR supplied.
    '''
    Dim pc As stdole.IPictureDisp
    Dim pn As String
    Dim sn As String
    Dim pd As String
    Dim dj As Single 'use to adjust
    '   form height and positions
    '   of command buttons
    
    rtAnswer = vbCancel
    If Not AiDoc Is Nothing Then
        ad = AiDoc
    End If

    If ad Is Nothing Then
        MsgBox "Review or Update requested" _
            & vbNewLine & "but no Document provided!" _
            & vbNewLine & "" _
            & vbNewLine & "" _
        , vbOKOnly, "No Document!"
        rtAnswer = vbNo
    ElseIf aiDocPartFromCCtr(ad) Is Nothing Then 'AiDoc
        ' ad = AiDoc
        With ad
            pc = .Thumbnail
            psDsn = .PropertySets(gnDesign)
            psUsr = .PropertySets(gnCustom)

            dcDsn = dcAiPropsInSet(psDsn)
            dcUsr = dcAiPropsInSet(psUsr)

            prFam = psDsn.Item(pnFamily)
            With dcUsr
                If .Exists(pnRawMaterial) Then
                    prStk = psUsr.Item(pnRawMaterial)
                Else
                    On Error Resume Next
                    Err.Clear()
                    prStk = psUsr.Add("", pnRawMaterial)
                    If Err.Number Then
                        Debug.Print Err.Number, Err.Description
                        Stop
                    Else
                        .Add pnRawMaterial, prStk
                    End If
                    On Error GoTo 0
                End If
            End With

            If Not prStk Is Nothing Then sn = prStk.Value
            pn = psDsn.Item(pnPartNum).Value
            pd = psDsn.Item(pnDesc).Value
        End With

        With Me
            .Caption = "Please Review Item: " & pn

            If pc Is Nothing Then
            Else
                .imThmNail.Picture = pc
            End If

            dj = fmHtAdjust(lblHtAdjust(.lbMsg,
                IIf(Len(txPre) > 0,
                    txPre & vbNewLine & vbNewLine, ""
                ) & Join(Array(pn & ": " & pd,
                    pnCatWebLink & ": " & psDsn.Item(pnCatWebLink).Value,
                    pnMaterial & ": " & psDsn.Item(pnMaterial).Value
                ), vbNewLine & vbNewLine) _
                & IIf(Len(txPost) > 0,
                    vbNewLine & vbNewLine & txPost, ""
                )
            ))
            '.dbFamily.Value = prFam.Value

            .Show 1
        End With
    Else
        MsgBox ad.DisplayName _
            & vbNewLine & "is a Content Center part" _
            & vbNewLine & "and cannot be updated." _
            & vbNewLine & "" _
            & vbNewLine & "" _
        , vbOKOnly, "Can't Update!" 'AiDoc
        rtAnswer = vbYes
    End If

    AskAbout = rtAnswer 'vbYes ' = 1
End Function

Public Function Using(
    AiDoc As Inventor.Document
) As fmTest2
    '''
    ''' NEWMETHOD[2021.12.13]
    ''' Using -- assign supplied Document
    '''     for use in all subsequent calls
    '''     to AskAbout without one.
    '''
    rtAnswer = vbCancel

    If Not AiDoc Is Nothing Then
        ad = AiDoc
    End If

    Using = Me
End Function

Public Function Document(
    Optional AiDoc As Inventor.Document = Nothing
) As Inventor.Document
    '''
    ''' NEWMETHOD[2021.12.13]
    ''' Document -- return currently active Document
    '''
    If AiDoc Is Nothing Then
        Document = ad
    Else
        Document = Me.Using(AiDoc).Document
    End If
End Function

Private Function fmHtAdjust(by As Long) As Single
    Dim cmdTop As Long

    With Me
        .Height = .Height + by

        .cmdLt.Top = .Height - dmFmHt2cmdTop
        .cmdCt.Top = .cmdLt.Top
        .cmdRt.Top = .cmdLt.Top

        fmHtAdjust = .Height
    End With
End Function

Private Function lblHtAdjust(
    lb As MSForms.Label, tx As String
) As Single
    Dim ct As MSForms.Control
    Dim au As Boolean
    Dim wd As Single
    Dim ht As Single

    ct = lb
    With ct
        wd = .Width
        ht = .Height

        With lb
            au = .AutoSize
            .Caption = tx
            .AutoSize = True
            ct.Width = dmLbMsWd
            .AutoSize = au
        End With

        lblHtAdjust = Int(.Height - ht)
    End With
End Function

Private Sub cmdCt_Click()
    rtAnswer = vbNo
    Me.Hide()
End Sub

Private Sub cmdLt_Click()
    rtAnswer = vbYes
    Me.Hide()
End Sub

Private Sub cmdRt_Click()
    rtAnswer = vbCancel
    Me.Hide()
End Sub

Private Sub UserForm_Initialize()
    '''
    With Me
        'dmFmHt = .Height
        'dmFmWd = .Width
        With .lbMsg
            'dmLbMsHt = .Height
            dmLbMsWd = .Width
        End With
        dmFmHt2cmdTop = .Height - .cmdLt.Top
    End With
    'dmDfFmMsWd = dmFmWd - dmLbMsWd
    'dmDfFmMsHt = dmFmHt - dmLbMsHt
    rtAnswer = vbCancel
End Sub

Private Sub UserForm_Click()
    '''
End Sub

Private Sub UserForm_Layout()
    'Stop
End Sub

Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
    Cancel = 1
    Me.Hide()
End Sub

Private Sub UserForm_Terminate()
    'cn.Close
    ' cn = Nothing
End Sub

