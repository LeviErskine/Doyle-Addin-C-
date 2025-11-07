

Public Function lbxPickedStr(lbx As MSForms.ListBox, _
    Optional dlm As String = vbVerticalTab _
) As String
    Dim dw As Long
    Dim mx As Long
    Dim dx As Long
    Dim rt As String
    
    dw = Len(dlm)
    With lbx
        rt = ""
        mx = .ListCount - 1
        For dx = 0 To mx
            If .Selected(dx) Then
                rt = rt & dlm _
                & CStr(.List(dx, 0))
            End If
        Next
        lbxPickedStr = Mid$(rt, 1 + dw)
    End With
End Function

Public Function lbxPicked(lbx As MSForms.ListBox, _
    Optional dlm As String = vbVerticalTab _
) As Variant
    lbxPicked = Split(lbxPickedStr(lbx, dlm), dlm)
End Function

Public Function nuSelector() As fmSelectorList
    Set nuSelector = New fmSelectorList
End Function

Public Function nuSelectorV2() As fmSelectorV2
    Set nuSelectorV2 = New fmSelectorV2
End Function

Public Function nuSelFromDict(dc As Scripting.Dictionary, _
    Optional hOhKay As String = "", Optional mOhKay As String = "", _
    Optional hCancl As String = "", Optional mCancl As String = "", _
    Optional hNoSel As String = "", Optional mNoSel As String = "" _
) As fmSelectorList
    Set nuSelFromDict = nuSelector( _
    ).SetHdrOK(IIf(Len(hOhKay) > 0, hOhKay, "Confirm Selection" _
    )).SetMsgOK(IIf(Len(mOhKay) > 0, mOhKay, Join(Array( _
            "Action will proceed using", "%%%", _
            "(Click CANCEL to quit with no action)" _
        ), vbNewLine) _
    )).SetHdrCancel(IIf(Len(hCancl) > 0, hCancl, "Cancel Operation?" _
    )).SetMsgCancel(IIf(Len(mCancl) > 0, mCancl, Join(Array( _
            "No action will be taken on", "%%%" _
        ), vbNewLine) _
    )).SetHdrNoSelection(IIf(Len(hNoSel) > 0, hNoSel, "No Item Selected!" _
    )).SetMsgNoSelection(IIf(Len(mNoSel) > 0, mNoSel, Join(Array( _
            "Do you wish to cancel the operation?", _
            "(Click NO to return to list)" _
        ), vbNewLine) _
    )).WithList( _
        dc.Keys _
    )
End Function

#If False Then '

Public Function nuSelWkBk( _
    Optional hOhKay As String = "", Optional mOhKay As String = "", _
    Optional hCancl As String = "", Optional mCancl As String = "", _
    Optional hNoSel As String = "", Optional mNoSel As String = "" _
) As fmSelectorList
    Set nuSelWkBk = nuSelector( _
    ).SetHdrOK(IIf(Len(hOhKay) > 0, hOhKay, "Proceed With Update?" _
    )).SetMsgOK(IIf(Len(mOhKay) > 0, mOhKay, Join(Array( _
        "The following workbook will be affected: ", _
        "%%%", "(Click CANCEL to quit with no changes)" _
        ), vbNewLine) _
    )).SetHdrCancel(IIf(Len(hCancl) > 0, hCancl, "Cancel Operation?" _
    )).SetMsgCancel(IIf(Len(mCancl) > 0, mCancl, Join(Array( _
        "No changes will be applied", _
        "to any open workbook." _
        ), vbNewLine) _
    )).SetHdrNoSelection(IIf(Len(hNoSel) > 0, hNoSel, "No Workbook Selected!" _
    )).SetMsgNoSelection(IIf(Len(mNoSel) > 0, mNoSel, Join(Array( _
        "Do you wish to cancel the operation?", _
        "(Click NO to return to list)" _
        ), vbNewLine) _
    )).WithList( _
        lsWorkbooks() _
    )
End Function

Public Function nuSelWkSht(inWkBk As Excel.Workbook, _
    Optional hOhKay As String = "", Optional mOhKay As String = "", _
    Optional hCancl As String = "", Optional mCancl As String = "", _
    Optional hNoSel As String = "", Optional mNoSel As String = "" _
) As fmSelectorList
    Set nuSelWkSht = nuSelector( _
    ).SetHdrOK(IIf(Len(hOhKay) > 0, hOhKay, "Confirm Selection" _
    )).SetMsgOK(IIf(Len(mOhKay) > 0, mOhKay, Join(Array( _
            "Action will proceed using Workwheet: ", "%%%", _
            "(Click CANCEL to quit with no action)" _
        ), vbNewLine) _
    )).SetHdrCancel(IIf(Len(hCancl) > 0, hCancl, "Cancel Operation?" _
    )).SetMsgCancel(IIf(Len(mCancl) > 0, mCancl, Join(Array( _
            "No action will be taken", _
            "to any open workbook." _
        ), vbNewLine) _
    )).SetHdrNoSelection(IIf(Len(hNoSel) > 0, hNoSel, "No Workbook Selected!" _
    )).SetMsgNoSelection(IIf(Len(mNoSel) > 0, mNoSel, Join(Array( _
            "Do you wish to cancel the operation?", _
            "(Click NO to return to list)" _
        ), vbNewLine) _
    )).WithList( _
        dcWkSheets(inWkBk).Keys _
    )
End Function

#End If

