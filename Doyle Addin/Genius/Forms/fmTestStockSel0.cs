
Attribute VB_Name = "fmTestStockSel0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private cn As ADODB.Connection
Private rsFam As ADODB.Recordset
Private rsItm As ADODB.Recordset

Private Sub lbxFamily_Change()
    With Me
        rsItm.Filter = "Family = '" & .lbxFamily.Value & "'"
        .lbxItem.List = m0g3f1(rsItm)
    End With
End Sub

Private Sub UserForm_Initialize()
    cn = cnGnsDoyle()
    With cn
        rsFam = .Execute(Join(Array(
            "select Family, Description1",
            "from vgMfiFamilies",
            "where FamilyGroup = 'RAW'"
        ), " "))
        rsItm = .Execute(Join(Array(
            "Select I.Family, I.Item, I.Description1",
            "From vgMfiItems as I",
            "Inner Join vgMfiFamilies as F",
            "On I.Family = F.Family",
            "Where F.FamilyGroup = 'RAW'"
        ), " "))
    End With
    
    Me.lbxFamily.List = m0g3f1(rsFam)
End Sub
