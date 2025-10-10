class libFmSelectors
{
    public string lbxPickedStr(MSForms.ListBox lbx, string dlm = Constants.vbVerticalTab)
    {
        long dw;
        long mx;
        long dx;
        string rt;

        dw = Strings.Len(dlm);
        {
            var withBlock = lbx;
            rt = "";
            mx = withBlock.ListCount - 1;
            for (dx = 0; dx <= mx; dx++)
            {
                if (withBlock.Selected(dx))
                    rt = rt + dlm;/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
            }
            lbxPickedStr = Mid(rt, 1 + dw);
        }
    }

    public Variant lbxPicked(MSForms.ListBox lbx, string dlm = Constants.vbVerticalTab)
    {
        lbxPicked = Split(lbxPickedStr(lbx, dlm), dlm);
    }

    public fmSelectorList nuSelector()
    {
        nuSelector = new fmSelectorList();
    }

    public fmSelectorV2 nuSelectorV2()
    {
        nuSelectorV2 = new fmSelectorV2();
    }

    public fmSelectorList nuSelFromDict(Scripting.Dictionary dc, string hOhKay = "", string mOhKay = "", string hCancl = "", string mCancl = "", string hNoSel = "", string mNoSel = "")
    {
        nuSelFromDict = nuSelector().SetHdrOK(Interaction.IIf(Strings.Len(hOhKay) > 0, hOhKay, "Confirm Selection")).SetMsgOK(IIf(Strings.Len(mOhKay) > 0, mOhKay, Join(Array("Action will proceed using", "%%%", "(Click CANCEL to quit with no action)"), Constants.vbNewLine))).SetHdrCancel(Interaction.IIf(Strings.Len(hCancl) > 0, hCancl, "Cancel Operation?")).SetMsgCancel(IIf(Strings.Len(mCancl) > 0, mCancl, Join(Array("No action will be taken on", "%%%"), Constants.vbNewLine))).SetHdrNoSelection(Interaction.IIf(Strings.Len(hNoSel) > 0, hNoSel, "No Item Selected!")).SetMsgNoSelection(IIf(Strings.Len(mNoSel) > 0, mNoSel, Join(Array("Do you wish to cancel the operation?", "(Click NO to return to list)"), Constants.vbNewLine))).WithList(dc.Keys);
    }
}