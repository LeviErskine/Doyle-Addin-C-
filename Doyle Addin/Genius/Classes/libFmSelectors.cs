using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class libFmSelectors
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="lbx"></param>
    /// <param name="dlm"></param>
    /// <returns></returns>
    public static string lbxPickedStr(ListBox lbx, string dlm = Constants.vbVerticalTab)
    {
        long dw = Strings.Len(dlm);
        {
            var rt = "";
            long mx = lbx.Items.Count - 1;
            for (long dx = 0; dx <= mx; dx++)
            {
                if (lbx.GetSelected((int)dx))
                    rt = rt + dlm + (lbx.Items[(int)dx]?.ToString() ?? "");
            }

            return Strings.Mid(rt, (int)(1 + dw));
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="lbx"></param>
    /// <param name="dlm"></param>
    /// <returns></returns>
    public static dynamic lbxPicked(ListBox lbx, string dlm = Constants.vbVerticalTab)
    {
        return Strings.Split(lbxPickedStr(lbx, dlm), dlm);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    public static fmSelectorList nuSelector()
    {
        return new fmSelectorList();
    }

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    public static fmSelectorV2 nuSelectorV2()
    {
        return new fmSelectorV2();
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="dc"></param>
    /// <param name="hOhKay"></param>
    /// <param name="mOhKay"></param>
    /// <param name="hCancl"></param>
    /// <param name="mCancl"></param>
    /// <param name="hNoSel"></param>
    /// <param name="mNoSel"></param>
    /// <returns></returns>
    public static fmSelectorList nuSelFromDict(System.Collections.IDictionary dc, string hOhKay = "",
        string mOhKay = "",
        string hCancl = "",
        string mCancl = "", string hNoSel = "", string mNoSel = "")
    {
        return nuSelector().SetHdrOK((string)Interaction.IIf(Strings.Len(hOhKay) > 0, hOhKay, "Confirm Selection"))
            .SetMsgOK(
                (string)Interaction.IIf(Strings.Len(mOhKay) > 0, mOhKay, Strings.Join([
                        "Action will proceed using",
                        "%%%",
                        "(Click CANCEL to quit with no action)"
                    ],
                    Constants.vbCrLf))).SetHdrCancel((string)Interaction.IIf(Strings.Len(hCancl) > 0,
                hCancl,
                "Cancel Operation?")).SetMsgCancel((string)Interaction.IIf(Strings.Len(mCancl) > 0,
                mCancl,
                Strings.Join([
                        "No action will be taken on",
                        "%%%"
                    ],
                    Constants.vbCrLf))).SetHdrNoSelection((string)Interaction.IIf(Strings.Len(hNoSel) > 0,
                hNoSel,
                "No Item Selected!")).SetMsgNoSelection((string)Interaction.IIf(Strings.Len(mNoSel) > 0,
                mNoSel,
                Strings.Join([
                        "Do you wish to cancel the operation?",
                        "(Click NO to return to list)"
                    ],
                    Constants.vbCrLf))).WithList(dc.Keys);
    }
}