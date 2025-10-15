using System.Net.Mime;
using Doyle_Addin.Genius.Forms;
using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class fmDvl00 : Form
{
    public dynamic fd0g1f1(dynamic ls)
    {
        dynamic wk;
        string rt;

        if (IsArray(ls))
            wk = ls;
        else if (IsObject(ls))
        {
            Dictionary dc = dcOb(obOf(ls));
            if (dc == null)
                wk = null;
            else
                wk = dcOb(ls).Keys;
        }
        else
            wk = null;

        if (IsEmpty(wk))
            wk = new[] { "*vvvvvvvvvvv*", "*Unsupported*", "*List Source*", "*^^^^^^^^^^^*" };

        var fm = nuFmEmpty();
        const float mg = 10;

        MSForms.ListBox lb = nuMsFmCtListBox(fm, nm: "lbxA");
        {
            var withBlock = obMsFmControl(lb);
            withBlock.Top = mg;
            withBlock.Left = mg;
            withBlock.Height = fm.InsideHeight - mg - mg;
            withBlock.Width = fm.InsideWidth - mg - mg;
        }

        {
            var withBlock = lb;
            withBlock.MultiSelect = fmMultiSelectMulti; // fmMultiSelectExtended
            withBlock.ListStyle = fmListStyleOption;

            withBlock.List = wk;
            fm.Show(vbModal);

            // Stop
            rt = "";
            long dx = 0;
            while (!dx == withBlock.ListCount)
            {
                if (withBlock.Selected(dx))
                    rt = rt + Constants.vbVerticalTab + withBlock.List(dx);
                dx = 1 + dx;
            }
        }

        return Split(Mid(rt, 2), Constants.vbVerticalTab);
    }

    public fmEmpty nuFmEmpty(dynamic f = )
    {
        {
            var withBlock = new fmEmpty();
            // 
            return withBlock.Itself;
        }
    }

    public Control obMsFmControl(dynamic it)
    {
        var ob = obOf(it);
        return ob as Control;
    }

    public ListBox nuMsFmCtListBox(fmEmpty fm, dynamic sp = null, dynamic nm = "", bool vs = true) // MSForms.UserForm
    {
        // nuMsFmCtListBox -- add new ListBox
        // Control to supplied fmEmpty dynamic
        // 
        // accepts, but does not yet use,
        // a specification sp laying out
        // the parameters defining the
        // desired control
        // 
        // note that fm MUST be fmEmpty
        // a general MSForms UserForm is
        // NOT accepted, because it does
        // NOT support certain essential
        // properties; for example, there
        // are not properties to set size
        // or position, which are essential
        // to the goals of this system
        // 

        ListBox rt =
            // Dim ct As MSForms.Control
            // fm.Left
            fm.Controls.Add("Forms.ListBox.1", null, vs);
        if (Len(nm) > 0)
            obMsFmControl(rt).Name = nm;

        {
            var withBlock = rt;
        }

        return rt;
    }
}