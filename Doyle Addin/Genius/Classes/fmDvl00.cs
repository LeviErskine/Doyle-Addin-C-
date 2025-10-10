class fmDvl00
{
    public Variant fd0g1f1(Variant ls)
    {
        fmEmpty fm;
        MSForms.ListBox lb;
        float mg;
        Variant wk;
        Scripting.Dictionary dc;
        long dx;
        string rt;

        if (IsArray(ls))
            wk = ls;
        else if (IsObject(ls))
        {
            dc = dcOb(obOf(ls));
            if (dc == null)
                wk = Empty;
            else
                wk = dcOb(ls).Keys;
        }
        else
            wk = Empty;

        if (IsEmpty(wk))
            wk = Array("*vvvvvvvvvvv*", "*Unsupported*", "*List Source*", "*^^^^^^^^^^^*");

        fm = nuFmEmpty();
        mg = 10;

        lb = nuMsFmCtListBox(fm, nm: "lbxA");
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
            dx = 0;
            while (!dx == withBlock.ListCount)
            {
                if (withBlock.Selected(dx))
                    rt = rt + Constants.vbVerticalTab + withBlock.List(dx);
                dx = 1 + dx;
            }
        }

        fd0g1f1 = Split(Mid(rt, 2), Constants.vbVerticalTab);
    }

    public fmEmpty nuFmEmpty(Variant f = )
    {
        {
            var withBlock = new fmEmpty();
            /// 
            nuFmEmpty = withBlock.Itself;
        }
    }

    public MSForms.Control obMsFmControl(Variant it)
    {
        object ob;

        ob = obOf(it);
        if (ob is MSForms.Control)
            obMsFmControl = ob;
        else
            obMsFmControl = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public MSForms.ListBox nuMsFmCtListBox(fmEmpty fm, Variant sp = Empty, Variant nm = "", bool vs = true) // MSForms.UserForm
    {
        /// nuMsFmCtListBox -- add new ListBox
        /// Control to supplied fmEmpty Object
        /// 
        /// accepts, but does not yet use,
        /// a specification sp laying out
        /// the parameters defining the
        /// desired control
        /// 
        /// note that fm MUST be fmEmpty
        /// a general MSForms UserForm is
        /// NOT accepted, because it does
        /// NOT support certain essential
        /// properties; for example, there
        /// are not properties to set size
        /// or position, which are essential
        /// to the goals of this system
        /// 
        MSForms.ListBox rt;
        // Dim ct As MSForms.Control

        // fm.Left

        rt = fm.Controls.Add("Forms.ListBox.1", null/* Conversion error: Set to default value for this argument */, vs);
        if (Len(nm) > 0)
            obMsFmControl(rt).Name = nm;

        {
            var withBlock = rt;
        }

        nuMsFmCtListBox = rt;
    }
}