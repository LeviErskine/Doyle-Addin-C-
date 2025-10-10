class SurroundingClass
{
    public string sqlTextLocal(string nm)
    {
        sqlTextLocal = sqlTextInProject(nm, vbProjectLocal());
    }

    public string sqlOf_()
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
        sqlOf_ = sqlTextLocal("sqlOf_");
    }

    public string sqlOf_simpleSelWhere(string FromView, string GetField, string WhereField, Variant Matches)
    {
        string mtExpr;

        // generate filter expression
        // based on type of supplied
        // matching value
        if (IsArray(Matches))
        {
            mtExpr = " in ('" + Join(Matches, "', '") + "') ";
            System.Diagnostics.Debugger.Break();
        }
        else if (IsNumeric(Matches))
            mtExpr = " = " + System.Convert.ToHexString(Matches) + " ";
        else if (VarType(Matches) == Constants.vbString)
            /// for String data, single quotes
            /// must be "escaped" by repeating
            /// each instance, that is, replace
            /// each one with two of the same.
            mtExpr = " = '" + Replace(Matches, "'", "''") + "' ";
        else if (IsDate(Matches))
        {
            mtExpr = " = #" + System.Convert.ToHexString(Format(Matches, "yyyy/mm/dd")) + "# ";
            System.Diagnostics.Debugger.Break();
        }
        else if (IsNull(Matches))
            mtExpr = " is null ";
        else
            System.Diagnostics.Debugger.Break();
        // NOTE: this block MIGHT want
        // exported to its own function

        sqlOf_simpleSelWhere = " select " + GetField + " from " + FromView + " where " + WhereField + mtExpr + ";";
    }

    public string sqlOf_gnsMatlSpec1ops()
    {
        sqlOf_gnsMatlSpec1ops = sqlOf_gnsMatlSpec1ops_v0_1();
    }

    public string sqlOf_gnsMatlSpec1ops_v0_1()
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
        sqlOf_gnsMatlSpec1ops_v0_1 = sqlTextLocal("sqlOf_gnsMatlSpec1ops_v0_1");
    }

    public string sqlOf_MatlSpecXref()
    {
        sqlOf_MatlSpecXref = sqlOf_MatlSpecXref_v0_1();
    }

    public string sqlOf_MatlSpecXref_v0_1()
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
        sqlOf_MatlSpecXref_v0_1 = sqlTextLocal("sqlOf_MatlSpecXref_v0_1");
    }

    public string sqlOf_GnsPartInfo(string Item)
    {
        sqlOf_GnsPartInfo = Replace(sqlTextLocal("sqlOf_GnsPartInfo"), "%%%", Item);
    }

    public string sqlOf_GnsPartMatl(string Item)
    {
        sqlOf_GnsPartMatl = Replace(sqlTextLocal("sqlOf_GnsPartMatl"), "%%%", Item);
    }

    public string sqlOf_GnsMatlOptions(string Matl, Variant Dims)
    {
        sqlOf_GnsMatlOptions = sqlOf_GnsMatlOptions_v0_2(Matl, Dims);
    }

    public string sqlOf_GnsMatlOptions_v0_1(string Matl, double Wdth, double Hght, double Thck = -1, double Lgth = 0)
    {
        /// DON'T try to do anything with this yet!
        /// see notes on where things are
        sqlOf_GnsMatlOptions_v0_1 = Replace(Replace(Replace(Replace(Replace(sqlTextLocal("sqlOf_GnsMatlOptions_v0_1"), "$MTL$", Matl), "#THK#", ""), "#WID#", ""), "#HGT#", ""), "#LNG#", "");// ''
    }

    public string sqlOf_GnsMatlOptions_v0_2(string Matl, Variant Dims)
    {
        /// DON'T try to do anything with this yet!
        /// see notes on where things are
        if (IsArray(Dims))
            sqlOf_GnsMatlOptions_v0_2 = Replace(Replace(sqlTextLocal("sqlOf_GnsMatlOptions_v0_2"), "%%S6%%", Matl), "%%LS%%", Join(Dims, "), ("));// ''
        else if (IsNumeric(Dims))
            sqlOf_GnsMatlOptions_v0_2 = sqlOf_GnsMatlOptions_v0_2(Matl, Array(Dims));
        else
        {
            System.Diagnostics.Debugger.Break(); // because this might be an issue
            /// will resort to a sane default for now
            sqlOf_GnsMatlOptions_v0_2 = sqlOf_GnsMatlOptions_v0_2(Matl, Array(0.075)); // should pick up 14GA sheet metal only
        }
    }

    public string sqlOf_GnsTubeHose(double Diam = 0)
    {
        sqlOf_GnsTubeHose = sqlOf_GnsTubeHose_v0_1(Diam);
    }

    public string sqlOf_GnsTubeHose_v0_1(double Diam = 0) // , Matl As String, Dims As Variant
    {
        /// DON'T try to do anything with this yet!
        /// see notes on where things are
        string txDiam;

        if (Diam > 0)
            txDiam = Join(Array("between", System.Convert.ToHexString(Diam - 0.01), "and", System.Convert.ToHexString(Diam + 0.01)), " ");
        else
            txDiam = "> 0.0";

        sqlOf_GnsTubeHose_v0_1 = Replace(sqlTextLocal("sqlOf_GnsTubeHose_v0_1"), "%%DI%%", txDiam); // ''
    }

    public string sqlOf_ASDF(string Item)
    {
        sqlOf_ASDF = Replace(sqlTextLocal("sqlOf_ASDF"), "%%%", Item);
    }

    public string sqlOf_03R4LC09_NOCOND()
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
        sqlOf_03R4LC09_NOCOND = sqlTextLocal("sqlOf_03R4LC09_NOCOND");
    }

    public string sqlOf_ERC_PTOSIZE()
    {
        /// SQL'''
        // -- ERC-PTOSIZE
        // select I.Item, I.Description1, I.OptionPrice, I.Specification7
        // , D.Item as PartsKit
        // 
        // from vgMfiItems I
        // inner join vgMfiItems D
        // on  I.Specification1 = D.Specification1
        // and I.Specification2 = D.Specification2
        // and I.Specification4 = D.Specification4
        // and I.Specification5 = D.Specification5
        // -- and
        // 
        // where I.Specification1 = 'SPREADER'
        // and I.Specification2 ='PTO'
        // and ISNULL(I.Specification3,'') = ''
        // and ISNULL(D.Specification3,'') <> ''
        // and I.Specification4 ='ALL'
        // and I.Specification5 ='DRIVE'
        // 
        // Order by Description1
        // ; --
        /// SQL'''
        sqlOf_ERC_PTOSIZE = sqlTextLocal("sqlOf_ERC_PTOSIZE"); // vbTextOfProcInDict
    }

    public string sqlOf_test2()
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
        sqlOf_test2 = sqlTextLocal("sqlOf_test2");
    }
}