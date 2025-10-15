using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class sql0
{
    public static  string sqlTextLocal(string nm)
    {
        return sqlTextInProject(nm, vbProjectLocal());
    }

    public static  string sqlOf_()
    {
        return sqlTextLocal("sqlOf_");
    }

    public static  string sqlOf_simpleSelWhere(string FromView, string GetField, string WhereField, dynamic Matches)
    {
        string mtExpr;

        // generate filter expression
        // based on type of supplied
        // matching value
        if (IsArray(Matches))
        {
            mtExpr = " in ('" + Join(Matches, "', '") + "') ";
            Debugger.Break();
        }
        else if (IsNumeric(Matches))
            mtExpr = " = " + Convert.ToHexString(Matches) + " ";
        else if (VarType(Matches) == Constants.vbString)
            // for String data, single quotes
            // must be "escaped" by repeating
            // each instance, that is, replace
            // each one with two of the same.
            mtExpr = " = '" + Replace(Matches, "'", "''") + "' ";
        else if (IsDate(Matches))
        {
            mtExpr = " = #" + Convert.ToHexString(Format(Matches, "yyyy/mm/dd")) + "# ";
            Debugger.Break();
        }
        else if (IsNull(Matches))
            mtExpr = " is null ";
        else
            Debugger.Break();
        // NOTE: this block MIGHT want
        // exported to its own function

        return " select " + GetField + " from " + FromView + " where " + WhereField + mtExpr + ";";
    }

    public static  string sqlOf_gnsMatlSpec1ops()
    {
        return sqlOf_gnsMatlSpec1ops_v0_1();
    }

    public static  string sqlOf_gnsMatlSpec1ops_v0_1()
    {
        return sqlTextLocal("sqlOf_gnsMatlSpec1ops_v0_1");
    }

    public static  string sqlOf_MatlSpecXref()
    {
        return sqlOf_MatlSpecXref_v0_1();
    }

    public static  string sqlOf_MatlSpecXref_v0_1()
    {
        return sqlTextLocal("sqlOf_MatlSpecXref_v0_1");
    }

    public static  string sqlOf_GnsPartInfo(string Item)
    {
        return Replace(sqlTextLocal("sqlOf_GnsPartInfo"), "%%%", Item);
    }

    public static  string sqlOf_GnsPartMatl(string Item)
    {
        return Replace(sqlTextLocal("sqlOf_GnsPartMatl"), "%%%", Item);
    }

    public static  string sqlOf_GnsMatlOptions(string Matl, dynamic Dims)
    {
        return sqlOf_GnsMatlOptions_v0_2(Matl, Dims);
    }

    public static  string sqlOf_GnsMatlOptions_v0_1(string Matl, double Wdth, double Hght, double Thck = -1, double Lgth = 0)
    {
        // DON'T try to do anything with this yet!
        // see notes on where things are
        return Replace(
            Replace(
                Replace(Replace(Replace(sqlTextLocal("sqlOf_GnsMatlOptions_v0_1"), "$MTL$", Matl), "#THK#", ""),
                    "#WID#", ""), "#HGT#", ""), "#LNG#", ""); // ''
    }

    public static  string sqlOf_GnsMatlOptions_v0_2(string Matl, dynamic Dims)
    {
        while (true)
        {
            // DON'T try to do anything with this yet!
            // see notes on where things are
            if (IsArray(Dims))
                return Replace(Replace(sqlTextLocal("sqlOf_GnsMatlOptions_v0_2"), "%%S6%%", Matl), "%%LS%%",
                    Join(Dims, "), (")); // ''
            if (IsNumeric(Dims))
            {
                Dims = new[] { Dims };
            }
            else
            {
                Debugger.Break(); // because this might be an issue
                // will resort to a sane default for now
                Dims = new[] { 0.075 };
            }
        }
    }

    public static  string sqlOf_GnsTubeHose(double Diam = 0)
    {
        return sqlOf_GnsTubeHose_v0_1(Diam);
    }

    public static  string sqlOf_GnsTubeHose_v0_1(double Diam = 0) // , Matl As String, Dims As dynamic
    {
        // DON'T try to do anything with this yet!
        // see notes on where things are

        var txDiam = Diam > 0
            ? Join(new[] { "between", Convert.ToHexString(Diam - 0.01), "and", Convert.ToHexString(Diam + 0.01) }, " ")
            : "> 0.0";

        return Replace(sqlTextLocal("sqlOf_GnsTubeHose_v0_1"), "%%DI%%", txDiam); // ''
    }

    public static  string sqlOf_ASDF(string Item)
    {
        return Replace(sqlTextLocal("sqlOf_ASDF"), "%%%", Item);
    }

    public static  string sqlOf_03R4LC09_NOCOND()
    {
        return sqlTextLocal("sqlOf_03R4LC09_NOCOND");
    }

    public static  string sqlOf_ERC_PTOSIZE()
    {
        // SQL'''
        // -- ERC-PTOSIZE
        // select I.Item, I.Description1, I.OptionPrice, I.Specification7
        // , D.Item as PartsKit
        // 
        // from vgMfiItems I
        // inner join vgMfiItems D
        // on I.Specification1 = D.Specification1
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
        // SQL'''
        return sqlTextLocal("sqlOf_ERC_PTOSIZE"); // vbTextOfProcInDict
    }

    public static  string sqlOf_test2()
    {
        return sqlTextLocal("sqlOf_test2");
    }
}