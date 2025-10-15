using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class sql1
{
    public string sqlListValues(Dictionary dc, string sep = "', '", string pfx = "('", string sfx = "')")
    {
        return pfx + Join(dc.Keys, sep) + sfx;
    }

    public string sqlValSelFromDict(Dictionary dc)
    {
        return sqlListValues(dc, "')," + Constants.vbCrLf + Constants.vbTab + "('", "(select pn from (VALUES ('",
            "')" + Constants.vbCrLf + ") as a(pn))");
    }

    public string sqlValsFromDict(Dictionary dc, string lsName = "ls", string fdName = "it")
    {
        return sqlListValues(dc, "')," + Constants.vbCrLf + Constants.vbTab + "('", "(values ('",
            "')" + Constants.vbCrLf + ") as " + lsName + "(" + fdName + ")");
    }

    public string sqlValsFromAssy(Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        var dc = dcRemapByPtNum(dcAiDocComponents(AiDoc));

        var ck = sqlListValues(dc, "')," + Constants.vbCrLf + Constants.vbTab + "('", "(values ('",
            "')" + Constants.vbCrLf + ") as " + lsName + "(" + fdName + ")");
        if (sqlValsFromDict(dc) == ck)
            Debugger.Break();
        return ck;
    }

    public string q1g0x0(Document AiDoc)
    {
        // SQL text function naming convention
        // q1 - "q" for "query", with module number
        // g1 - "g" for "group" (typical usage)
        // x1 - "x" for "text" (stands out better than "t")
        // 
        return "-- SQL text begins here" + Constants.vbCrLf + "" + Constants.vbCrLf + "-- SQL text ends here";
    }

    public string q1g1x1(Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        // SQL text function naming convention
        // q1 - "q" for "query", with module number
        // g1 - "g" for "group" (typical usage)
        // x1 - "x" for "text" (stands out better than "t")
        // 
        return "from vgMfiItems i inner join" + Constants.vbCrLf + sqlValsFromAssy(AiDoc, lsName, fdName) +
               Constants.vbCrLf + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g1x1v2(Document AiDoc, string gnsTbl = "vgMfiItems", string lsName = "ls", string fdName = "it")
    {
        // SQL text function naming convention
        // q1 - "q" for "query", with module number
        // g1 - "g" for "group" (typical usage)
        // x1 - "x" for "text" (stands out better than "t")
        // 

        // REV[2021.08.18] (REVERSED)
        // changed inner join to right (outer) join
        // to pick up Inventor Items not (yet) in Genius
        // REVERSED -- since all returned fields are null,
        // no information is returned for missing Items.
        return "from " + gnsTbl + " i inner join" + Constants.vbCrLf + sqlValsFromAssy(AiDoc, lsName, fdName) +
               Constants.vbCrLf + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g1x1v3(Dictionary dc, string gnsTbl = "vgMfiItems", string lsName = "ls", string fdName = "it")
    {
        // SQL text function naming convention
        // q1 - "q" for "query", with module number
        // g1 - "g" for "group" (typical usage)
        // x1 - "x" for "text" (stands out better than "t")
        // 

        // REV[2021.08.18] (REVERSED)
        // changed inner join to right (outer) join
        // to pick up Inventor Items not (yet) in Genius
        // REVERSED -- since all returned fields are null,
        // no information is returned for missing Items.
        return "from " + gnsTbl + " i inner join" + Constants.vbCrLf + sqlValsFromDict(dc, lsName, fdName) +
               Constants.vbCrLf + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g1x2(Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        return "select" + " i." + IIf(false, "*",
            Join(
                new[]
                {
                    "Item", "Family", "Description1", "Description3", "Unit", "Thickness", "Width", "Length",
                    "Height", "Diameter", "Weight", "Specification1", "Specification2", "Specification3",
                    "Specification4", "Specification5", "Specification6", "Specification7", "Specification8",
                    "Specification9"
                }, ", i.")) + vbCrLf;
    }

    public string q1g1x2v2(Dictionary dc, string gnsTbl = "vgMfiItems", string lsName = "ls", string fdName = "it")
    {
        return "select" + " i." + IIf(false, "*",
            Join(
                new[]
                {
                    "Item", "Family", "Description1", "Description3", "Unit", "Thickness", "Width", "Length",
                    "Height", "Diameter", "Weight", "Specification1", "Specification2", "Specification3",
                    "Specification4", "Specification5", "Specification6", "Specification7", "Specification8",
                    "Specification9"
                }, ", i.")) + vbCrLf;
    }

    public string q1g2x1(Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        // SQL text function naming convention
        // q1 - "q" for "query", with module number
        // g1 - "g" for "group" (typical usage)
        // x1 - "x" for "text" (stands out better than "t")
        // 
        return "from vgIcoBillOfMaterials b inner join" + Constants.vbCrLf + sqlValsFromAssy(AiDoc, lsName, fdName) +
               Constants.vbCrLf + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g2x2(Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        return "select" + " b." + IIf(false, "*", Join(new[]
            {
                "Product", "ItemOrder", "Item", "QuantityInConversionUnit", "ConversionUnit", "ItemType", "Reserved"
            },
            ", b.")) + vbCrLf;
    }

    public string sqlSelAiPurch01fromTextV01(string txList)
    {
        return "-- " + Constants.vbCrLf + "with" + Constants.vbCrLf + Constants.vbTab + "ls as " + txList;
    }

    public string sqlSelAiPurch01fromTextV02(string txList)
    {
        // 
        // 
        dynamic a0;

        long n0 = InStr(1, txList, "'");
        string s0 = Mid(txList, 1 + n0);
        n0 = InStr(1, s0, "'");
        long n1 = InStr(1 + n0, s0, "'"); // - n0 + 1
        if (n1 > 0)
        {
            n1 = n1 - n0 + 1;
            string s1 = Mid(s0, n0, n1);
            a0 = string.Split(s0, s1);
            n0 = UBound(a0);
            a0(n0) = string.Split(a0(n0), "'")(0);
        }
        else
            a0 = new[] { Left(s0, n0 - 1) };
        // s0 = Join(a0, "', '")

        Debug.Print("");

        return "-- " + Constants.vbCrLf + "select" + Constants.vbCrLf + Constants.vbTab + "it.Item, it.Type, it.Family"
               + Constants.vbCrLf + "from" + Constants.vbCrLf + Constants.vbTab + "vgMfiItems as it" + Constants.vbCrLf
               + "where" + Constants.vbCrLf + Constants.vbTab + "it.Item in ('" + Join(a0,
                   "', '") + "')" + Constants.vbCrLf + Constants.vbTab + "and (it.Type = 'R'" + Constants.vbCrLf +
               Constants.vbTab + "or it.Family in ('" + Join(new[] { "D-HDWR", "D-PTO", "D-PTS", "R-PTO", "R-PTS" },
                   "', '") + "'))" + "" + "";
    }

    public string sqlSelAiPdParts01fromTextV01(string txList)
    {
        long n0 = InStr(1, txList, "'");
        string s0 = Mid(txList, 1 + n0);
        n0 = InStr(1, s0, "'");
        var n1 = InStr(1 + n0, s0, "'") - n0 + 1;
        string s1 = Mid(s0, n0, n1);
        var a0 = string.Split(s0, s1);
        n0 = UBound(a0);
        a0(n0) = Split(a0(n0), "'")(0);
        // s0 = Join(a0, "', '")

        Debug.Print("");

        return "-- " + Constants.vbCrLf + "select" + Constants.vbCrLf + Constants.vbTab + Join(new[]
               {
                   "bm.Product", "bm.Item", "pd.Type pType", "pd.Family pFam", "it.Type iType", "it.Family iFam"
               }, ", ") +
               Constants.vbCrLf + "from" + Constants.vbCrLf + Constants.vbTab + "vgIcoBillOfMaterials bm Inner Join" +
               Constants.vbCrLf + Constants.vbTab + "vgMfiItems pd On bm.Product = pd.Item" + Constants.vbCrLf +
               Constants.vbTab + "Left Join vgMfiItems it On bm.Item = it.Item" + Constants.vbCrLf + "where" +
               Constants.vbCrLf +
               Constants.vbTab + "bm.Product in " + txList;
    }

    public string sqlSelTestFromText(string txList)
    {
        return sqlSelAiPdParts01fromTextV01(txList);
    }

    public string sqlSelAiPurch01fromText(string txList)
    {
        return sqlSelAiPurch01fromTextV02(txList);
    }

    public string sqlSelAiPdParts01fromText(string txList)
    {
        return sqlSelAiPdParts01fromTextV01(txList);
    }

    public string sqlSelAiPurch01fromDict(Dictionary dc)
    {
        return sqlSelAiPurch01fromText(Join(string.Split(sqlValSelFromDict(dc), Constants.vbCrLf),
            Constants.vbCrLf + Constants.vbTab));
    }

    public string sqlSelAiPdParts01fromDict(Dictionary dc)
    {
        return sqlSelAiPdParts01fromText(sqlListValues(dc));
    }

    public string sqlSelAiPurch01fromAssy(Document AiDoc)
    {
        return sqlSelAiPurch01fromDict(dcRemapByPtNum(dcAiDocComponents(AiDoc)));
    }

    public string sqlSelAiPdParts01fromAssy(Document AiDoc)
    {
        return sqlSelAiPdParts01fromDict(dcRemapByPtNum(dcAiDocComponents(AiDoc)));
    }
}