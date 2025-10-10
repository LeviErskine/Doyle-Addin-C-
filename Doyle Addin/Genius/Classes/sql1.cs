class SurroundingClass
{
    public string sqlListValues(Scripting.Dictionary dc, string sep = "', '", string pfx = "('", string sfx = "')")
    {
        /// 
        /// 
        sqlListValues = pfx + Join(dc.Keys, sep) + sfx;
    }

    public string sqlValSelFromDict(Scripting.Dictionary dc)
    {
        /// 
        /// 
        sqlValSelFromDict = sqlListValues(dc, "')," + Constants.vbNewLine + Constants.vbTab + "('", "(select pn from (VALUES ('", "')" + Constants.vbNewLine + ") as a(pn))");
    }

    public string sqlValsFromDict(Scripting.Dictionary dc, string lsName = "ls", string fdName = "it")
    {
        /// 
        /// 
        sqlValsFromDict = sqlListValues(dc, "')," + Constants.vbNewLine + Constants.vbTab + "('", "(values ('", "')" + Constants.vbNewLine + ") as " + lsName + "(" + fdName + ")");
    }

    public string sqlValsFromAssy(Inventor.Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        /// 
        /// 
        Scripting.Dictionary dc;
        string ck;

        dc = dcRemapByPtNum(dcAiDocComponents(AiDoc));

        ck = sqlListValues(dc, "')," + Constants.vbNewLine + Constants.vbTab + "('", "(values ('", "')" + Constants.vbNewLine + ") as " + lsName + "(" + fdName + ")");
        if (sqlValsFromDict(dc) == ck)
            System.Diagnostics.Debugger.Break();
        sqlValsFromAssy = ck;
    }

    public string q1g0x0(Inventor.Document AiDoc)
    {
        /// SQL text function naming convention
        /// q1 - "q" for "query", with module number
        /// g1 - "g" for "group" (typical usage)
        /// x1 - "x" for "text" (stands out better than "t")
        /// 
        q1g0x0 = "-- SQL text begins here" + Constants.vbNewLine + "" + Constants.vbNewLine + "-- SQL text ends here";
    }

    public string q1g1x1(Inventor.Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        /// SQL text function naming convention
        /// q1 - "q" for "query", with module number
        /// g1 - "g" for "group" (typical usage)
        /// x1 - "x" for "text" (stands out better than "t")
        /// 
        q1g1x1 = "from vgMfiItems i inner join" + Constants.vbNewLine + sqlValsFromAssy(AiDoc, lsName, fdName) + Constants.vbNewLine + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g1x1v2(Inventor.Document AiDoc, string gnsTbl = "vgMfiItems", string lsName = "ls", string fdName = "it")
    {
        /// SQL text function naming convention
        /// q1 - "q" for "query", with module number
        /// g1 - "g" for "group" (typical usage)
        /// x1 - "x" for "text" (stands out better than "t")
        /// 

        // REV[2021.08.18] (REVERSED)
        // changed inner join to right (outer) join
        // to pick up Inventor Items not (yet) in Genius
        // REVERSED -- since all returned fields are null,
        // no information is returned for missing Items.
        q1g1x1v2 = "from " + gnsTbl + " i inner join" + Constants.vbNewLine + sqlValsFromAssy(AiDoc, lsName, fdName) + Constants.vbNewLine + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g1x1v3(Scripting.Dictionary dc, string gnsTbl = "vgMfiItems", string lsName = "ls", string fdName = "it")
    {
        /// SQL text function naming convention
        /// q1 - "q" for "query", with module number
        /// g1 - "g" for "group" (typical usage)
        /// x1 - "x" for "text" (stands out better than "t")
        /// 

        // REV[2021.08.18] (REVERSED)
        // changed inner join to right (outer) join
        // to pick up Inventor Items not (yet) in Genius
        // REVERSED -- since all returned fields are null,
        // no information is returned for missing Items.
        q1g1x1v3 = "from " + gnsTbl + " i inner join" + Constants.vbNewLine + sqlValsFromDict(dc, lsName, fdName) + Constants.vbNewLine + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g1x2(Inventor.Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        q1g1x2 = "select" + " i." + IIf(false, "*", Join(Array("Item", "Family", "Description1", "Description3", "Unit", "Thickness", "Width", "Length", "Height", "Diameter", "Weight", "Specification1", "Specification2", "Specification3", "Specification4", "Specification5", "Specification6", "Specification7", "Specification8", "Specification9"), ", i.")) + vbNewLine; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
    }

    public string q1g1x2v2(Scripting.Dictionary dc, string gnsTbl = "vgMfiItems", string lsName = "ls", string fdName = "it")
    {
        q1g1x2v2 = "select" + " i." + IIf(false, "*", Join(Array("Item", "Family", "Description1", "Description3", "Unit", "Thickness", "Width", "Length", "Height", "Diameter", "Weight", "Specification1", "Specification2", "Specification3", "Specification4", "Specification5", "Specification6", "Specification7", "Specification8", "Specification9"), ", i.")) + vbNewLine; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
    }

    public string q1g2x1(Inventor.Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        /// SQL text function naming convention
        /// q1 - "q" for "query", with module number
        /// g1 - "g" for "group" (typical usage)
        /// x1 - "x" for "text" (stands out better than "t")
        /// 
        q1g2x1 = "from vgIcoBillOfMaterials b inner join" + Constants.vbNewLine + sqlValsFromAssy(AiDoc, lsName, fdName) + Constants.vbNewLine + "on i.Item = " + lsName + "." + fdName;
    }

    public string q1g2x2(Inventor.Document AiDoc, string lsName = "ls", string fdName = "it")
    {
        q1g2x2 = "select" + " b." + IIf(false, "*", Join(Array("Product", "ItemOrder", "Item", "QuantityInConversionUnit", "ConversionUnit", "ItemType", "Reserved"), ", b.")) + vbNewLine; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
    }

    public string sqlSelAiPurch01fromTextV01(string txList)
    {
        /// 
        /// 
        sqlSelAiPurch01fromTextV01 = "-- " + Constants.vbNewLine + "with" + Constants.vbNewLine + Constants.vbTab + "ls as " + txList; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
    }

    public string sqlSelAiPurch01fromTextV02(string txList)
    {
        /// 
        /// 
        long n0;
        long n1;
        string s0;
        string s1;
        Variant a0;

        n0 = InStr(1, txList, "'");
        s0 = Mid(txList, 1 + n0);
        n0 = InStr(1, s0, "'");
        n1 = InStr(1 + n0, s0, "'"); // - n0 + 1
        if (n1 > 0)
        {
            n1 = n1 - n0 + 1;
            s1 = Mid(s0, n0, n1);
            a0 = Split(s0, s1);
            n0 = UBound(a0);
            a0(n0) = Split(a0(n0), "'")(0);
        }
        else
            a0 = Array(Left(s0, n0 - 1));
        // s0 = Join(a0, "', '")

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */

        sqlSelAiPurch01fromTextV02 = "-- " + Constants.vbNewLine + "select" + Constants.vbNewLine + Constants.vbTab + "it.Item, it.Type, it.Family" + Constants.vbNewLine + "from" + Constants.vbNewLine + Constants.vbTab + "vgMfiItems as it" + Constants.vbNewLine + "where" + Constants.vbNewLine + Constants.vbTab + "it.Item in ('" + Join(a0, "', '") + "')" + Constants.vbNewLine + Constants.vbTab + "and (it.Type = 'R'" + Constants.vbNewLine + Constants.vbTab + "or it.Family in ('" + Join(Array("D-HDWR", "D-PTO", "D-PTS", "R-PTO", "R-PTS"), "', '") + "'))" + "" + "";
    }

    public string sqlSelAiPdParts01fromTextV01(string txList)
    {
        /// 
        /// 
        long n0;
        long n1;
        string s0;
        string s1;
        Variant a0;

        n0 = InStr(1, txList, "'");
        s0 = Mid(txList, 1 + n0);
        n0 = InStr(1, s0, "'");
        n1 = InStr(1 + n0, s0, "'") - n0 + 1;
        s1 = Mid(s0, n0, n1);
        a0 = Split(s0, s1);
        n0 = UBound(a0);
        a0(n0) = Split(a0(n0), "'")(0);
        // s0 = Join(a0, "', '")

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */

        sqlSelAiPdParts01fromTextV01 = "-- " + Constants.vbNewLine + "select" + Constants.vbNewLine + Constants.vbTab + Join(Array("bm.Product", "bm.Item", "pd.Type pType", "pd.Family pFam", "it.Type iType", "it.Family iFam"), ", ") + Constants.vbNewLine + "from" + Constants.vbNewLine + Constants.vbTab + "vgIcoBillOfMaterials bm Inner Join" + Constants.vbNewLine + Constants.vbTab + "vgMfiItems pd On bm.Product = pd.Item" + Constants.vbNewLine + Constants.vbTab + "Left Join vgMfiItems it On bm.Item = it.Item" + Constants.vbNewLine + "where" + Constants.vbNewLine + Constants.vbTab + "bm.Product in " + txList; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
    }

    public string sqlSelTestFromText(string txList)
    {
        /// 
        /// 

        sqlSelTestFromText = sqlSelAiPdParts01fromTextV01(txList);
    }

    public string sqlSelAiPurch01fromText(string txList)
    {
        /// 
        /// 

        sqlSelAiPurch01fromText = sqlSelAiPurch01fromTextV02(txList);
    }

    public string sqlSelAiPdParts01fromText(string txList)
    {
        /// 
        /// 

        sqlSelAiPdParts01fromText = sqlSelAiPdParts01fromTextV01(txList);
    }

    public string sqlSelAiPurch01fromDict(Scripting.Dictionary dc)
    {
        /// 
        /// 
        sqlSelAiPurch01fromDict = sqlSelAiPurch01fromText(Join(Split(sqlValSelFromDict(dc), Constants.vbNewLine), Constants.vbNewLine + Constants.vbTab));
    }

    public string sqlSelAiPdParts01fromDict(Scripting.Dictionary dc)
    {
        /// 
        /// 
        sqlSelAiPdParts01fromDict = sqlSelAiPdParts01fromText(sqlListValues(dc));
    }

    public string sqlSelAiPurch01fromAssy(Inventor.Document AiDoc)
    {
        /// 
        /// 
        sqlSelAiPurch01fromAssy = sqlSelAiPurch01fromDict(dcRemapByPtNum(dcAiDocComponents(AiDoc)));
    }

    public string sqlSelAiPdParts01fromAssy(Inventor.Document AiDoc)
    {
        /// 
        /// 
        sqlSelAiPdParts01fromAssy = sqlSelAiPdParts01fromDict(dcRemapByPtNum(dcAiDocComponents(AiDoc)));
    }
}