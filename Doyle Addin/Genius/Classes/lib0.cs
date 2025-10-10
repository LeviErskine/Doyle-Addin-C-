class lib0
{
    // Measurement Unit Conversion Factors
    public const double cvArSqCm2SqFt = 0.00107639;
    // 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
    // 
    // /  1ft | 1in    \2     2                2
    // ( ------+-------- ) * cm  = 0.00107639 ft
    // \ 12in | 2.54cm /
    public const double cvMassKg2LbM = 2.20462;
    public const double cvLenIn2cm = 2.54;

    /// 
    public const string guidRegPart = "{4D29B490-49B2-11D0-93C3-7E0706000000}";
    public const string guidSheetMetal = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
    public const string guidDesignAccl = "{BB8FE430-83BF-418D-8DF9-9B323D3DB9B9}";
    public const string guidPipingSgmt = "{4D39D5F1-0985-4783-AA5A-FC16C288418C}";
    public const string guidILogicAdIn = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
    /// 
    public const string guidRegAssy = "{E60F81E1-49B3-11D0-93C3-7E0706000000}";
    public const string guidWeldment = "{28EC8354-9024-440F-A8A2-0E0E55D635B0}";

    /// 
    public const string guidPrSetSumm = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"; // Summary Information (Inventor Summary Information)
    public const string guidPrSetDocu = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}"; // Document Summary Information (Inventor Document Summary Information)
    public const string guidPrSetTrkg = "{32853F0F-3444-11D1-9E93-0060B03C1CA6}"; // Design Tracking Properties (Design Tracking Properties)
    public const string guidPrSetUser = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"; // User Defined Properties (Inventor User Defined Properties)
    public const string guidPrSetCLib = "{B9600981-DEE8-4547-8D7C-E525B3A1727A}"; // Content Library Component Properties (Content Library Component Properties)
    public const string guidPrSetCCtr = "{CEAAEE65-91D8-444E-ACBA-BE54A5FB9D4D}"; // ContentCenter (ContentCenter)
                                                                                  // Public Const guidPrSet____  As String = "{00000000-0000-0000-0000-000000000000}" 'Display Name (Name)
    /// 

    public const string gnDesign = "Design Tracking Properties";
    public const string pnMaterial = "Material";          // 
    public const string pnPartNum = "Part Number";       // 
    public const string pnStockNum = "Stock Number";      // 
    public const string pnFamily = "Cost Center";       // 
    public const string pnDesc = "Description";       // 
    public const string pnCatWebLink = "Catalog Web Link";  // 

    public const string gnCustom = "Inventor User Defined Properties";
    public const string pnMass = "GeniusMass";    // 
    public const string pnRawMaterial = "RM";            // 
    public const string pnRmQty = "RMQTY";         // 
    public const string pnRmUnit = "RMUNIT";        // (replaces "RMUOM")
                                                    // '
    public const string pnArea = "Extent_Area";   // 
    public const string pnLength = "Extent_Length"; // 
    public const string pnWidth = "Extent_Width";  // 
                                                   // 
    public const string pnThickness = "Thickness";     // 
                                                       // 

    public VBIDE.VBProject vbProjectLocal()
    {
        vbProjectLocal = ThisDocument.VBAProject.VBProject;
    }

    public ADODB.Connection cnGnsDoyle()
    {
        ADODB.Connection rt;
        /// NOTE[2021.12.08]:
        /// Might consider make rt a Static Object.
        /// If it can be created and opened just once
        /// during a run, this could potentially save
        /// a LOT of overhead from repeated open/close
        /// operations, and might save a little load
        /// on the server, as well.

        rt = new ADODB.Connection();
        {
            var withBlock = rt;
            withBlock.Provider = "SQLOLEDB"; // "SQLNCLI11"
            withBlock.CursorLocation = adUseClient;
            withBlock.Open("Data Source=DOYLE-ERP02", "GeniusReporting", "geniusreporting");
            withBlock.DefaultDatabase = "DoyleDB";
        }
        cnGnsDoyle = rt;
    }

    public Scripting.Dictionary dcIvObjTypeEnum()
    {
        Scripting.Dictionary dc;
        Inventor.ObjectTypeEnum en;

        dc = new Scripting.Dictionary();

        {
            var withBlock = dc;
            en = k3dAViewObject;
            en = kAliasFreeformFeatureObject;
            en = kAliasFreeformFeatureProxyObject;
            en = kAliasFreeformFeaturesObject;
            en = kAnalysisManagerObject;
            en = kAnalyticEdgeWorkAxisDefObject;
            en = kAngleConstraintObject;
            en = kAngleConstraintProxyObject;
            en = kAngleExtentObject;
            en = kAngleiMateDefinitionObject;
        }

        dcIvObjTypeEnum = dc;
    }
    // 
    // =====

    public Scripting.Dictionary dcIvDocTypeEnum()
    {
        Scripting.Dictionary dc;
        Inventor.DocumentTypeEnum en;

        dc = new Scripting.Dictionary();

        {
            var withBlock = dc;
            withBlock.Add(kUnknownDocumentObject, "kUnknownDocumentObject");
            withBlock.Add(kSATFileDocumentObject, "kSATFileDocumentObject");
            withBlock.Add(kPresentationDocumentObject, "kPresentationDocumentObject");
            withBlock.Add(kPartDocumentObject, "kPartDocumentObject");
            withBlock.Add(kNoDocument, "kNoDocument");
            withBlock.Add(kForeignModelDocumentObject, "kForeignModelDocumentObject");
            withBlock.Add(kDrawingDocumentObject, "kDrawingDocumentObject");
            withBlock.Add(kDesignElementDocumentObject, "kDesignElementDocumentObject");
            withBlock.Add(kAssemblyDocumentObject, "kAssemblyDocumentObject");
        }

        dcIvDocTypeEnum = dc;
    }
    // 
    // =====

    public string txDumpLs(Variant ls, string bk = Constants.vbNewLine)
    {
        Variant rt;
        Variant tx;
        long mx;
        long bs;
        long dx;

        if (IsArray(ls))
        {
            bs = LBound(ls);
            mx = UBound(ls);
            if (bs > mx)
                txDumpLs = "";
            else
            {
                ;/* Cannot convert RedimClauseSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.RangeArgumentSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.SimpleArgumentSyntax'.
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.<ConvertArrayBounds>b__20_0(ArgumentSyntax a) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 342
   at System.Linq.Enumerable.SelectEnumerableIterator`2.ToList()
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitRedimClause(RedimClauseSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 140
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
rt(bs To mx)

 */
                for (dx = LBound(ls); dx <= mx; dx++)
                    rt(dx) = txDumpLs(ls(dx));
                txDumpLs = Join(rt, bk);
            }
        }
        else if (IsObject(ls))
        {
            if (ls is Scripting.Dictionary)
                txDumpLs = txDumpLs(ls.Keys);
            else
            {
            }
        }
        else
            txDumpLs = System.Convert.ToHexString(ls);
    }

    public void lsDump(Variant ls, string bk = Constants.vbNewLine)
    {
        Debug.Print(txDumpLs(ls, bk));
    }

    /// The following is copied over from the Excel project file libExt.xlsm

    /// to provide a means of dumping Key-Value pairs from a Dictionary.

    /// 

    public string dumpLsKeyVal(Scripting.Dictionary dc, string dlmField = ",", string dlmLine = Constants.vbNewLine, string nullTxt = "<null>", string emptyTx = "<empty>")
    {
        Scripting.Dictionary d2;
        Variant ky;
        Variant vl;
        // Dim rt As String

        // rt = ""
        if (dc == null)
            dumpLsKeyVal = "";
        else
        {
            d2 = new Scripting.Dictionary();
            {
                var withBlock = dc;
                foreach (var ky in withBlock.Keys)
                {
                    // rt = rt & ky & "," & .Item(ky) & vbNewLine

                    vl = Array(withBlock.Item(ky));
                    // '  Any values which have
                    // '  no direct String conversion
                    // '  are replaced with String defaults
                    // '  (   user supplied, or see
                    // '      Function declaration above )
                    if (IsNull(vl(0)))
                        vl = nullTxt;
                    if (IsEmpty(vl(0)))
                        vl = emptyTx;
                    // If IsMissing(vl) Then vl = ""
                    // If IsError(vl) Then vl = ""
                    if (IsObject(vl(0)))
                    {
                        if (vl(0) == null)
                            vl = "<ob:Nothing>";
                        else
                            vl = "<ob:" + TypeName(vl(0)) + ">";
                    }
                    if (IsArray(vl))
                    {
                        if (IsArray(vl(0)))
                            vl = "<array>";
                        else
                            vl = vl(0);
                    }

                    d2.Add(Join(Array(ky, vl), dlmField), ky);
                }
            }
            dumpLsKeyVal = Join(d2.Keys, dlmLine);
        }
    }
}