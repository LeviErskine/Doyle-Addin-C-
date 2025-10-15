using System.Collections;
using ADODB;
using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class lib0
{
    // Measurement Unit Conversion Factors
    public const double cvArSqCm2SqFt = 0.00107639;
    // 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
    // 
    // / 1ft | 1in \2 2 2
    // ( ------+-------- ) * cm = 0.00107639 ft
    // \ 12in | 2.54cm /
    public const double cvMassKg2LbM = 2.20462;
    public const double cvLenIn2cm = 2.54;

    // 
    public const string guidRegPart = "{4D29B490-49B2-11D0-93C3-7E0706000000}";
    public const string guidSheetMetal = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
    public const string guidDesignAccl = "{BB8FE430-83BF-418D-8DF9-9B323D3DB9B9}";
    public const string guidPipingSgmt = "{4D39D5F1-0985-4783-AA5A-FC16C288418C}";
    public const string guidILogicAdIn = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
    // 
    public const string guidRegAssy = "{E60F81E1-49B3-11D0-93C3-7E0706000000}";
    public const string guidWeldment = "{28EC8354-9024-440F-A8A2-0E0E55D635B0}";

    // 
    public const string guidPrSetSumm = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"; // Summary Information (Inventor Summary Information)
    public const string guidPrSetDocu = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}"; // Document Summary Information (Inventor Document Summary Information)
    public const string guidPrSetTrkg = "{32853F0F-3444-11D1-9E93-0060B03C1CA6}"; // Design Tracking Properties (Design Tracking Properties)
    public const string guidPrSetUser = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"; // User Defined Properties (Inventor User Defined Properties)
    public const string guidPrSetCLib = "{B9600981-DEE8-4547-8D7C-E525B3A1727A}"; // Content Library Component Properties (Content Library Component Properties)
    public const string guidPrSetCCtr = "{CEAAEE65-91D8-444E-ACBA-BE54A5FB9D4D}"; // ContentCenter (ContentCenter)
    // Public Const guidPrSet____ As String = "{00000000-0000-0000-0000-000000000000}" 'Display Name (Name)
    // 

    public const string gnDesign = "Design Tracking Properties";
    public const string pnMaterial = "Material"; // 
    public const string pnPartNum = "Part Number"; // 
    public const string pnStockNum = "Stock Number"; // 
    public const string pnFamily = "Cost Center"; // 
    public const string pnDesc = "Description"; // 
    public const string pnCatWebLink = "Catalog Web Link"; // 

    public const string gnCustom = "Inventor User Defined Properties";
    public const string pnMass = "GeniusMass"; // 
    public const string pnRawMaterial = "RM"; // 
    public const string pnRmQty = "RMQTY"; // 
    public const string pnRmUnit = "RMUNIT"; // (replaces "RMUOM")
    // '
    public const string pnArea = "Extent_Area"; // 
    public const string pnLength = "Extent_Length"; // 
    public const string pnWidth = "Extent_Width"; // 
    // 
    public const string pnThickness = "Thickness"; // 
    // 

    public static VBIDE.VBProject vbProjectLocal()
    {
        // Not supported in this context; returning null to avoid compile errors.
        return null;
    }

    public static Connection cnGnsDoyle()
    {
        var rt =
            // NOTE[2021.12.08]:
            // Might consider make rt a Static dynamic.
            // If it can be created and opened just once
            // during a run, this could potentially save
            // a LOT of overhead from repeated open/close
            // operations, and might save a little load
            // on the server, as well.
            new Connection();
        {
            rt.Provider = "SQLOLEDB"; // "SQLNCLI11"
            rt.CursorLocation = CursorLocationEnum.adUseClient;
            rt.Open("Data Source=DOYLE-ERP02", "GeniusReporting", "geniusreporting");
            rt.DefaultDatabase = "DoyleDB";
        }
        return rt;
    }

    public static Dictionary dcIvObjTypeEnum()
    {
        var dc = new Dictionary();

        {
            var withBlock = dc;
            var en = kAngleiMateDefinitionObject;
        }

        return dc;
    }

    public static Dictionary dcIvDocTypeEnum()
    {
        DocumentTypeEnum en;

        var dc = new Dictionary();

        {
            dc.Add(kUnknownDocumentObject, "kUnknownDocumentObject");
            dc.Add(kSATFileDocumentObject, "kSATFileDocumentObject");
            dc.Add(kPresentationDocumentObject, "kPresentationDocumentObject");
            dc.Add(kPartDocumentObject, "kPartDocumentObject");
            dc.Add(kNoDocument, "kNoDocument");
            dc.Add(kForeignModelDocumentObject, "kForeignModelDocumentObject");
            dc.Add(kDrawingDocumentObject, "kDrawingDocumentObject");
            dc.Add(kDesignElementDocumentObject, "kDesignElementDocumentObject");
            dc.Add(kAssemblyDocumentObject, "kAssemblyDocumentObject");
        }

        return dc;
    }

    public static string txDumpLs(object ls, string bk = Constants.vbCrLf)
    {
        if (ls == null) return "";
        // If it's a Scripting.Dictionary, dump keys
        if (ls is Dictionary dict)
        {
            var keys = new ArrayList();
            foreach (var k in (IEnumerable)dict.Keys()) keys.Add(k);
            var arr = new string[keys.Count];
            for (var i = 0; i < keys.Count; i++) arr[i] = Convert.ToString(keys[i]);
            return Join(bk, arr);
        }
        // If it's an array, dump each element
        if (ls is not Array a) return Convert.ToString(ls) ?? "";
        var list = new List<string>(a.Length);
        list.AddRange(from object item in a select txDumpLs(item, bk));
        return Join(bk, list);
    }

    public static void lsDump(dynamic ls, string bk = Constants.vbCrLf)
    {
        Debug.Print(txDumpLs(ls, bk));
    }

    // The following is copied over from the Excel project file libExt.xlsm

    // to provide a means of dumping Key-Value pairs from a Dictionary.

    // 

    public static string dumpLsKeyVal(Dictionary dc, string dlmField = ",", string dlmLine = Constants.vbCrLf, string nullTxt = "<null>", string emptyTx = "<null>")
    {
        if (dc == null) return "";
        var lines = new List<string>();
        foreach (var ky in (IEnumerable)dc.Keys())
        {
            object v = dc.get_Item(ky);
            var sv = ToDisplayString(v, nullTxt, emptyTx);
            lines.Add(Convert.ToString(ky) + dlmField + sv);
        }
        return Join(dlmLine, lines);
    }

    private static string ToDisplayString(object v, string nullTxt, string emptyTx)
    {
        while (true)
        {
            switch (v)
            {
                case null:
                    return "<ob:Nothing>";
                case DBNull:
                    return nullTxt;
                case Array { Length: > 0 } arr when arr.GetValue(0) is Array:
                    return "<array>";
                case Array arr:
                {
                    var first = arr.Length > 0 ? arr.GetValue(0) : null;
                    v = first;
                    continue;
                }
            }

            var s = Convert.ToString(v);
            if (s == null) return "";
            return s.Length == 0 ? emptyTx : s;
        }
    }
}