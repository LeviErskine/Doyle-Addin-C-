using Doyle_Addin.Genius.Forms;
using Microsoft.VisualBasic;
using Point = Inventor.Point;
using static Doyle_Addin.Genius.Classes.lib0;
using static Inventor.ViewOrientationTypeEnum;

namespace Doyle_Addin.Genius.Classes;

/// <summary>
/// 
/// </summary>
public class dvl0
{
    private Dictionary d0g0f0(SheetMetalComponentDefinition cd, Dictionary dc = null)
    {
        while (true)
        {
            // New Sheet Metal Part processing function
            // Reference function d0g0f0
            // 
            // '
            // '
            // Dim op As Boolean
            // '
            // '
            // '

            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            var rt = dc;
            if (cd == null)
            {
            }
            else
            {
                var pt = (PartDocument)cd.Document;
                var ps = pt.PropertySets.GetType(gnCustom);
                var v1 = d0g1f0(pt);
                // op = pt.Open

                Information.Err().Clear();
                var prThk = cd.Thickness;
                long ec = Information.Err().Number;
                var ed = Information.Err().Description;

                if (ec != 0) return rt;
                if (!cd.HasFlatPattern)
                {
                    MsgBoxResult ck = newFmTest2().AskAbout(pt,
                        null /* Conversion error: Set to default value for this argument */,
                        "NO FLAT PATTERN!" + Constants.vbCrLf + "Try to generate one?");
                    if (ck == Constants.vbYes)
                    {
                        // If Not op Then Stop
                        // Want to see if forcing an Unfold
                        // causes an unopened document to open.
                        // Also want to see how Open Property
                        // relates to a referenced Document
                        // not (yet) separately opened.

                        Information.Err().Clear();
                        cd.Unfold();
                        if (Information.Err().Number == 0)
                        {
                            if (cd.HasFlatPattern) cd.FlatPattern.ExitEdit();
                        }
                        else
                            Debugger.Break(); // Couldn't make Flat Pattern

                        Information.Err().Clear();

                        var v2 = d0g1f0(pt);
                        if (v2 == null)
                        {
                            if (v1 == null) v2.Close();
                        }
                    }
                }

                if (cd.HasFlatPattern)
                {
                    {
                        var withBlock = cd.FlatPattern;
                        // First, make sure it's VALID
                        double dLength;
                        double dWidth;
                        double dHeight;
                        double dArea;
                        double dfHtThk;
                        {
                            var withBlock1 = withBlock.Body.RangeBox;
                            // Check height against thickness
                            // Valid flat pattern should return
                            // zero or VERY minimal difference
                            dHeight = (withBlock1.MaxPoint.Z - withBlock1.MinPoint.Z);
                            dfHtThk = double.Abs(dHeight - prThk.Value);

                            // the extent of the face.
                            // Extract the width, length and area from the range.
                            dLength = (withBlock1.MaxPoint.X - withBlock1.MinPoint.X);
                            dWidth = (withBlock1.MaxPoint.Y - withBlock1.MinPoint.Y);
                            dArea = dLength * dWidth;
                        }
                        // Stop
                        // At this point, we should have enough
                        // to check at least a few things,
                        // and possibly pick out stock.
                        // 
                        if (dfHtThk > 0.01)
                        {
                            // Stop 'and prep for machined (non sheet metal) specs
                            // Pretty sure dimension values
                            // come through in centimeters
                            // so try converting them here
                            // sort3dimsUp
                            d0g1f4(cd);
                            rt = d0g1f3(pt,
                                sort3dimsUp(dHeight / cvLenIn2cm, dWidth / cvLenIn2cm, dLength / cvLenIn2cm), rt);
                        }

                        if (dArea > 0)
                        {
                            // an invalid flat pattern SHOULD have no geometry,
                            // which means it SHOULD have no area to speak of.
                            // '
                            // One would think this obvious, in retrospect,
                            // but one would not be surprised to be proven wrong.
                            // Again.

                            {
                                // Convert values into document units.
                                // This will result in strings that are identical
                                // to the strings shown in the Extent dialog.
                                {
                                    var withBlock2 = pt.UnitsOfMeasure;
                                    var strWidth = withBlock2.GetStringFromValue(dWidth,
                                        withBlock2.GetStringFromType(withBlock2.LengthUnits));
                                    var strLength = withBlock2.GetStringFromValue(dLength,
                                        withBlock2.GetStringFromType(withBlock2.LengthUnits));
                                    var strArea = withBlock2.GetStringFromValue(dArea,
                                        withBlock2.GetStringFromType(withBlock2.LengthUnits) + "^2");

                                    var strDVNs = dfHtThk > 0.01
                                        ? withBlock2.GetStringFromValue(dfHtThk,
                                            withBlock2.GetStringFromType(withBlock2.LengthUnits))
                                        : "";
                                }
                            }
                        }
                        else
                        {
                            if (MessageBox.Show(
                                    Join(
                                        new[]
                                        {
                                            "The flat pattern for this", "part has no features,",
                                            "and is likely not valid.", "", "Pause here to review?",
                                            "(Click 'NO' to just keep going)"
                                        }, Constants.vbCrLf), Constants.vbYesNo, "Invalid Flat Pattern") ==
                                Constants.vbYes) Debugger.Break(); // and let the user look into it
                            Debug.Print(aiDocument(withBlock.Document).FullDocumentName);
                        }
                    }

                    // Add area to custom property set
                    // rt = dcWithProp(aiPropSet, pnRmQty, dArea * cvArSqCm2SqFt, rt)

                    // Add Width to custom property set
                    // rt = dcWithProp(aiPropSet, pnWidth, strWidth, rt)

                    // Add Length to custom property set
                    // rt = dcWithProp(aiPropSet, pnLength, strLength, rt)

                    // Add AreaDescription to custom property set
                    // rt = dcWithProp(aiPropSet, pnArea, strArea, rt)

                    if (Strings.Len(strDVNs) > 0)
                    {
                    }
                }
            }

            else
            Debugger.Break();

            return rt;

            break;
        }
    }
    // For Each dc In aiDocAssy(aiDocActive).ComponentDefinition.Occurrences: Debug.Print aiDocument(aiCompOcc(obOf(dc)).Definition.Document).Open, aiDocument(aiCompOcc(obOf(dc)).Definition.Document).FullDocumentName: Next
    // Looks like Open property will NOT distinguish documents in tab list from those not
    // All entries came up True

    private static  Inventor.View d0g1f0(Document rf)
    {
        Inventor.View rt = null;
        foreach (Inventor.View vw in ThisApplication.Views.Cast<dynamic>().Where(vw => vw.Document == rf))
        {
            rt = vw;
        }

        return rt;
    }

    public static  SheetMetalComponentDefinition d0g1f1(PartDocument rf)
    {
        return rf == null ? null : aiCompDefShtMetal(rf.ComponentDefinition);
    }

    public static  dynamic noVal(VbVarType vt = )
    {
        if (vt & Constants.vbArray)
            return Array.Empty<dynamic>();
        switch (vt)
        {
            case dynamic _ when Constants.vbString:
            {
                return "";
                break;
            }

            case dynamic _ when Constants.vbLong:
            {
                return Convert.ToInt64(0);
                break;
            }

            case dynamic _ when Constants.vbVariant:
            {
                return null;
                break;
            }

            case dynamic _ when Constants.vbInteger:
            {
                return Convert.ToInt32(0);
                break;
            }

            case dynamic _ when Constants.vbSingle:
            {
                return Convert.ToSingle(0);
                break;
            }

            case dynamic _ when Constants.vbDouble:
            {
                return Convert.ToDouble(0);
                break;
            }

            case dynamic _ when Constants.vbDecimal:
            {
                return Convert.ToDecimal(0);
                break;
            }

            case dynamic _ when Constants.vbCurrency:
            {
                return CCur(0);
                break;
            }

            case dynamic _ when Constants.vbBoolean:
            {
                return Convert.ToBoolean(0);
                break;
            }

            case dynamic _ when Constants.vbByte:
            {
                return Convert.ToByte(0);
                break;
            }

            case Constants.vbEmpty:
            {
                return null;
                break;
            }

            case dynamic _ when Constants.vbNull:
            {
                return Null;
                break;
            }

            case dynamic _ when Constants.vbObject:
            {
                return null;
                break;
            }

            case dynamic _ when Constants.vbDate:

            case dynamic _ when vbError:

            case dynamic _ when vbDataObject:

            case dynamic _ when Constants.vbUserDefinedType:
            {
                Debugger.Break(); // noVal = null
                break;
            }
        }
    }

    private static  double[] pt3d(double d0 = 0, double d1 = 0, double d2 = 0)
    {
        var rt = new double[3];

        rt[0] = d0;
        rt[1] = d1;
        rt[2] = d2;

        return rt;
    }

    private static double[] sort3dimsUp(double d0, double d1, double d2)
    {
        double[] rt;

        if (d1 < d0)
            rt = sort3dimsUp(d1, d0, d2);
        else if (d2 < d1)
            rt = sort3dimsUp(d2, d0, d1);
        else
        {
            rt = new double[3];
            rt[0] = d0;
            rt[1] = d1;
            rt[2] = d2;
        }

        return rt;
    }

    public static  double[] sort3dimsDn(double d0, double d1, double d2)
    {
        double[] rt;

        if (d1 > d0)
            rt = sort3dimsDn(d1, d0, d2);
        else if (d2 > d1)
            rt = sort3dimsDn(d2, d0, d1);
        else
        {
            rt = new double[3];
            rt[0] = d0;
            rt[1] = d1;
            rt[2] = d2;
        }

        return rt;
    }

    public static  double[] aiBoxDims(Box RefBox)
    {
        var rt = Array.Empty<double>();
        Point mx;
        Point mn;

        {
            mx = RefBox.MaxPoint;
            mn = RefBox.MinPoint;
        }

        rt[0] = mx.X - mn.X;
        rt[1] = mx.Y - mn.Y;
        rt[2] = mx.Z - mn.Z;

        return rt;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="RefBox"></param>
    /// <returns></returns>
    public static Box aiBoxSortDown(Box RefBox)
    {
        Box rt;

        {
            var withBlock = ThisApplication.TransientGeometry;
            rt = withBlock.CreateBox();
            {
                var mx = RefBox.MaxPoint;
                var mn = RefBox.MinPoint;
                rt.PutBoxData(pt3d(), sort3dimsDn(mx.X - mn.X, mx.Y - mn.Y, mx.Z - mn.Z));
            }
        }

        return rt;
    }

    public static  Dictionary dcSteelType2Spec6()
    {
        var rt = new Dictionary();
        {
            rt.Add("Steel, Mild", "MS");
            rt.Add("Stainless Steel", "SS");
            rt.Add("Stainless Steel, Austenitic", "SS");
        }

        return rt;
    }

    public static  string steelSpec6(string stl, long Ask = 0)
    {
        string rt;
        // With dcSteelType2Spec6()
        // If .Exists(stl) Then
        // steelSpec6 = .get_Item(stl)
        // Else
        // steelSpec6 = ""
        // End If
        // End With

        switch (stl)
        {
            case "Stainless Steel":
            case "Stainless Steel, Austenitic":
            case "Stainless Steel 304":
            {
                rt = "SS";
                break;
            }

            case "Steel, Mild":
            {
                rt = "MS";
                break;
            }
            case "Rubber":
            case "Rubber, Silicone":
            case "UHMW, White":
            {
                rt = ""; // LG
                break;
            }

            default:
            {
                if (Ask)
                {
                    Debug.Print("=== UNKNOWN MATERIAL ===");
                    Debug.Print(" (" + stl + ")");
                    Debug.Print("Please supply a code for Specification 6,");
                    Debug.Print("if applicable, on the line below, and");
                    Debug.Print("press [ENTER] or [RETURN] to modify.");
                    Debug.Print("Press [F5] when ready to continue.");
                    Debug.Print("rt = \"\" '<-( place code between double quotes )");
                    Debugger.Break();
                }
                else
                    rt = "";

                break;
            }
        }

        return rt;
    }

    public static  Dictionary d0g1f3(PartDocument rf, double[] dm, Dictionary dc = )
    {
        Dictionary rt;
        // '
        // '

        if (dc == null)
            rt = d0g1f3(rf, dm, new Dictionary());
        else
        {
            rt = dc;

            var aspect = dm[1] / dm[0];
            var offSqr = aspect - 1;
            var length = dm[2];
            string rmType = rf.PropertySets(gnDesign).get_Item(pnMaterial).Value;
            var rmSpc6 = steelSpec6(rmType);

            // Debug.Print ".Add """ & rmType & """, """""
            Debug.Print("Material: " + rmType + " (" + rmSpc6 + ")");
            // 
            Debug.Print("Cross Section: " + Format(dm[1], "0.000") + " X " + Format(dm[0], "0.000"));
            Debug.Print("Length: " + Format(dm[2], "0.000"));
            switch (offSqr)
            {
                case < 0.01:
                    Debug.Print("Likely Square or Round");
                    break;
                case > 20:
                    Debug.Print("Likely Sheet or Plate");
                    break;
                default:
                    Debug.Print("Likely Rectangular, Uneven?");
                    break;
            }

            var rmItem = rmType + "???";
            var rmQty = Format(dm[2], "0.000");
            const string rmUnit = "IN";

            dc.Add("RM", rmItem);
            dc.Add("RMQTY", rmQty);
            dc.Add("RMUNIT", rmUnit);
            Debugger.Break();
        }

        return rt;
    }

    private static  Dictionary d0g1f4(SheetMetalComponentDefinition cd)
    {
        byte[] kyBytes;
        var d0 = new Dictionary();
        var d1 = new Dictionary();

        var rkMgr = aiDocument(cd.Document).ReferenceKeyManager;

        var kyContx = rkMgr.CreateKeyContext;

        foreach (SurfaceBody sb in cd.SurfaceBodies)
        {
            foreach (var fc in sb.Faces)
            {
                fc.GetHashCode(kyBytes, kyContx);
                var kyLabel = rkMgr.KeyToString(kyBytes);
                d1.Add(kyLabel, fc);
                {
                    var withBlock = fc.Evaluator;
                    Debug.Print(Information.TypeName(fc.Geometry) + "(" +
                                Convert.ToHexString(0 + IIf(withBlock.IsExtrudedShape, 1, 0) +
                                                    IIf(withBlock.IsRevolvedShape, 2, 0)) + ")");
                    {
                        var withBlock1 = withBlock.RangeBox;
                        d0.Add(d0.Count, withBlock1.MaxPoint);
                        d0.Add(d0.Count, withBlock1.MinPoint);
                    }
                }
            }

            foreach (dynamic ky in d0.Keys)
            {
                Point pt = d0.get_Item(ky);
                foreach (dynamic k1 in d1.Keys) // fc In sb.Faces
                {
                    Face fc = d1.get_Item(k1);
                    // If d1.Exists(k1) Then
                    if (!fc.Evaluator.RangeBox.Contains(pt))
                        d1.Remove(k1);
                }

                Debugger.Break();
            }
        }
    }

    public static  Dictionary d0g1f5(SurfaceBody sb)
    {
        var dFc = new Dictionary();
        var dPt = new Dictionary();

        foreach (Face fc in sb.Faces)
        {
            dFc.Add(fc.InternalName, fc);
            {
                var withBlock = fc.Evaluator;
                Debug.Print(TypeName(fc.Geometry) + "(" +
                            Convert.ToHexString(0 + IIf(withBlock.IsExtrudedShape, 1, 0) +
                                                IIf(withBlock.IsRevolvedShape, 2, 0)) + ")");
                {
                    var withBlock1 = withBlock.RangeBox;
                    dPt.Add(dPt.Count, withBlock1.MaxPoint);
                    dPt.Add(dPt.Count, withBlock1.MinPoint);
                }
            }
        }

        foreach (dynamic kPt in dPt.Keys)
        {
            Point pt = dPt.get_Item(kPt);
            foreach (var fc in sb.Faces)
            {
                if (!dFc.Exists(fc.InternalName)) continue;
                if (fc.Evaluator.RangeBox.Contains(pt)) continue;
                dFc.Remove(fc.InternalName);
                if (dFc.Count == 0)
                    Debugger.Break();
            }

            Debugger.Break();
        }

        return dFc;
    }

    public static  string d0g1f6(Face fc)
    {
        byte[] kyBytes;
        string rt;

        {
            var withBlock =
                    aiDocument(fc.SurfaceBody.ComponentDefinition.Document).ReferenceKeyManager // .CreateKeyContext
                ;
            var kyContx = withBlock.CreateKeyContext;
            fc.GetReferenceKey(kyBytes, kyContx);
            rt = withBlock.KeyToString(kyBytes);
        }
        return rt;
    }

    public static  Point aiPoint(dynamic ob)
    {
        return ob as Point;
    }

    // d0g2: Testing

    // 

    // 

    public static  void d0g2f1()
    {
        // Verify 3-way sorting function sort3dimsUp
        var ck = sort3dimsUp(2, 3, 5);
        Debugger.Break();
        ck = sort3dimsUp(2, 5, 3);
        Debugger.Break();
        ck = sort3dimsUp(3, 2, 5);
        Debugger.Break();
        ck = sort3dimsUp(3, 5, 2);
        Debugger.Break();
        ck = sort3dimsUp(5, 2, 3);
        Debugger.Break();
        ck = sort3dimsUp(5, 3, 2);
        Debugger.Break();
    }

    public static  void d0g2f2()
    {
        // Testing a new spec pickup system

        {
            var withBlock = dcAiDocComponents(ThisApplication.ActiveDocument);
            foreach (string ky in withBlock.Keys)
            {
                Debug.Print(ky);
                withBlock.get_Item(ky) = d0g0f0(aiCompDefShtMetal(aiCompDefOf(aiDocPart(withBlock.get_Item(ky)))));
                if (withBlock.get_Item(ky) == null)
                    withBlock.Remove(ky);
            }
        }
    }

    public static  void d0g2f3()
    {
        // Checking some behaviors
        // on string arrays
        // vs variants
        {
            var withBlock = new aiPropSetter();
            Debug.Print(Join(withBlock.PropList(), "|"));
            foreach (var ky in withBlock.PropList())
                Debug.Print(ky);
        }
    }

    public static  Dictionary d0g2f4(Dictionary dc)
    {
        // Return Dictionary of ALLOCATED Property
        // Values (True/False) attached to all components
        // and subcomponents of the active Document.
        // 
        // Where the ALLOCATED Property is not present,
        // represent it as "<default>"
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                // Debug.Print ky
                {
                    var withBlock1 = aiDocument(dc.get_Item(ky)).PropertySets.get_Item(gnCustom);
                    Information.Err().Clear();
                    Property pr = withBlock1.get_Item("ALLOCATED");
                    if (Information.Err().Number == 0)
                        rt.Add(ky, Convert.ToHexString(pr.Value) + "|" + ky);
                    else
                        rt.Add(ky, "<default>" + "|" + ky);
                }
            }
        }
        return rt;
    }
    // Debug.Print Join(d0g2f4(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbCrLf)

    public static  Dictionary d0g2f6(Dictionary dc, string pn, string gn = gnCustom, string df = "<NOPROP>")
    {
        // Return Dictionary of named Property Values
        // attached to all Inventor Documents
        // in supplied Dictionary.
        // 
        // 
        // Where the ALLOCATED Property is not present,
        // represent it as "<default>"
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document ad = aiDocument(dc.get_Item(ky));
                if (ad == null)
                {
                }

                {
                    var withBlock1 = ad.PropertySets.get_Item(gn);
                    Information.Err().Clear();
                    var pr = withBlock1.get_Item("ALLOCATED");
                    if (Information.Err().Number == 0)
                        rt.Add(ky, Convert.ToHexString(pr.Value) + "|" + ky);
                    else
                        rt.Add(ky, "<default>" + "|" + ky);
                }
            }
        }
        return rt;
    }
    // Debug.Print Join(d0g2f6(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbCrLf)

    public static  Dictionary d0g2f5(Dictionary dc)
    {
        // Attempt to "transpose" contents of Dictionary
        // and return a dictionary of Items mapped
        // to sub-Dictionaries containing all keys
        // which mapped to each value
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                var kv = d0g2f5a(dc.get_Item(ky));
                {
                    if (!rt.Exists(kv))
                        rt.Add(kv, new Dictionary());

                    {
                        var withBlock2 = dcOb(rt.get_Item(kv));
                        withBlock2.Add.Count(null /* Conversion error: Set to default value for this argument */, ky);
                    }
                }
            }
        }
    }

    private static  dynamic d0g2f5a(dynamic vr)
    {
        // Return any dynamic that is NOT an dynamic
        // dynamic handling MAY be addressed later.
        // 
        if (vr != null)
            Debugger.Break();
        else
            return null;
    }

    public static  Dictionary dcMaterialUsage(Dictionary dc)
    {
        var rt = new Dictionary();
        foreach (PartDocument pt in from dynamic ky in dc select aiDocPart(obOf(dc.get_Item(ky))))
        {
            if (pt == null)
            {
            }
            else
            {
                string pn = pt.PropertySets(gnDesign).get_Item(pnPartNum).Value;
                string mt = pt.PropertySets(gnDesign).get_Item(pnMaterial).Value;
                {
                    if (rt.Exists(mt))
                        rt.get_Item(mt) = rt.get_Item(mt) + Constants.vbCrLf + Constants.vbTab + pn;
                    else
                        rt.Add(mt, mt + Constants.vbCrLf + Constants.vbTab + pn);
                }
                pt = null;
            }
        }

        return rt;
    }
    // lsDump dcMaterialUsage(dcAiDocsOfType(kPartDocumentObject, dcAiDocComponents(aiDocActive()))).Items

    public static  Dictionary d0g3f0()
    {
        var rt = new Dictionary();
        {
            var withBlock = ThisApplication.ActiveMaterialLibrary;
            foreach (Asset mt in withBlock.MaterialAssets)
                rt.Add(mt.DisplayName, mt);
        }
        return rt;
    }
    // lsDump d0g3f0().Keys

    public static  Dictionary dcGrpByPtNum(Dictionary dc)
    {
        // '
        // ' Returns Dictionary of Dictionaries
        // ' grouping Inventor Documents in
        // ' supplied Dictionary by their Part
        // ' Numbers.
        // '
        // ' Ideally, each Document's Part Number
        // ' should be unique, and each sub Dictionary
        // ' should contain only one Document, however,
        // ' it is possible for more than one Document
        // ' to have the same Part Number.
        // '
        // ' By returning a Dictionary of Dictionaries,
        // ' this function provides a way for the client
        // ' to detect and respond to any conflicts.
        // '

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document pt = aiDocument(dc.get_Item(ky));
                var dn = pt.FullDocumentName;
                var pn = Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));
                {
                    if (!rt.Exists(pn))
                        rt.Add(pn, new Dictionary());

                    {
                        var withBlock2 = dcOb(rt.get_Item(pn));
                        if (withBlock2.Exists(dn))
                            Debugger.Break(); // because something went wrong
                        else
                            withBlock2.Add(dn, pt);
                    }
                }
                Debug.Print("");
            }
        }

        return rt;
    }

    /// <summary>
    ///' Returns Dictionary of Inventor
    ///' Documents keyed on Part Number
    /// </summary>
    /// <param name="dc"></param>
    /// <returns></returns>
    public static  Dictionary dcRemapByPtNum(Dictionary dc)
    {


        var rt = new Dictionary();
        var xt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document pt = aiDocument(dc.get_Item(ky));
                var pn = Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));
                // 

                // 
                // If pt.DocumentType = kPartDocumentObject Then
                // ''' UPDATE[2021.06.22]
                // ''' moving hardware component check outside of
                // ''' and before collision check in order to skip
                // ''' known hardware items entirely.
                // ''' '
                // ''' UPDATE[2021.06.21]
                // ''' implementing a set of checks for hardware components
                // ''' probably want to move outside Dictionary check
                // ''' to catch hardware elements before they're added.
                // ''' Could lead to trouble if this prevents new hardware
                // ''' Items from being added to Genius, but don't believe
                // ''' this is a high risk.
                // If aiDocPart(pt).ComponentDefinition.IsContentMember Then
                // 'it's commodity hardware
                // With cnGnsDoyle().Execute(' "select Family from vgMfiItems where Item = '"' & pn & "';"' )
                // If .BOF And .EOF Then
                // 'probably not in Genius
                // 'keep it -- may need added
                // ElseIf Split(.GetString(' adClipString, , "", vbVerticalTab' ), vbVerticalTab)(0) = "D-HDWR" Then
                // pt = Nothing 'and move on
                // Else
                // Debug.Print ; 'Breakpoint Landing
                // 'Stop
                // End If
                // End If
                // End With
                // ElseIf pt.PropertySets.get_Item(gnDesign).get_Item(pnFamily).Value = "D-HDWR" Then
                // 'it's in commodity hardware family
                // pt = Nothing 'and move on
                // ElseIf InStr(1, "|D-HDWR|D-PTS|R-PTS|", "|" & pt.PropertySets.get_Item(gnDesign).get_Item(pnFamily).Value & "|") > 0 Then
                // 'it's PROBABLY hardware
                // 'but keep it, just in case
                // Debug.Print ; 'Breakpoint Landing
                // Else 'nothing special to worry about, probably
                // Debug.Print ; 'Breakpoint Landing
                // 'Stop
                // End If
                // Else 'we've got an Assembly
                // Debug.Print ; 'Breakpoint Landing
                // 'Stop
                // End If
                // 

                // 
                if (Strings.Len(pn) > 0)
                {
                    {
                        if (rt.Exists(pn))
                        {
                            {
                                if (!xt.Exists(pn))
                                    xt.Add(pn, new Dictionary());

                                {
                                    var withBlock3 = dcOb(xt.get_Item(pn));
                                    withBlock3.Add(pt.FullDocumentName, pt);
                                }
                            }
                        }
                        else
                            rt.Add(pn, pt);
                    }
                    Debug.Print("");
                }
                else
                {
                    Debug.Print(InputBox(
                        "This component has no part number:" + Constants.vbCrLf + pt.DisplayName + Constants.vbCrLf +
                        Convert.ToHexString(aiDocPropVal(pt, pnDesc, gnDesign)) + Constants.vbCrLf + Constants.vbCrLf +
                        "Copy file path from text box for later review.", pt.DisplayName, pt.FullDocumentName));
                    if (getFromClipBdWin10() == pt.FullDocumentName)
                    {
                    }
                    else if (MessageBox.Show(
                                 "Are you sure you want to continue" + Constants.vbCrLf +
                                 "without recording this file path?", Constants.vbExclamation + Constants.vbYesNo,
                                 "File Path not copied!") == Constants.vbNo)
                        Debugger.Break();
                }
            }
        }

        if (xt.Count > 0)
            Debug.Print(MessageBox.Show(
                Join(
                    new[]
                    {
                        "The following Part Numbers are", "assigned to more than one Model:", "",
                        Constants.vbTab + Join(xt.Keys, Constants.vbCrLf + Constants.vbTab), ""
                    }, Constants.vbCrLf), Constants.vbOKOnly | Constants.vbInformation, "Duplicate Part Numbers!"));

        return rt;
    }

    public static  Dictionary dcRemapByFilePath(Dictionary dc)
    {
        // ' Returns Dictionary of Inventor
        // ' Documents re-keyed to File Path.
        // ' Typically for a Dictionary
        // ' previously remapped to another
        // ' key (most likely Part Number)

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document pt = aiDocument(dc.get_Item(ky));
                var pn = pt.FullDocumentName;
                rt.Add(pn, pt);
            }
        }
        return rt;
    }

    public static  Dictionary dcRemapByPtNumFilePath(Dictionary dc)
    {
        // ' Returns Dictionary of Inventor
        // ' Documents keyed on Part Number
        // ' combined with original key
        // ' (which SHOULD be full doc path)

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document pt = aiDocument(dc.get_Item(ky));
                var pn = Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign)) + Constants.vbTab +
                         pt.FullDocumentName; // ky
                rt.Add(pn, pt);
            }
        }
        return rt;
    }

    public static  Dictionary dcGeniusItems()
    {
        // ' Generates Dictionary of Items in Genius
        // ' Formerly d0g3f2

        var rt = new Dictionary();
        {
            var withBlock = cnGnsDoyle();
            {
                var withBlock1 = withBlock.Execute("select Item, ItemID from vgMfiItems");
                ADODB.Field ky = withBlock1.Fields("Item");
                ADODB.Field vl = withBlock1.Fields("ItemID");
                while (!withBlock1.EOF | withBlock1.BOF)
                {
                    rt.Add(ky.Value, vl.Value);
                    withBlock1.MoveNext();
                }
            }
        }
        return rt;
    }

    public static  Dictionary[] d0g3f3(Dictionary dc)
    {
        // ' Appears intended to separate
        // ' parts in Genius from those
        // ' not there yet. I don't think
        // ' this one was working quite
        // ' right yet. New kyPick system
        // ' should handle this properly,
        // ' now, in any case.
        var rt = new Dictionary[2];

        rt[0] = new Dictionary();
        rt[1] = new Dictionary();
        {
            var withBlock = dcGeniusItems();
            foreach (var ky in dc)
            {
                if (withBlock.Exists(ky))
                    rt[1].Add(ky, withBlock.get_Item(ky));
                else
                    rt[0].Add(ky, "");
            }
        }
        return rt;
    }

    private static  long deConstrainAssyComponent(ComponentOccurrence co)
    {
        // ' Deletes all constraints on an occurrence
        // ' !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
        long ct = 0;
        foreach (AssemblyConstraint cs in co.Constraints)
        {
            cs.Delete();
            ct = ct + 1;
        }

        return ct;
    }

    public static  long deConstrainAssyDocument(AssemblyDocument ad)
    {
        // ' Calls deConstrainAssyComponent over all occurrences
        // ' in an assembly to remove all their constraints
        // ' !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
        // ' !!!That goes DOUBLE for THIS function!!!
        ComponentOccurrence co;

        return ad.ComponentDefinition.Occurrences.Cast<dynamic>()
            .Aggregate<dynamic, long>(0, (current, co) => current + deConstrainAssyComponent(co));
    }

    public static  Dictionary dcPartsInGeniusOrNot()
    {
        var rt = new Dictionary();
        var dcInGns = new Dictionary();
        var dcNotIn = dcRemapByPtNum(dcAiDocComponents(aiDocActive()));
        // dcNotIn = dcAiDocComponents(aiDocActive())
        // dcNotIn = dcAssyDocsByPtNum(aiDocActive())
        {
            var withBlock = dcFrom2Fields(cnGnsDoyle().Execute("select Item from vgMfiItems"), "Item", "Item");
            foreach (var ky in withBlock.Keys)
            {
                {
                    if (!dcNotIn.Exists(ky)) continue;
                    dcInGns.Add(ky, dcNotIn.get_Item(ky));
                    dcNotIn.Remove(ky);
                }
            }
        }
        // Stop

        rt.Add("INGNS", dcInGns);
        rt.Add("NOTIN", dcNotIn);
        return rt;
    }

    public static  Dictionary d0g4f1()
    {
        var rt = new Dictionary();
        var cn = cnGnsDoyle();
        var rs = cn.Execute("select Item from vgMfiItems");
        return rt;
    }

    public static  ComponentOccurrence compOccFromProxy(ComponentOccurrence oc)
    {
        ComponentOccurrenceProxy px;

        // If TypeOf oc Is Inventor.ComponentOccurrenceProxy Then
        // px = oc
        // Stop
        // compOccFromProxy = compOccFromProxy(px.ContainingOccurrence)
        // Else
        return oc;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="ob"></param>
    /// <returns></returns>
    public static  kyPick nuPicker(kyPick ob = null)
    {
        return ob ?? new kyPick();
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="ob"></param>
    /// <returns></returns>
    public static  dcSplitter nuSplitter(dcSplitter ob = null)
    {
        return ob ?? new dcSplitter();
    }

    public static Dictionary dcAiPartDocsWithRMv0(Dictionary dcIn, long WantOut = 0)
    {
        while (true)
        {
            var rt = new Dictionary[2];

            if (WantOut < 0 | WantOut > 1)
            {
                WantOut = 1;
                continue;
            }

            // With nuSplitter().WithSel(New kyPickAiDocWithRM)
            {
                var withBlock = new kyPickAiDocWithRM();
                rt[0] = withBlock.dcIn();
                rt[1] = withBlock.dcOut();
                foreach (var ky in dcIn.Keys)
                {
                    {
                        var withBlock1 = withBlock.dcFor(dcIn.get_Item(ky));
                        if (withBlock1.Exists(ky))
                            Debugger.Break();
                        else
                            withBlock1.Add(ky, dcIn.get_Item(ky));
                    }
                }
            }
            return rt[WantOut];
            break;
        }
    }
    // Debug.Print txDumpLs(dcAiPartDocsWithRMv0(dcAiDocComponents(aiDocActive()), 1).Keys)

    public static  kyPick kyScanned(Dictionary dcIn, kyPick pkr = null)
    {
        while (true)
        {
            if (pkr == null)
            {
                pkr = new kyPick();
            }
            else
                return pkr.AfterScanning(dcIn);
        }
    }
    // Debug.Print txDumpLs(kyScanned(dcAiDocComponents(aiDocActive()), New kyPickAiPartVsAssy).dcIn().Keys)

    public static  Dictionary dcAiDocsPicked(Dictionary dcIn, kyPick pkr = null, long WantOut = 0)
    {
        while (true)
        {
            var rt = new Dictionary[2];

            if (pkr == null)
            {
                pkr = new kyPick();
            }
            else if (WantOut < 0 | WantOut > 1)
            {
                WantOut = 1;
            }
            else
            {
                {
                    rt[0] = pkr.dcIn();
                    rt[1] = pkr.dcOut();
                    foreach (var ky in dcIn.Keys)
                    {
                        {
                            var withBlock1 = pkr.dcFor(dcIn.get_Item(ky));
                            if (withBlock1.Exists(ky))
                                Debugger.Break();
                            else
                                withBlock1.Add(ky, dcIn.get_Item(ky));
                        }
                    }
                }
                return rt[WantOut];
            }
        }
    }
    // Debug.Print txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), 1).Keys)
    // Debug.Print txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), New kyPickAiDocContentCtr, 0).Keys)

    public static  Dictionary dcAiPartDocsWithRM(Dictionary dcIn, long WantOut = 0)
    {
        return dcAiDocsPicked(dcAiDocsPicked(dcIn, new kyPickAiPartVsAssy(), 0), new kyPickAiDocWithRM(), WantOut);
    }
    // Debug.Print txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)

    public static  void d0g5f0()
    {
    }

    public static Dictionary dcAiDocGrpsByForm(Dictionary dcIn)
    {
        // dcAiDocGrpsByForm -- Separate a Dictionary
        // of Inventor Documents into
        // categorical sub-Dictionaries
        // according to various criteria:
        // - PRCH Purchased Items
        // - ASSY Assemblies
        // - HDWR Hardware (Content Center)
        // - DBAR Structural Parts (was BSTK)
        // Subtype NOT Sheet Metal
        // (some "Sheet Metal" Parts
        // might also technically
        // belong here. see below)
        // - MAYB Likely Structural Parts
        // Sheet Metal subtype, but
        // has either no flat pattern,
        // or an invalid one.
        // - SHTM Sheet Metal Parts
        // Indicated both by Subtype
        // and presence of a valid
        // flat pattern.
        // 
        // Presence in Genius, a distinction
        // originally intended to be made here,
        // is now planned to be made to a separate
        // Dictionary, possibly also subcategorized,
        // to be processed in conjunction with
        // the results of this function.
        // 
        // The notion of passing different subgroups
        // of this Dictionary to separate handlers
        // for more specialized processing, while
        // still an option, is no longer considered
        // its primary role. Instead, the set is
        // expected to be used in a form application
        // which will present the various groups to
        // the user for review, and modification as
        // desired or necessary.
        // 
        // REV[2022.03.08.1212] All new text
        // in function description above.
        // see notes_2022-0308_general-01.txt
        // for prior description
        // 
        // Dim pkGns As kyPick
        kyPick pkPrt;
        kyPick pkCtC;
        kyPick pkSht;
        kyPick pkMbe;

        var rt = new Dictionary();
        // REV[2022.03.08.1112]
        // Disabled split on presence
        // in Genius. Believe better
        // addressed separately
        // ' separate items already in Genius
        // ' from those not yet in
        // pkGns = nuPicker(' New kyPickInGenius').AfterScanning(dcIn)
        // ' NOTE: no further processing
        // ' implemented on this yet
        // ' MIGHT be better applied
        // ' at a different stage?
        // REV[2022.03.08.1115]
        // Add division on Purchased Parts
        // with "out" Dictionary replacing
        // main for Part/Assy separation.
        var pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);

        {
            rt.Add("PRCH", pkBuy.dcIn);

            // ' separate parts from assemblies
            pkPrt = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(pkBuy.dcOut);
        }

        {
            rt.Add("ASSY", pkPrt.dcOut);

            // ' isolate Content Center
            // ' parts from the rest
            pkCtC = nuPicker(new kyPickAiDocContentCtr()).AfterScanning(pkPrt.dcIn);
        }

        {
            rt.Add("HDWR", pkCtC.dcIn);

            // ' separate (potential) sheet
            // ' metal parts from non-sheet
            pkSht = nuPicker(new kyPickAiSheetMetal()).AfterScanning(pkCtC.dcOut);
        }

        {
            rt.Add("DBAR", pkSht.dcOut);
            pkMbe = nuPicker(new kyPickAiShMtl4sure()).AfterScanning(pkSht.dcIn);
        }

        {
            rt.Add("SHTM", pkMbe.dcIn);
            rt.Add("MAYB", pkMbe.dcOut);
        }

        Debug.Print(""); // Breakpoint Landing

        return rt;
    }

    public static  Dictionary dcAiDocGrpsByFormAndIfac(Dictionary dcIn)
    {
        // dcAiDocGrpsByFormAndIfac -- Separate a Dictionary
        // of Inventor Documents into
        // categorical sub-Dictionaries
        // according to various criteria:
        // - PRCH Purchased Items
        // - ASSY Assemblies
        // - IASM iAssembly Factories
        // - IPRT iPart Factories
        // - HDWR Hardware (Content Center)
        // - DBAR Structural Parts (was BSTK)
        // Subtype NOT Sheet Metal
        // (some "Sheet Metal" Parts
        // might also technically
        // belong here. see below)
        // - MAYB Likely Structural Parts
        // Sheet Metal subtype, but
        // has either no flat pattern,
        // or an invalid one.
        // - SHTM Sheet Metal Parts
        // Indicated both by Subtype
        // and presence of a valid
        // flat pattern.
        // 
        // REV[2023.01.24.0921]
        // copied from dcAiDocGrpsByForm to produce
        // new dynamic with additional groupings for
        // iPart (IPRT) and iAssembly (IASM) members.
        // additional groups for their corresponding
        // Factories will likely also be added.
        // 
        // text of prior REV[2022.03.08.1212] removed.
        // see dcAiDocGrpsByForm for that.
        // 
        Dictionary wk;
        // REV[2023.01.24.1156]
        // add working Dictionary
        // to collect iAssembly
        // and iPart Factories
        Document mb;
        Document md;
        string fp;

        // Dim pkGns As kyPick
        kyPick pkPrt;
        kyPick pkCtC;
        kyPick pkSht;
        kyPick pkMbe;

        // REV[2023.01.24.1009]
        // add new pickers for
        // iAssemblies and iParts
        kyPick pkIas;
        kyPick pkIpt;

        var rt = new Dictionary();
        // REV[2022.03.08.1112]
        // Disabled split on presence
        // in Genius. Believe better
        // addressed separately
        // ' separate items already in Genius
        // ' from those not yet in
        // pkGns = nuPicker(' New kyPickInGenius').AfterScanning(dcIn)
        // ' NOTE: no further processing
        // ' implemented on this yet
        // ' MIGHT be better applied
        // ' at a different stage?
        // REV[2022.03.08.1115]
        // Add division on Purchased Parts
        // with "out" Dictionary replacing
        // main for Part/Assy separation.
        var pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);

        {
            if (dcIn.Count > 0)
                rt.Add("PRCH", pkBuy.dcIn);

            // ' separate parts from assemblies
            pkPrt = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(pkBuy.dcOut);
        }

        {
            // ' separate iAssembly members
            // ' from stand-alone assemblies
            pkIas = nuPicker(new kyPickAiAssyMember()).AfterScanning(pkPrt.dcOut);

            // ' isolate Content Center
            // ' parts from the rest
            pkCtC = nuPicker(new kyPickAiDocContentCtr()).AfterScanning(pkPrt.dcIn);
        }

        {
            if (pkIas.dcOut.Count > 0)
                rt.Add("ASSY", pkIas.dcOut);
            {
                var withBlock1 = pkIas.dcIn;
                wk = new Dictionary();

                foreach (var ky in withBlock1.Keys)
                {
                    {
                        var withBlock2 = aiDocAssy(withBlock1.get_Item(ky)).ComponentDefinition;
                        mb = withBlock2.Parent;
                        {
                            var withBlock3 = withBlock2.iAssemblyMember.ParentFactory.Parent;
                            md = withBlock3.Parent;
                            fp = md.FullDocumentName;

                            {
                                if (!wk.Exists(fp))
                                {
                                    wk.Add(fp, new Dictionary());
                                    dcOb(wk.get_Item(fp)).Add("", md);
                                }

                                {
                                    var withBlock5 = dcOb(wk.get_Item(fp));
                                    if (withBlock5.Exists(ky))
                                        Debugger.Break();
                                    else
                                        withBlock5.Add(ky, mb);
                                }
                            }
                        }
                    }
                }

                if (wk.Count > 0)
                    rt.Add("IASM", wk);
            }
        }

        {
            if (pkCtC.dcIn.Count > 0)
                rt.Add("HDWR", pkCtC.dcIn);

            // ' separate iPart members
            // ' from stand-alone parts
            pkIpt = nuPicker(new kyPickAiPartMember()).AfterScanning(pkCtC.dcOut);
        }

        {
            {
                var withBlock1 = pkIpt.dcIn;
                wk = new Dictionary();

                foreach (var ky in withBlock1.Keys)
                {
                    {
                        var withBlock2 = aiDocPart(withBlock1.get_Item(ky)).ComponentDefinition;
                        mb = withBlock2.Document; // .Parent
                        md = aiDocPart(withBlock2.iPartMember.ParentFactory.Parent); // .PropertySets.Parent ' .Parent
                        {
                            var withBlock3 = md;
                            fp = md.FullDocumentName;

                            {
                                if (!wk.Exists(fp))
                                {
                                    wk.Add(fp, new Dictionary());
                                    dcOb(wk.get_Item(fp)).Add("", md);
                                }

                                {
                                    var withBlock5 = dcOb(wk.get_Item(fp));
                                    if (withBlock5.Exists(ky))
                                        Debugger.Break();
                                    else
                                        withBlock5.Add(ky, mb);
                                }
                            }
                        }
                    }
                }

                if (wk.Count > 0)
                    rt.Add("IPRT", wk);
            }

            // ' add iPart Factories to
            // ' Dictionary of non-Members
            {
                var withBlock1 = pkIpt.dcOut;
                foreach (var ky in wk.Keys)
                {
                    if (withBlock1.Exists(ky))
                        Debugger.Break();
                    else
                        withBlock1.Add(ky, dcOb(wk.get_Item(ky)).get_Item(""));
                }
            }

            // ' separate (potential) sheet
            // ' metal parts from non-sheet
            pkSht = nuPicker(new kyPickAiSheetMetal()).AfterScanning(pkIpt.dcOut);
        }

        {
            if (pkSht.dcOut.Count > 0)
                rt.Add("DBAR", pkSht.dcOut);
            pkMbe = nuPicker(new kyPickAiShMtl4sure()).AfterScanning(pkSht.dcIn);
        }

        {
            if (pkMbe.dcIn.Count > 0)
                rt.Add("SHTM", pkMbe.dcIn);
            if (pkMbe.dcOut.Count > 0)
                rt.Add("MAYB", pkMbe.dcOut);
        }

        Debug.Print(""); // Breakpoint Landing

        return rt;
    }

    public static  Dictionary d0g5f2(Dictionary dcIn)
    {
        // function d0g5f2
        // 
        // INITIATED[2021.03.23]
        // this dynamic on dcAiDocGrpsByForm
        // is intended to separate items
        // in Genius from those not yet in
        // and purchased items from those
        // to be made, cross-referencing
        // the two to determine individual
        // needs for processing.
        // 
        // presently in a nonfunctional state
        // as End of Day approaches. will hope
        // to continue development tomorrow
        // 
        kyPick pkPvA;
        kyPick pkCtC;
        kyPick pkSht;

        var rt = new Dictionary();

        // ' separate items already in Genius
        // ' from those not yet in
        var pkGns = nuPicker(new kyPickInGenius()).AfterScanning(dcIn);

        // ' separate purchased items
        // ' from those to be made
        var pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);

        // ' NOTE: no further processing
        // ' implemented on this yet
        // ' MIGHT be better applied
        // ' at a different stage?
        // ' separate parts from assemblies
        // pkPvA = nuPicker(' New kyPickAiPartVsAssy').AfterScanning(dcIn)
        // 'rt.Add "ASSY", dck pkPvA.dcOut
        // ' isolate Content Center
        // ' parts from the rest
        // pkCtC = nuPicker(' New kyPickAiDocContentCtr').AfterScanning(pkPvA.dcIn)
        // 'rt.Add "HDWR", pkCtC.dcIn
        // ' separate (potential)
        // ' sheet metal parts
        // ' from non-sheet
        // pkSht = nuPicker(' New kyPickAiSheetMetal').AfterScanning(pkCtC.dcOut)
        // rt.Add "SHTM", pkSht.dcIn
        // rt.Add "BSTK", pkSht.dcOut
        Debug.Print(""); // Breakpoint Landing

        return rt;
    }

    public static  Dictionary d0g5f3(Dictionary dcIn)
    {
        // function d0g5f3 -- essentially a recreation of dcAiDocGrpsByForm
        // 
        kyPick pkPrt;
        kyPick pkCtC;
        kyPick pkSht;
        kyPick pkMbe;

        // Dim pkGns As kyPick
        // Dim pk___ As kyPick
        var rt = new Dictionary();

        // pkGns = nuPicker(New kyPickInGenius).AfterScanning(dcIn)
        var pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);

        {
            rt.Add("PRCH", pkBuy.dcIn);
            pkPrt = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(pkBuy.dcOut);
        }

        {
            rt.Add("ASSY", pkPrt.dcOut);
            pkCtC = nuPicker(new kyPickAiDocContentCtr()).AfterScanning(pkPrt.dcIn);
        }

        {
            rt.Add("HDWR", pkCtC.dcIn);
            pkSht = nuPicker(new kyPickAiSheetMetal()).AfterScanning(pkCtC.dcOut);
        }

        {
            rt.Add("DBAR", pkSht.dcOut);
            pkMbe = nuPicker(new kyPickAiShMtl4sure()).AfterScanning(pkSht.dcIn);
        }

        {
            rt.Add("MAYB", pkMbe.dcOut);
            rt.Add("SHTM", pkMbe.dcIn);
        }

        return rt;
    }

    public static  long dcDepthAiDocGrp(Dictionary dc)
    {
        dynamic vr;
        long rt;

        {
            long mx = dc.Count;

            if (mx == 0)
                return 0; // indeterminate
            long dx = 0;

            long ck;
            do
            {
                // vr = new [] {.get_Item(.Keys(dx)))
                var ob = obOf(dc.get_Item(dc.Keys(dx)));
                switch (ob)
                {
                    case null:
                        ck = -1;
                        break;
                    case Dictionary:
                    {
                        ck = dcDepthAiDocGrp(ob);
                        if (ck > 0)
                            ck = 1 + ck;
                        break;
                    }
                    case Document:
                        ck = 1;
                        break;
                    default:
                        ck = -1;
                        break;
                }

                dx = dx + 1;
                if (dx > mx)
                    ck = -1;
            } while (ck == 0) // vr(0)// invalid// invalid // invalid
                ; // indeterminate

            return ck;
        }
    }

    public static fmIfcTest05A nu_fmIfcTest05A(Dictionary dcIn = null)
    {
        {
            var withBlock = new fmIfcTest05A();
            return withBlock.Using(dcIn);
        }
    }

    public static  fmTest05 nu_fmTest05A(Dictionary dcIn = null)
    {
        Debugger.Break(); // DO NOT USE THIS FUNCTION!
        // instead, use the Interface
        // generator nu_fmIfcTest05A

        {
            var withBlock = new fmTest05();
            return withBlock.Holding(dcIn); // .Using(dcIn)
        }
    }

    public static  string lsAssyMembers(AssemblyDocument aiAssy)
    {
        var dc = dcAiDocsByPtNum(dcAssyComponentsImmediate(aiAssy)); // dcAiDocPartNumbers
        var pn = Constants.vbCrLf + aiAssy.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value + Constants.vbTab;
        var rt = pn + Join(dc.Keys, pn);

        {
            var withBlock = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(dc);
            {
                var withBlock1 = withBlock.dcOut;
                foreach (var ky in withBlock1.Keys)
                    rt = rt + lsAssyMembers(aiDocument(obOf(withBlock1.get_Item(ky))));
            }
        }

        return rt;
    }

    public static  string d0g6f0(Document AiDoc)
    {
        // ' Try to pick a distinct listing name
        // ' for a supplied Inventor Document

        {
            string ds;
            string rt;
            {
                var withBlock1 = AiDoc.PropertySets(gnDesign);
                rt = Trim(withBlock1.get_Item(pnPartNum).Value);
                ds = Trim(withBlock1.get_Item(pnDesc).Value);
            }

            if (Strings.Len(rt) > 0)
            {
                if (Strings.Len(ds) > 0)
                    rt = rt + ": " + ds;
            }
            else if (Strings.Len(ds) > 0)
                rt = ds;

            if (Strings.Len(rt) != 0) return rt;
            {
                ds = AiDoc.FullFileName;
                if (Strings.Len(ds) > 0)
                {
                    {
                        var withBlock1 = nuFso().GetFile(ds);
                        rt = withBlock1.Name + " (" + withBlock1.ParentFolder.Path + ")";
                    }
                }
                else
                    rt = AiDoc.DisplayName;
            }

            return rt;
        }
    }

    public static  void d0g6f1()
    {
        // '
        // ' testing form class fmTest0
        // '
        {
            var withBlock = new fmTest0();
            withBlock.imTNail.Visible = false;
            Debug.Print.Controls.Count();
            Debugger.Break();
        }
    }

    public static  Dictionary d0g6f2(Dictionary dc)
    {
        // Call this one from inside dcAiDocGrpsByForm (above)
        // Try: debug.Print txDumpLs(d0g6f2(pkPvA.dcIn).Keys)
        // 

        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                Document ad = aiDocument(dc.get_Item(ky));
                if (ad == null)
                {
                }
                else
                {
                    Property pr = ad.PropertySets(gnDesign).get_Item(pnFamily);
                    rt.Add(ky, pr);
                    {
                        pr.Value = "R-PTS";
                    }
                }
            }
        }

        return rt;
    }

    public static  void d0g6f3()
    {
        // '
        // ' testing new null form class fmEmpty
        // '
        {
            var withBlock = new fmEmpty();
            // .imTNail.Visible = False
            {
                var withBlock1 = withBlock.Controls.Add("Forms.ComboBox.1", "test", true);
                Debug.Print.Name();
                withBlock1.Left = 10;
                withBlock1.Top = 10;
            }
            Debug.Print.Controls.Count();
            withBlock.Show(1);
            Debugger.Break();
        }
    }

    public static  string d0g7f0()
    {
        // This function used to transfer Property Values
        // from blank model files GR12 ~ GR20
        // to new versions generated from Intraflo's
        // supplied STEP files. Save for reference,
        // but this version should not likely be used
        // as is for other tasks without review.
        // 
        Property prTg;

        var rt = dcAiDocsByPtNum(dcAssyComponentsImmediate(aiDocActive()));
        {
            foreach (var ky in rt.Keys)
            {
                Debug.Print(ky);
                Document sd = aiDocument(obOf(rt.get_Item(ky)));
                if (UCase(Left(ky, 2)) != "GR") continue;
                string sn = sd.PropertySets(gnDesign).get_Item(pnStockNum).Value;
                if (!rt.Exists(sn)) continue;
                Document td = aiDocument(obOf(rt.get_Item(sn)));
                Debug.Print(""); // Stop
                PropertySet psTg = td.PropertySets(gnCustom);
                var dcPr = dcAiPropsInSet(psTg);
                PropertySet psSc = sd.PropertySets(gnCustom);
                foreach (Property prSc in psSc)
                {
                    if (dcPr.Exists(prSc.Name))
                    {
                        {
                            var withBlock1 = psTg.get_Item(prSc.Name);
                            withBlock1.Value = prSc.Value;
                            Debug.Print(""); // Landing Point -- Ctrl-F8 to here
                        }
                    }
                    else
                    {
                        psTg.Add(null, prSc.Name);
                        Debug.Print(""); // Landing Point -- Ctrl-F8 to here
                    }
                }

                psSc = sd.PropertySets(gnDesign);
                {
                    var withBlock1 = td.PropertySets(gnDesign);
                    foreach (var pn in new[] { pnPartNum, pnStockNum, pnFamily, pnDesc, pnCatWebLink })
                    {
                        // .get_Item(pnStockNum).Value = psSc.get_Item(pnStockNum).Value
                        // .get_Item(pnFamily).Value = psSc.get_Item(pnFamily).Value
                        // .get_Item(pnCatWebLink).Value = psSc.get_Item(pnCatWebLink).Value
                        // .get_Item(pnDesc).Value = psSc.get_Item(pnDesc).Value
                        // .get_Item(pnPartNum).Value = psSc.get_Item(pnPartNum).Value
                        // .get_Item(pn).Value = psSc.get_Item(pn).Value
                        withBlock1.get_Item(Convert.ToHexString(pn)).Value =
                            psSc.get_Item(Convert.ToHexString(pn)).Value;
                        Debug.Print(""); // Landing Point -- Ctrl-F8 to here
                    }
                }
            }
        }
    }

    public static  string d0g8f0()
    {
        for (long dx = 1; dx <= 16; dx++)
        {
            var fn = "Specification" + Convert.ToHexString(dx);
            {
                var withBlock = cnGnsDoyle().Execute(Join(new[]
                {
                    "select distinct", fn, "from vgMfiItems", "where Family = 'D-BAR'", "and", fn, "is not null", "and",
                    fn, "<> ''", "order by", fn, ";"
                }, " "));
                if (withBlock.BOF | withBlock.EOF)
                {
                }
                else
                {
                    Debug.Print("[" + fn + "]");
                    Debug.Print.GetString();
                }
            }
        }
    }

    public static  string d0g9f0(Document ad = null, string pn = "")
    {
        while (true)
        {
            if (ad == null)
            {
                ad = ThisApplication.ActiveDocument;
                continue;
            }

            if (Strings.Len(pn) < 1)
            {
                pn = d0g9f3(ad);
                continue;
            }

            const string bp = @"C:\Doyle_Vault\Designs\Misc\andrewT\";
            var vw = ad.Views.get_Item(1); // ThisApplication.ActiveView
            var cm = vw.Camera;

            {
                // Debug.Print .Left, .Top
                // Debug.Print .Width, .Height

                Debug.Print(""); // breakpoint anchor

                {
                    var withBlock1 = vw.Camera;
                    withBlock1.ViewOrientationType = kIsoTopRightViewOrientation;
                    withBlock1.Fit();
                    withBlock1.Apply();
                }
                vw.Fit();
                vw.Update();
                Debug.Print(""); // breakpoint anchor
                // .SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                vw.SaveAsBitmap(bp + pn + "-I.png", vw.Width, vw.Height);
                {
                    var withBlock1 = vw.Camera;
                    withBlock1.ViewOrientationType = kFrontViewOrientation;
                    withBlock1.Fit();
                    withBlock1.Apply();
                }
                vw.Fit();
                vw.Update();
                Debug.Print(""); // breakpoint anchor
                // .SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                vw.SaveAsBitmap(bp + pn + "-F.png", 0, 0);
                {
                    var withBlock1 = vw.Camera;
                    withBlock1.ViewOrientationType = kTopViewOrientation;
                    withBlock1.Fit();
                    withBlock1.Apply();
                }
                vw.Fit();
                vw.Update();
                Debug.Print(""); // breakpoint anchor
                // .SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                vw.SaveAsBitmap(bp + pn + "-T.png", 0, 0);
                vw.GoHome();
                vw.Update();
            }

            Debug.Print(""); // breakpoint anchor

            return "";
        }
    }

    public static  string d0g9f1(AssemblyDocument ad)
    {
        {
            var withBlock = ad.ComponentDefinition;
            if (!withBlock.IsiAssemblyFactory) return "";
            {
                var withBlock1 = withBlock.iAssemblyFactory;
                foreach (iAssemblyTableRow rw in withBlock1.TableRows)
                {
                    {
                        var withBlock2 = rw;
                        Debug.Print.MemberName();
                    }
                }
            }
        }

        return "";
    }

    public static  void d0g9f2(ComponentOccurrence oc)
    {
        {
            if (oc.IsiAssemblyMember || oc.IsiPartMember)
            {
            }

            Debugger.Break();
        }
    }

    public static  void d0g9f2as(AssemblyComponentDefinition cd)
    {
        {
            // .IsiAssemblyMember
            // .iAssemblyMember
            {
                var withBlock1 = cd.iAssemblyMember;
            }
        }
    }

    public static  string d0g9f3(AssemblyDocument ad)
    {
        {
            var withBlock = ad.ComponentDefinition.Occurrences.get_Item(1);
            var cp = aiDocAssy(withBlock.Definition.Document);
            return cp == null ? "NO-NUM-ASSY" : cp.PropertySets(gnDesign).get_Item(pnPartNum).Value;
        }
    }

    public static  void PlaceInAssembly()
    {
        {
            var withBlock = ThisApplication;
            if (withBlock.ActiveDocumentType == kPartDocumentObjectOr.ActiveDocumentType !=
                kAssemblyDocumentObjectThen) return;
            Document cd = withBlock.ActiveDocument;
            var dc = dcAiAssyDocs(dcAiDocsVisible());
            dc.Remove(cd.FullDocumentName);
            Document ad;
            {
                var withBlock1 = nuSelAiDoc().WithList(dc.Keys);
                VbMsgBoxResult rp;
                do
                {
                    var nm = withBlock1.GetReply();
                    if (dc.Exists(nm))
                    {
                        ad = dc.get_Item(nm);
                        rp = Constants.vbOK;
                    }
                    else
                    {
                        ad = null;
                        rp = MessageBox.Show("No Valid Assembly Selected.", "No Assembly", Constants.vbRetryCancel);
                    }
                } while (rp == Constants.vbRetry) // Try Again?
                    ;
            }

            if (ad == null)
                Debug.Print("");
            else
            {
                ad.Activate();
                var cm = withBlock.CommandManager;
                {
                    cm.PostPrivateEvent(kFileNameEvent, cd.FullDocumentName);
                    cm.ControlDefinitions.get_Item("AssemblyPlaceComponentCmd").Execute();
                }
            }
        }
    }
}