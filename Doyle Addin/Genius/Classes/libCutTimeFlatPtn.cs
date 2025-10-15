// MessageBox

// Autodesk Inventor API types

namespace Doyle_Addin.Genius.Classes;

public class libCutTimeFlatPtn
{
    // Local fallback conversion (cm per inch) if not provided elsewhere
    private const double cvLenIn2cm = 2.54;

    public static  void showPerimeterInches()
    {
        var rt = fpPerimeter(ThisApplication.ActiveDocument);
        switch (rt)
        {
            // MessageBox.Show "Total length of all loops on face: " & rt & " cm"
            case > 0:
                MessageBox.Show(@"Total length of all loops on face: " + rt / cvLenIn2cm + @" in");
                break;
            case < 0:
                MessageBox.Show(@"No Valid Flat Pattern Found");
                break;
            default:
                MessageBox.Show(@"Flat Pattern Has No Measurable Perimeter");
                break;
        }
    }

    public static  double fpPerimeterInch(Document oDoc, double ld = 0)
    {
        var rt = fpPerimeter(oDoc); // , ld / cvLenIn2cm)
        if (rt > 0)
            return rt / cvLenIn2cm;
        return rt;
    }

    public  static double fpPerimeter(Document oDoc, double ld = 0)
    {
        double rt;

        double ct = 0;
        if (oDoc is PartDocument pd)
        {
            var oFace = smPartFlatPatternTopFace(pd);
            if (oFace == null)
                // ' Simple 'error' indicator
                rt = -1;
            else
            {
                rt = 0;
                foreach (Edge oEdge in oFace.Edges)
                {
                    rt = rt + edgeLength(oEdge);
                    ct = 1 + ct;
                }
            }
        }
        else
            rt = -1;
        return rt; // + ct * ld
    }

    public static  double edgeLength(Edge ed)
    {
        double lg;

        {
            var withBlock = ed.Evaluator;
            withBlock.GetParamExtents(out var mn, out var mx);
            withBlock.GetLengthAtParam(mn, mx, out lg);
        }
        return lg;
    }

    public static  Face smPartFlatPatternTopFace(PartDocument oDoc)
    {
        return oDoc == null ? null : fpTopFaceIfShtMetal((ComponentDefinition)oDoc.ComponentDefinition);
    }

    public static  Face fpTopFaceIfShtMetal(ComponentDefinition oDef)
    {
        return oDef is SheetMetalComponentDefinition smcd ? smcdFlatPtnTopFace(smcd) : null;
    }

    public  static Face smcdFlatPtnTopFace(SheetMetalComponentDefinition oDef)
    {
        return fpTopFace(oDef.FlatPattern);
    }

    public  static Face fpTopFace(FlatPattern fp)
    {
        Face rt = null;

        var oZAxis = ThisApplication.TransientGeometry.CreateUnitVector();

        foreach (var oFace in fp.Body.Faces.Cast<Face>()
                     .TakeWhile(_ => rt == null)
                     .Where(oFace => oFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                     .Select(oFace => new { oFace, plane = (Plane)oFace.Geometry })
                     .Where(@t => @t.plane.Normal.IsParallelTo(oZAxis))
                     .Where(@t => @t.plane.RootPoint.Z <= 0.0000001)
                     .Select(@t => @t.oFace))
        {
            rt = oFace;
        }

        return rt;
    }
}