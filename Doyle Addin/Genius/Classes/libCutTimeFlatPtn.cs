class libCutTimeFlatPtn
{
    public void showPerimeterInches()
    {
        double rt;

        rt = fpPerimeter(ThisApplication.ActiveDocument);
        if (rt > 0)
            // MsgBox "Total length of all loops on face: " & rt & " cm"
            MsgBox("Total length of all loops on face: " + rt / cvLenIn2cm + " in");
        else if (rt < 0)
            MsgBox("No Valid Flat Pattern Found");
        else
            MsgBox("Flat Pattern Has No Measurable Perimeter");
    }

    public double fpPerimeterInch(Inventor.Document oDoc, double ld = 0)
    {
        double rt;

        rt = fpPerimeter(oDoc); // , ld / cvLenIn2cm)
        if (rt > 0)
            fpPerimeterInch = rt / cvLenIn2cm;
        else
            fpPerimeterInch = rt;
    }

    public double fpPerimeter(Inventor.Document oDoc, double ld = 0)
    {
        Inventor.Face oFace;
        Inventor.Edge oEdge;
        double rt;
        double ct;

        ct = 0;
        if (oDoc is Inventor.PartDocument)
        {
            oFace = smPartFlatPatternTopFace(oDoc);
            if (oFace == null)
                // ' Simple 'error' indicator
                rt = -1;
            else
            {
                rt = 0;
                foreach (var oEdge in oFace.Edges)
                {
                    rt = rt + edgeLength(oEdge);
                    ct = 1 + ct;
                }
            }
        }
        else
            rt = -1;
        fpPerimeter = rt; // + ct * ld
    }

    public double edgeLength(Inventor.Edge ed)
    {
        double mn;
        double mx;
        double lg;

        {
            var withBlock = ed.Evaluator;
            withBlock.GetParamExtents(mn, mx);
            withBlock.GetLengthAtParam(mn, mx, lg);
        }
        edgeLength = lg;
    }

    public Inventor.Face smPartFlatPatternTopFace(PartDocument oDoc)
    {
        if (oDoc == null)
            smPartFlatPatternTopFace = null/* TODO Change to default(_) if this is not a reference type */;
        else
            smPartFlatPatternTopFace = fpTopFaceIfShtMetal(oDoc.ComponentDefinition);
    }

    public Inventor.Face fpTopFaceIfShtMetal(Inventor.ComponentDefinition oDef)
    {
        if (oDef is Inventor.SheetMetalComponentDefinition)
            fpTopFaceIfShtMetal = smcdFlatPtnTopFace(oDef);
        else
            fpTopFaceIfShtMetal = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.Face smcdFlatPtnTopFace(Inventor.SheetMetalComponentDefinition oDef)
    {
        smcdFlatPtnTopFace = fpTopFace(oDef.FlatPattern);
    }

    public Inventor.Face fpTopFace(Inventor.FlatPattern fp)
    {
        Inventor.UnitVector oZAxis;
        Inventor.Face oFace;
        Inventor.Face rt;

        oZAxis = ThisApplication.TransientGeometry.CreateUnitVector(0, 0, 1);

        foreach (var oFace in fp.Body.Faces)
        {
            // Only looking until we find a match
            if (rt == null)
            {
                {
                    var withBlock = oFace;
                    // Only interested in planar faces
                    if (withBlock.SurfaceType == kPlaneSurface)
                    {
                        {
                            var withBlock1 = aiPlane(withBlock.Geometry);
                            // Only interested in faces that have z-direction normal
                            if (withBlock1.Normal.IsParallelTo(oZAxis))
                            {
                                // Look for the face with Z = 0
                                if (withBlock1.RootPoint.Z <= 0.0000001)
                                    rt = oFace;
                            }
                        }
                    }
                }
            }
        }

        fpTopFace = rt;
    }
}