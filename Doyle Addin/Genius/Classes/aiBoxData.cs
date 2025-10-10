class aiBoxData
{
    private const string f01 = "#,##0.000";
    // Private Const f02 As String = "#,##0.0000 '"

    private Inventor.Box bx;
    private Inventor.Point mn;
    private Inventor.Point mx;

    private double sc;

    private void Class_Initialize()
    {
        sc = 1#;
    }

    public var Box
    {
        set
        {
            bx = ThisBox;
            mn = bx.MinPoint;
            mx = bx.MaxPoint;
        }
    }

    public Inventor.Box Box
    {
        get
        {
            Box = bx;
        }
    }

    public aiBoxData UsingBox(Inventor.Box ThisOne)
    {
        this.Box = ThisOne;
        UsingBox = this;
    }

    public aiBoxData UsingOrBox(Inventor.OrientedBox ThisOne)
    {
        bx = ThisApplication.TransientGeometry.CreateBox();

        {
            var withBlock = ThisOne;
            bx.Extend(ThisApplication.TransientGeometry.CreatePoint(withBlock.DirectionOne.Length, withBlock.DirectionTwo.Length, withBlock.DirectionThree.Length));
        }

        // Me.Box = ThisOne
        UsingOrBox = this.UsingBox(bx);
    }

    public aiBoxData UsingBoxOb(object ThisOne)
    {
        if (ThisOne == null)
            UsingBoxOb = this;
        else if (ThisOne is Inventor.Box)
            UsingBoxOb = UsingBox(ThisOne);
        else if (ThisOne is Inventor.OrientedBox)
            UsingBoxOb = UsingOrBox(ThisOne);
        else
            UsingBoxOb = this;
    }

    public aiBoxData UsingModel(Inventor.Document ThisOne, long Oriented = 0)
    {
        UsingModel = UsingPart(aiDocPart(ThisOne), Oriented).UsingAssy(aiDocAssy(ThisOne), Oriented);
    }

    public aiBoxData UsingPart(Inventor.PartDocument ThisOne, long Oriented = 0)
    {
        if (ThisOne == null)
            UsingPart = this;
        else
        {
            var withBlock = ThisOne.ComponentDefinition;
            UsingPart = UsingBoxOb(IIf(Oriented == 0, withBlock.RangeBox, withBlock.OrientedMinimumRangeBox));
        }
    }

    public aiBoxData UsingAssy(Inventor.AssemblyDocument ThisOne, long Oriented = 0)
    {
        if (ThisOne == null)
            UsingAssy = this;
        else
        {
            var withBlock = ThisOne.ComponentDefinition;
            UsingAssy = UsingBoxOb(IIf(Oriented == 0, withBlock.RangeBox, withBlock.OrientedMinimumRangeBox));
        }
    }

    public aiBoxData SortingDims(Inventor.Box ThisBox = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (ThisBox == null)
        {
            if (bx == null)
                SortingDims = this;
            else
                SortingDims = SortingDims(bx);
        }
        else
        {
            this.Box = aiBoxSortDown(ThisBox);
            SortingDims = this;
        }
    }

    private double Span(double ptMin, double ptMax)
    {
        Span = sc * (ptMax - ptMin);
    }

    public double SpanX()
    {
        SpanX = Span(mn.X, mx.X);
    }

    public double SpanY()
    {
        SpanY = Span(mn.Y, mx.Y);
    }

    public double SpanZ()
    {
        SpanZ = Span(mn.Z, mx.Z);
    }

    public double[] SpansXYZ()
    {
        double[] rt = new double[3];

        rt[0] = SpanX();
        rt[1] = SpanY();
        rt[2] = SpanZ();

        SpansXYZ = rt;
    }

    public double[] SpansOrdered()
    {
        SpansOrdered = sort3dimsUp(SpanX, SpanY, SpanZ);
    }

    public aiBoxData UsingInches(long Yes = 1)
    {
        if (Yes)
            sc = 1 / 2.54;
        else
            sc = 1;
        UsingInches = this;
    }

    public string Dump(long Form = 0)
    {
        Dump = "";
        // ConvertToJson(nuDcPopulator().Setting("X SPAN", Format$(me.SpanX, "#,##0.0000 '")).Setting("Y SPAN", Format$(me.SpanY, "#,##0.0000 '")).Setting("Z SPAN", Format$(me.SpanZ, "#,##0.0000 '")).Dictionary,vbTab)
        // ConvertToJson(nuDcPopulator().Setting("X SPAN", Round(me.SpanX,4)).Setting("Y SPAN", Round(me.SpanY,4)).Setting("Z SPAN", Round(me.SpanZ,4)).Dictionary,vbTab)
        switch (Form)
        {
            case 67518582:
                {
                    {
                        var withBlock = nuDcPopulator().Setting("X SPAN", Format(this.SpanX(), "#,##0.0000 '")).Setting("Y SPAN", Format(this.SpanY(), "#,##0.0000 '")).Setting("Z SPAN", Format(this.SpanZ(), "#,##0.0000 '")).Dictionary;
                    }

                    break;
                }

            default:
                {
                    Dump = "X SPAN" + Constants.vbTab + "Y SPAN" + Constants.vbTab + "Z SPAN" + Constants.vbNewLine + Strings.Format(this.SpanX(), f01) + Constants.vbTab + Strings.Format(this.SpanY(), f01) + Constants.vbTab + Strings.Format(this.SpanZ(), f01);
                    break;
                }
        }
    }

    public Scripting.Dictionary Dictionary(long Form = 3)
    {
        /// Dictionary -- return Dictionary of dimensions
        /// keyed according to Form, a sum of:
        /// 1 - "X", "Y", "Z", per Model
        /// 2 - magnitudes "Min", "Mid", "Max"
        /// (note that sorting keys in descending order
        /// produces values sorted in ascending order)
        /// 3 - BOTH sets of keys (1 + 2)
        /// 
        /// REV[2022.08.31.1444] Method Dictionary
        /// added to Class to support extraction
        /// of Dictionary Object for data export
        /// (see dcGnsPtProps_Rev20220830_inProg)
        /// 
        Scripting.Dictionary rt;
        double[] dm;

        if ((Form & 3) == 0)
            rt = Dictionary(3);
        else
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = rt;
                if (Form & 1)
                {
                    withBlock.Add("X", SpanX());
                    withBlock.Add("Y", SpanY());
                    withBlock.Add("Z", SpanZ());
                }

                if (Form & 2)
                {
                    dm = SpansOrdered();
                    withBlock.Add("Min", dm[0]);
                    withBlock.Add("Mid", dm[1]);
                    withBlock.Add("Max", dm[2]);
                }
            }
        }

        Dictionary = rt;
    }
}