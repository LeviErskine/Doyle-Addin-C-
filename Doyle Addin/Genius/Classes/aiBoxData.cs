using System;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using Inventor;

namespace Doyle_Addin.Genius.Classes
{
    // Safe C# rewrite of aiBoxData.
    // Where external dependencies were unclear or unavailable, code is commented rather than removed.
    class aiBoxData
    {
        private const string f01 = "#,##0.000";

        private Inventor.Point mn;
        private Inventor.Point mx;

        private double sc = 1;

        private Box Box { get; set; }

        private void SetBox(Box thisBox)
        {
            Box = thisBox;
            if (Box != null)
            {
                mn = Box.MinPoint;
                mx = Box.MaxPoint;
            }
            else
            {
                mn = null;
                mx = null;
            }
        }

        public aiBoxData UsingBox(Box ThisOne)
        {
            SetBox(ThisOne);
            return this;
        }

        public aiBoxData UsingOrBox(OrientedBox ThisOne)
        {
            Box = ThisApplication.TransientGeometry.CreateBox();

            {
                Box.Extend(ThisApplication.TransientGeometry.CreatePoint(ThisOne.DirectionOne.Length,
                    ThisOne.DirectionTwo.Length, ThisOne.DirectionThree.Length));
            }

            // Me.Box = ThisOne
            return UsingBox(Box);
        }

        public aiBoxData UsingBoxOb(object ThisOne)
        {
            return ThisOne switch
            {
                null => this,
                Box b => UsingBox(b),
                OrientedBox ob => UsingOrBox(ob),
                _ => this
            };
        }

        public aiBoxData UsingModel(Document ThisOne, long Oriented = 0)
        {
            return UsingPart(aiDocPart(ThisOne), Oriented).UsingAssy(aiDocAssy(ThisOne), Oriented);
        }

        private aiBoxData UsingPart(PartDocument ThisOne, long Oriented = 0)
        {
            if (ThisOne == null) return this;
            var withBlock = ThisOne.ComponentDefinition;
            return UsingBoxOb(Oriented == 0 ? withBlock.RangeBox : withBlock.OrientedMinimumRangeBox);
        }

        private aiBoxData UsingAssy(AssemblyDocument ThisOne, long Oriented = 0)
        {
            if (ThisOne == null) return this;
            var withBlock = ThisOne.ComponentDefinition;
            return UsingBoxOb(Oriented == 0 ? withBlock.RangeBox : withBlock.OrientedMinimumRangeBox);
        }

        public aiBoxData SortingDims(Box ThisBox = null)
        {
            while (true)
            {
                if (ThisBox == null)
                {
                    if (Box == null) return this;
                    ThisBox = Box;
                    continue;
                }

                // Original called aiBoxSortDown(ThisBox) which is not available here.
                // Keeping the incoming box as-is; comment left for later implementation.
                // SetBox(aiBoxSortDown(ThisBox));
                SetBox(ThisBox);
                return this;
                break;
            }
        }

        private double Span(double ptMin, double ptMax)
        {
            return sc * (ptMax - ptMin);
        }

        public double SpanX()
        {
            return mn == null || mx == null ? 0 : Span(mn.X, mx.X);
        }

        public double SpanY()
        {
            return mn == null || mx == null ? 0 : Span(mn.Y, mx.Y);
        }

        public double SpanZ()
        {
            return mn == null || mx == null ? 0 : Span(mn.Z, mx.Z);
        }

        public double[] SpansXYZ()
        {
            return new[] { SpanX(), SpanY(), SpanZ() };
        }

        public double[] SpansOrdered()
        {
            var arr = SpansXYZ();
            Array.Sort(arr); // ascending
            return arr;
        }

        public aiBoxData UsingInches(long Yes = 1)
        {
            if (Yes != 0)
                sc = 1 / 2.54;
            else
                sc = 1;
            return this;
        }

        public string Dump(long Form = 0)
        {
            // JSON and nuDcPopulator path kept for future work.
            // switch (Form)
            // {
            //     case 67518582:
            //         var _ = nuDcPopulator().Setting("X SPAN", Format(SpanX(), "#,##0.0000 '"))
            //             .Setting("Y SPAN", Format(SpanY(), "#,##0.0000 '"))
            //             .Setting("Z SPAN", Format(SpanZ(), "#,##0.0000 '"));
            //         break;
            //     default:
            //         break;
            // }
            return "X SPAN" + Constants.vbTab + "Y SPAN" + Constants.vbTab + "Z SPAN" + Constants.vbCrLf +
                   Strings.Format(SpanX(), f01) + Constants.vbTab + Strings.Format(SpanY(), f01) + Constants.vbTab +
                   Strings.Format(SpanZ(), f01);
        }

        public Dictionary<string, double> Dictionary(long Form = 3)
        {
            while (true)
            {
                // Dictionary -- return Dictionary of dimensions
                // keyed according to Form, a sum of:
                // 1 - "X", "Y", "Z", per Model
                // 2 - magnitudes "Min", "Mid", "Max"
                // 3 - BOTH sets of keys (1 + 2)
                var rt = new Dictionary<string, double>();

                if ((Form & 1) != 0)
                {
                    rt["X"] = SpanX();
                    rt["Y"] = SpanY();
                    rt["Z"] = SpanZ();
                }

                if ((Form & 2) != 0)
                {
                    var dm = SpansOrdered();
                    if (dm.Length >= 3)
                    {
                        rt["Min"] = dm[0];
                        rt["Mid"] = dm[1];
                        rt["Max"] = dm[2];
                    }
                }

                if (rt.Count != 0) return rt;
                Form = 3;
            }
        }
    }
}