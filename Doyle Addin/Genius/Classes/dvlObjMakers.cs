class SurroundingClass
{
    private const string txVersion = "dvlObjMakers REV[2022.03.16.0930]";

    public wkgCls0 nu_wkgCls0(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        {
            var withBlock = new wkgCls0();
            nu_wkgCls0 = withBlock.Using(AiDoc);
        }
    }

    public gnsIfcAiDoc nu_gnsIfcAiDoc()
    {
        nu_gnsIfcAiDoc = new gnsIfcAiDoc();
    }

    public iLogicIfc nuILogicIfc(Inventor.Document Using = null/* TODO Change to default(_) if this is not a reference type */)
    {
        {
            var withBlock = new iLogicIfc() // ifcVault
      ;
            if (RuleSource == RuleSourceEnum.Using)
            {
                if (Using == null)
                {
                    nuILogicIfc = WithRulesIn(Using);
                }
                else
                {
                    nuILogicIfc = Itself;
                }
            }
            else if (Using == null)
            {
                nuILogicIfc = Itself;
            }
            else
            {
                nuILogicIfc = WithRulesIn(Using);
            }
        }
    }

    public fmIfcTest04A nu_fmIfcTest04A(Scripting.Dictionary About = null/* TODO Change to default(_) if this is not a reference type */)
    {
        {
            var withBlock = new fmIfcTest04A();
            nu_fmIfcTest04A = withBlock.Using(About);
        }
    }

    public fmIfcMatlQty01 nu_fmIfcMatlQty01()
    {
        nu_fmIfcMatlQty01 = new fmIfcMatlQty01();
    }

    public fmGetList nu_FmGetList()
    {
        nu_FmGetList = new fmGetList();
    }

    public fmTest0 newFmTest0()
    {
        newFmTest0 = new fmTest0();
    }
    // Debug.Print newFmTest0().ft0g0f0(aiDocument(ThisApplication.ActiveDocument).Thumbnail)

    public fmTest1 newFmTest1()
    {
        newFmTest1 = new fmTest1();
    }

    public fmTest2 newFmTest2()
    {
        newFmTest2 = new fmTest2();
    }

    public aiBoxData nuAiBoxData()
    {
        // '  Using "blank" version at this point
        nuAiBoxData = nuAiBoxDataRC0();
    }
    // Debug.Print nuAiBoxData().UsingInches().Sorted(aiDocPart(aiDocActive()).ComponentDefinition.RangeBox).Dump(0)

    /// TESTING SECTION

    /// 

    public void tstFmTest1()
    {
        Variant ky;
        string nm;

        // nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-047.ipt"
        nm = @"C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-072.ipt";
        // nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-050.ipt"
        // nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-051.ipt"

        {
            var withBlock = newFmTest1();
            if (withBlock.AskAbout(ThisApplication.Documents.ItemByName(nm)) == Constants.vbYes)
            {
                {
                    var withBlock1 = withBlock.ItemData;
                    foreach (var ky in withBlock1.Keys)
                        Debug.Print(ky, withBlock1.Item(ky));
                    System.Diagnostics.Debugger.Break();
                }
            }
            else
            {
            }
        }
    }

    /// VERSION / VARIANT SECTION

    /// 

    public aiBoxData nuAiBoxDataRC1(Variant arg1, long UseInches = -1)
    {
        object ob;
        aiBoxData rt;

        if (UseInches < 0)
        {
            if (IsMissing(arg1))
            {
            }
            else if (IsObject(arg1))
            {
            }
            else
            {
            }
        }
        else
        {
        }

        {
            var withBlock = new aiBoxData();
            rt = withBlock.UsingInches(UseInches);
        }

        nuAiBoxDataRC1 = rt;
    }

    public aiBoxData nuAiBoxDataRC0()
    {
        nuAiBoxDataRC0 = new aiBoxData();
    }

    /// END of MODULE dvlObjMakers

    /// 
    public string dvlObjMakers()
    {
        dvlObjMakers = txVersion;
    }
}