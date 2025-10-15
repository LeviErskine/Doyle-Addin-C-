using Doyle_Addin.Genius.Forms;
using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class dvlObjMakers
{
    private const string txVersion = "dvlObjMakers REV[2022.03.16.0930]";

    public static wkgCls0 nu_wkgCls0(Document AiDoc = null)
    {
        {
            var withBlock = new wkgCls0();
            return withBlock.Using(AiDoc);
        }
    }

    public static gnsIfcAiDoc nu_gnsIfcAiDoc()
    {
        return new gnsIfcAiDoc();
    }

    public static iLogicIfc nuILogicIfc(Document Using = null)
    {
        {
            var withBlock = new iLogicIfc() // ifcVault
                ;
            if (RuleSource == RuleSourceEnum.Using)
            {
                if (Using == null)
                {
                    return WithRulesIn(Using);
                }

                return Itself;
            }

            if (Using == null)
            {
                return Itself;
            }

            return WithRulesIn(Using);
        }
    }

    public static fmIfcTest04A nu_fmIfcTest04A(Dictionary About = null)
    {
        {
            var withBlock = new fmIfcTest04A();
            return withBlock.Using(About);
        }
    }

    public static fmIfcMatlQty01 nu_fmIfcMatlQty01()
    {
        return new fmIfcMatlQty01();
    }

    public static fmGetList nu_FmGetList()
    {
        return new fmGetList();
    }

    public static fmTest0 newFmTest0()
    {
        return new fmTest0();
    }
    // Debug.Print newFmTest0().ft0g0f0(aiDocument(ThisApplication.ActiveDocument).Thumbnail)

    public static fmTest1 newFmTest1()
    {
        return new fmTest1();
    }

    public static fmTest2 newFmTest2()
    {
        return new fmTest2();
    }

    public static aiBoxData nuAiBoxData()
    {
        // ' Using "blank" version at this point
        return nuAiBoxDataRC0();
    }
    // Debug.Print nuAiBoxData().UsingInches().Sorted(aiDocPart(aiDocActive()).ComponentDefinition.RangeBox).Dump(0)

    // TESTING SECTION

    // 

    public static void tstFmTest1()
    {
        dynamic ky;

        const string nm = @"C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-072.ipt";

        // nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-050.ipt"
        // nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-051.ipt"
        {
            var withBlock = newFmTest1();
            if (withBlock.AskAbout(ThisApplication.Documents.ItemByName(nm)) != Constants.vbYes) return;
            {
                var withBlock1 = withBlock.ItemData;
                foreach (var ky in withBlock1.Keys)
                    Debug.Print(ky, withBlock1.get_Item(ky));
                Debugger.Break();
            }
        }
    }

    // VERSION / dynamic SECTION

    // 

    public static aiBoxData nuAiBoxDataRC1(dynamic arg1, long UseInches = -1)
    {
        dynamic ob;
        aiBoxData rt;

        if (UseInches < 0)
        {
            if (IsMissing(arg1))
            {
            }
            else if (IsObject(arg1))
            {
            }
        }

        {
            var withBlock = new aiBoxData();
            rt = withBlock.UsingInches(UseInches);
        }

        return rt;
    }

    public static aiBoxData nuAiBoxDataRC0()
    {
        return new aiBoxData();
    }

    // END of MODULE dvlObjMakers

    // 
    public static string dvlObjMakers()
    {
        return txVersion;
    }
}