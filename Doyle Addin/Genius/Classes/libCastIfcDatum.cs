namespace Doyle_Addin.Genius.Classes;

public class libCastIfcDatum
{
    private const string txVersion = "module libCastIfcDatum REV[2022.03.18.1136]";
    // 

    // 

    /// <summary>
    /// 
    /// </summary>
    /// <param name="ob"></param>
    /// <returns></returns>
    public static ifcDatum obIfcDatum(dynamic ob)
    {
        if (ob is Property)
        {
            {
                var withBlock = new ifcAiProperty();
                var caster = new libCastOb();
                return withBlock.Connect(caster.obAiProp(ob));
            }
        }

        {
            var withBlock = new ifcDatum();
            return withBlock.Connect(ob);
        }
    }

    // END of Module libCastIfcDatum

    // 
    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    public static string Version()
    {
        return txVersion;
    }
}