class libCastIfcDatum
{
    private const string txVersion = "module libCastIfcDatum REV[2022.03.18.1136]";
    /// 

    /// 

    public ifcDatum obIfcDatum(object ob)
    {
        if (ob is Inventor.Property)
        {
            {
                var withBlock = new ifcAiProperty();
                obIfcDatum = withBlock.Connect(obAiProp(ob));
            }
        }
        else
        {
            var withBlock = new ifcDatum();
            obIfcDatum = withBlock.Connect(ob);
        }
    }

    /// END of Module libCastIfcDatum

    /// 
    public string libCastIfcDatum()
    {
        libCastIfcDatum = txVersion;
    }
}