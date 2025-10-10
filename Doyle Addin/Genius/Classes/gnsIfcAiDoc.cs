class gnsIfcAiDoc
{
    private ADODB.Connection cn;
    private ADODB.Recordset rs;
    private ADODB.Recordset rsFamBuy;
    private ADODB.Field fdItem;
    private ADODB.Field fdFamily;
    private ADODB.Field fd;
    // Private fd          As ADODB.Field

    /// NOTE: the following SQL text constants

    /// are left over from the initial design

    /// of this class (under a different name).

    /// 

    /// Their code will remain in place until

    /// such time as their value may be better

    /// ascertained. Assuming they remain useful,

    /// they should be exported to a separate

    /// library module for storage of SQL source.

    /// 
    private const string sql01 = "" + "select Item, Family " + "from vgMfiItems " + "" + "";

    private const string sql02 = "" + "Select F.Family, F.Description1, " + "F.DefaultPlanningId As pln, " + "F.ProductCategory As cat " + "From vgMfiFamilies F " + "Where F.Type = 'R' " + "And F.FamilyGroup = 'PARTS' " + "";

    public Scripting.Dictionary Props(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (dc == null)
            Props = Props(AiDoc, new Scripting.Dictionary());
        else
            Props = assyProps(aiDocAssy(AiDoc), partProps(aiDocPart(AiDoc), gnrlProps(AiDoc, dc)));
    }

    private Scripting.Dictionary gnrlProps(Inventor.Document AiDoc, Scripting.Dictionary dc)
    {
        /// gnrlProps -- derived 2023.01.13 from partProps
        /// to collect (and set) Properties applicable
        /// to both Parts and Assemblies
        /// '
        /// original intent to remove Family assignment
        /// from Part Property section to a more general
        /// applicable to both Parts and Assemblies,
        /// however, that's proving a more challenging
        /// task than anticipated
        /// '
        /// presently just a stub, pending further review
        /// 
        if (AiDoc == null)
            gnrlProps = dc;
        else
            // With AiDoc
            // With .PropertySets
            // With .Item(gnDesign)
            // End With
            // End With
            // End With
            gnrlProps = dc;
    }

    private Scripting.Dictionary partProps(Inventor.PartDocument AiDoc, Scripting.Dictionary dc)
    {
        if (AiDoc == null)
            partProps = dc;
        else
        {
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            partProps = dcGeniusPropsPart(AiDoc, dc);
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }
    }

    private Scripting.Dictionary assyProps(Inventor.AssemblyDocument AiDoc, Scripting.Dictionary dc)
    {
        if (AiDoc == null)
            assyProps = dc;
        else
        {
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            assyProps = dcGeniusPropsAssy(AiDoc, dc);
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }
    }

    private void Class_Initialize()
    {
    }

    /// OBSOLETE SECTION

    /// 

    /// All functions below this comment

    /// are left over from the initial

    /// effort to create some form of

    /// Genius-oriented interface to

    /// Autodesk Inventor Documents,

    /// and particularly Part Documents.

    /// 

    /// The TestXX functions, originally Public,

    /// have been rendered Private to hide them

    /// from any client processes. Their code

    /// remains in place pending possible use

    /// in some future process(es).

    /// 

    /// They should at some point be removed,

    /// once their value is better established,

    /// and any useful portions incorporated

    /// into appropriate procedures.

    /// 

    private Inventor.BOMStructureEnum Test01(Inventor.PartDocument invDoc)
    {
        /// Present Role: Categorize Part Document
        /// 
        /// 
        string nmFamily;
        Inventor.BOMStructureEnum bomStruct;
        // 
        /// 

        {
            var withBlock = invDoc;
            nmFamily = withBlock.PropertySets(gnDesign).Item(pnFamily).Value;

            if (withBlock.ComponentDefinition.IsContentMember)
            {
                if (withBlock.ComponentDefinition.BOMStructure == kPurchasedBOMStructure)
                {
                    bomStruct = kPurchasedBOMStructure;
                    nmFamily = "D-HDWR";
                }
                else
                    System.Diagnostics.Debugger.Break();
            }
            else if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") > 0)
                bomStruct = kPurchasedBOMStructure;
            else if (g0f0(nmFamily) == kPurchasedBOMStructure)
                // this is almost certainly a purchased part
                bomStruct = kPurchasedBOMStructure;
            else
                bomStruct = kDefaultBOMStructure;
        }
        Test01 = bomStruct;
        /// 
        // 
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing Line for Debug use. Do not disable.
    }

    private Inventor.BOMStructureEnum Test02(Inventor.PartDocument invDoc)
    {
        /// Present Role: Categorize Part Document
        /// 
        /// 
        Scripting.Dictionary rt;
        // ''
        Inventor.PropertySet aiPropsUser;
        Inventor.PropertySet aiPropsDesign;
        // ''
        Inventor.Property prFamily;
        // Dim prPartNum   As Inventor.Property
        // Dim prRawMatl   As Inventor.Property 'pnRawMaterial
        // Dim prRmUnit    As Inventor.Property 'pnRmUnit
        // Dim prRmQty     As Inventor.Property 'pnRmQty
        // ''
        string nmFamily;
        // Dim mtFamily As String
        // ''' UPDATE[2018.05.30]:
        // '''     Rename variable Family to nmFamily
        // '''     to minimize confusion between code
        // '''     and comment text in searches.
        // '''     Also add variable mtFamily
        // '''     for raw material Family name
        // Dim pnStock As String
        // Dim qtUnit As String
        Inventor.BOMStructureEnum bomStruct;
        // Dim ck As VbMsgBoxResult
        // '
        // '''
        // 
        {
            var withBlock = invDoc;
            // Property Sets
            // With .PropertySets
            // aiPropsUser = .Item(gnCustom)
            // aiPropsDesign = .Item(gnDesign)
            // End With

            // Custom Properties
            // prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
            // prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
            // prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)

            // Family property is from Design, NOT Custom set
            // prFamily = aiGetProp(aiPropsDesign, pnFamily)
            // prPartNum = aiGetProp(aiPropsDesign, pnPartNum)
            // nmFamily = prFamily.Value
            nmFamily = withBlock.PropertySets(gnDesign).Item(pnFamily).Value;

            if (withBlock.ComponentDefinition.IsContentMember)
            {
                if (withBlock.ComponentDefinition.BOMStructure == kPurchasedBOMStructure)
                {
                    bomStruct = kPurchasedBOMStructure;
                    nmFamily = "D-HDWR";
                }
                else
                    System.Diagnostics.Debugger.Break();
            }
            else if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") > 0)
                // this is PROBABLY a purchased part
                bomStruct = kPurchasedBOMStructure;
            else if (g0f0(nmFamily) == kPurchasedBOMStructure)
                // this is almost certainly a purchased part
                bomStruct = kPurchasedBOMStructure;
            else
                bomStruct = kNormalBOMStructure;
        }
        Test02 = bomStruct;
        /// 
        // 
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing Line for Debug use. Do not disable.
    }

    private Inventor.BOMStructureEnum g0f0(string f)
    {
        {
            var withBlock = rsFamBuy;
            withBlock.Filter = "Family = '" + f + "'";
            if (withBlock.BOF | withBlock.EOF)
                g0f0 = kDefaultBOMStructure;
            else
                g0f0 = kPurchasedBOMStructure;
        }
    }
}