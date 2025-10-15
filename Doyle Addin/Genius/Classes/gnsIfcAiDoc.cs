namespace Doyle_Addin.Genius.Classes;

class gnsIfcAiDoc
{
    private ADODB.Connection cn;
    private ADODB.Recordset rs;
    private ADODB.Recordset rsFamBuy;
    private ADODB.Field fdItem;
    private ADODB.Field fdFamily;
    private ADODB.Field fd;
    // Private fd As ADODB.Field

    // NOTE: the following SQL text constants

    // are left over from the initial design

    // of this class (under a different name).

    // 

    // Their code will remain in place until

    // such time as their value may be better

    // ascertained. Assuming they remain useful,

    // they should be exported to a separate

    // library module for storage of SQL source.

    // 
    private const string sql01 = "" + "select Item, Family " + "from vgMfiItems " + "" + "";

    private const string sql02 = "" + "Select F.Family, F.Description1, " + "F.DefaultPlanningId As pln, " +
                                 "F.ProductCategory As cat " + "From vgMfiFamilies F " + "Where F.Type = 'R' " +
                                 "And F.FamilyGroup = 'PARTS' " + "";

    public Dictionary Props(Document AiDoc, Dictionary dc = null)
    {
        while (true)
        {
            if (dc != null) return assyProps(aiDocAssy(AiDoc), partProps(aiDocPart(AiDoc), gnrlProps(AiDoc, dc)));
            dc = new Dictionary();
            continue;

            break;
        }
    }

    private Dictionary gnrlProps(Document AiDoc, Dictionary dc)
    {
        // gnrlProps -- derived 2023.01.13 from partProps
        // to collect (and set) Properties applicable
        // to both Parts and Assemblies
        // '
        // original intent to remove Family assignment
        // from Part Property section to a more general
        // applicable to both Parts and Assemblies,
        // however, that's proving a more challenging
        // task than anticipated
        // '
        // presently just a stub, pending further review
        // 
        if (AiDoc == null)
        {
        }

        // With AiDoc
        // With .PropertySets
        // With .get_Item(gnDesign)
        // End With
        // End With
        // End With
        return dc;
    }

    private Dictionary partProps(PartDocument AiDoc, Dictionary dc)
    {
        if (AiDoc == null)
            return dc;
        Debug.Print(""); // Breakpoint Landing
        return dcGeniusPropsPart(AiDoc, dc);
        Debug.Print(""); // Breakpoint Landing
    }

    private Dictionary assyProps(AssemblyDocument AiDoc, Dictionary dc)
    {
        if (AiDoc == null)
            return dc;
        Debug.Print(""); // Breakpoint Landing
        return dcGeniusPropsAssy(AiDoc, dc);
        Debug.Print(""); // Breakpoint Landing
    }

    private void Class_Initialize()
    {
    }

    // OBSOLETE SECTION

    // 

    // All functions below this comment

    // are left over from the initial

    // effort to create some form of

    // Genius-oriented interface to

    // Autodesk Inventor Documents,

    // and particularly Part Documents.

    // 

    // The TestXX functions, originally Public,

    // have been rendered Private to hide them

    // from any client processes. Their code

    // remains in place pending possible use

    // in some future process(es).

    // 

    // They should at some point be removed,

    // once their value is better established,

    // and any useful portions incorporated

    // into appropriate procedures.

    // 

    private BOMStructureEnum Test01(PartDocument invDoc)
    {
        // Present Role: Categorize Part Document
        // 
        // 
        BOMStructureEnum bomStruct;
        // 
        // 

        {
            string nmFamily = invDoc.PropertySets(gnDesign).get_Item(pnFamily).Value;

            if (invDoc.ComponentDefinition.IsContentMember)
            {
                if (invDoc.ComponentDefinition.BOMStructure == kPurchasedBOMStructure)
                {
                    bomStruct = kPurchasedBOMStructure;
                    nmFamily = "D-HDWR";
                }
                else
                    Debugger.Break();
            }
            else if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") > 0)
                bomStruct = kPurchasedBOMStructure;
            else if (g0f0(nmFamily) == kPurchasedBOMStructure)
                // this is almost certainly a purchased part
                bomStruct = kPurchasedBOMStructure;
            else
                bomStruct = kDefaultBOMStructure;
        }
        return bomStruct;
        // 
        // 
        Debug.Print(""); // Landing Line for Debug use. Do not disable.
    }

    private BOMStructureEnum Test02(PartDocument invDoc)
    {
        // Present Role: Categorize Part Document
        // 
        // 
        Dictionary rt;
        // ''
        PropertySet aiPropsUser;
        PropertySet aiPropsDesign;
        // ''
        Property prFamily;
        // Dim prPartNum As Inventor.Property
        // Dim prRawMatl As Inventor.Property 'pnRawMaterial
        // Dim prRmUnit As Inventor.Property 'pnRmUnit
        // Dim prRmQty As Inventor.Property 'pnRmQty
        // ''
        // Dim mtFamily As String
        // ''' UPDATE[2018.05.30]:
        // ''' Rename variable Family to nmFamily
        // ''' to minimize confusion between code
        // ''' and comment text in searches.
        // ''' Also add variable mtFamily
        // ''' for raw material Family name
        // Dim pnStock As String
        // Dim qtUnit As String
        BOMStructureEnum bomStruct;
        // Dim ck As VbMsgBoxResult
        // '
        // '''
        // 
        {
            // Property Sets
            // With .PropertySets
            // aiPropsUser = .get_Item(gnCustom)
            // aiPropsDesign = .get_Item(gnDesign)
            // End With
            // Custom Properties
            // prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
            // prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
            // prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
            // Family property is from Design, NOT Custom set
            // prFamily = aiGetProp(aiPropsDesign, pnFamily)
            // prPartNum = aiGetProp(aiPropsDesign, pnPartNum)
            // nmFamily = prFamily.Value
            string nmFamily = invDoc.PropertySets(gnDesign).get_Item(pnFamily).Value;

            if (invDoc.ComponentDefinition.IsContentMember)
            {
                if (invDoc.ComponentDefinition.BOMStructure == kPurchasedBOMStructure)
                {
                    bomStruct = kPurchasedBOMStructure;
                    nmFamily = "D-HDWR";
                }
                else
                    Debugger.Break();
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
        return bomStruct;
        // 
        // 
        Debug.Print(""); // Landing Line for Debug use. Do not disable.
    }

    private BOMStructureEnum g0f0(string f)
    {
        {
            var withBlock = rsFamBuy;
            withBlock.Filter = "Family = '" + f + "'";
            if (withBlock.BOF | withBlock.EOF)
                return kDefaultBOMStructure;
            return kPurchasedBOMStructure;
        }
    }
}