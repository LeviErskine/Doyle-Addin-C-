global using static Doyle_Addin.Genius.Classes.dvl0;
global using static Doyle_Addin.Genius.Classes.dvl1;
global using static Doyle_Addin.Genius.Classes.dvl4;
global using static Doyle_Addin.Genius.Classes.dvlAiNameValMap;
global using static Doyle_Addin.Genius.Classes.dvlDict0;
global using static Doyle_Addin.Genius.Classes.dvlGnsIfc201904;
global using static Doyle_Addin.Genius.Classes.dvlObjMakers;
global using static Doyle_Addin.Genius.Classes.jsonConverter;
global using static Doyle_Addin.Genius.Classes.kyPickAiPartMember;
global using static Doyle_Addin.Genius.Classes.lib0;
global using static Doyle_Addin.Genius.Classes.libCastIfcDatum;
global using static Doyle_Addin.Genius.Classes.libCastOb;
global using static Doyle_Addin.Genius.Classes.libClipboardWin10;
global using static Doyle_Addin.Genius.Classes.libCutTimeFlatPtn;
global using static Doyle_Addin.Genius.Classes.libDcGeneral;
global using static Doyle_Addin.Genius.Classes.libDict;
global using static Doyle_Addin.Genius.Classes.libFmSelectors;
global using static Doyle_Addin.Genius.Classes.libFSys;
global using static Doyle_Addin.Genius.Classes.mod1;
global using static Doyle_Addin.Genius.Classes.modDcFilters;
global using static Doyle_Addin.Genius.Classes.modDcFilters;
global using static Doyle_Addin.Genius.Classes.modGPUpdateAT;
global using static Doyle_Addin.Genius.Classes.Module1;
global using static Doyle_Addin.Genius.Classes.sql0;


// 
// Modules imported from aiVba02.ivb [2017-09-29]
// AssyComponentsCollector
// lib0
// libCastOb
// mod1
// modDcFilters
// modGPUpdateAT
// modMacros
// src0
// 
// Attempt to track changes/revisions in this module

// Mods to dcGeniusPropsPart [2017-10-02]
// Rearrange sheet metal properties collection process
// to check for purchased parts before anything else.
// Sheet metal details should not needed for purchased parts.
// Resolution:
// 1: Pre-assign BOMStructure to check variable
// 2: Rewrite later checks to use new variable
// 3: In Sheet Metal branch,
// move Sheet Metal prop collection
// inside normalBOM sub-branch

// 
// Debug.Print cnGnsDoyle().Execute("select Item, Specification5, Specification4, Specification6 from vgMfiItems where Specification1 = 'CONVEYOR' and Specification2 = 'FIELDLOADER' ;").GetString
// "Specification = '' and "

