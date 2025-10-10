

// 
// Modules imported from aiVba02.ivb [2017-09-29]
//     AssyComponentsCollector
//     lib0
//     libCastOb
//     mod1
//     modDcFilters
//     modGPUpdateAT
//     modMacros
//     src0
// 
// Attempt to track changes/revisions in this module

// Mods to dcGeniusPropsPart [2017-10-02]
// Rearrange sheet metal properties collection process
// to check for purchased parts before anything else.
// Sheet metal details should not needed for purchased parts.
// Resolution:
//     1: Pre-assign BOMStructure to check variable
//     2: Rewrite later checks to use new variable
//     3: In Sheet Metal branch,
//         move Sheet Metal prop collection
//         inside normalBOM sub-branch

// 
// Debug.Print cnGnsDoyle().Execute("select Item, Specification5, Specification4, Specification6 from vgMfiItems where Specification1 = 'CONVEYOR' and Specification2 = 'FIELDLOADER' ;").GetString
// "Specification = '' and "

