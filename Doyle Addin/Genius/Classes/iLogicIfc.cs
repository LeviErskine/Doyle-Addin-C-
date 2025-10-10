class ifLogicIfc
{
    private Inventor.Document md;
    private Inventor.Application ap;
    // Private pj As Inventor.DesignProject
    private Inventor.ApplicationAddIn ad;
    private object il; // Inventor.ApplicationAddIn
    private Scripting.Dictionary dc;

    // Private mp As Inventor.NameValueMap 'for WithArgs
    /// NOTE: this NameValueMap is not presently used

    /// as method WithArgs is not yet implemented,

    /// or even defined.

    /// 

    public Inventor.Document RuleSource()
    {
        RuleSource = md;
    }

    public iLogicIfc WithRulesIn(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// DO NOT CALL this Method with ThisDocument
        /// For some unknown reason, this will trigger
        /// a Type Mismatch Error (Err 13) on exit.
        /// 
        Variant rl;

        if (AiDoc == null)
            md = ThisDocument;
        else
            md = AiDoc;

        ap = md.Parent;
        {
            var withBlock = ap;
            {
                var withBlock1 = withBlock.ApplicationAddIns;
                ad = withBlock1.ItemById(guidILogicAdIn);
                {
                    var withBlock2 = ad;
                    if (!withBlock2.Activated)
                        withBlock2.Activate();
                    dc.RemoveAll();

                    if (withBlock2.Activated)
                    {
                        il = withBlock2.Automation;
                        foreach (var rl in il.rules(md))
                            dc.Add(rl.Name, rl);
                    }
                    else
                        il = null;
                }
            }
        }

        WithRulesIn = this;
    }

    public Variant RuleNames(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (AiDoc == null)
            RuleNames = dc.Keys;
        else
            RuleNames = WithRulesIn(AiDoc).RuleNames();
    }

    public Scripting.Dictionary RuleDefs(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Variant nm;

        rt = new Scripting.Dictionary();
        foreach (var nm in RuleNames(AiDoc))
            rt.Add(nm, TextOf(System.Convert.ToHexString(nm)));
        RuleDefs = rt;
    }

    public string TextOf(string ruleName)
    {
        /// TextOf - retrieve text of rule indicated
        /// NOTE[2023.03.06.1424] will need error traps
        /// 
        object rl;

        if (il == null)
            rl = null;
        else
        {
            Information.Err.Clear();
            rl = il.GetRule(md, ruleName);
            /// NOTE[2023.03.06.1511] might want
            /// to use Dictionary instead, and thus
            /// reduce anticipated need for error traps.
            /// NOTE, however, that a rule's name
            /// MIGHT be changed during run time,
            /// so the Dictionary approach
            /// might still fail.
            if (Information.Err.Number != 0)
            {
                System.Diagnostics.Debugger.Break();
                rl = null;
            }
        }

        if (rl == null)
            TextOf = "''' NORULE '''";
        else
            TextOf = rl.Text();
    }

    public Scripting.Dictionary Apply(string ruleName, object Args = null)
    {
        // Optional dcArgs As Scripting.Dictionary = Nothing,'Optional NVMap As Inventor.NameValueMap = Nothing,'
        /// NOTE: Some iLogic Rules might behave
        /// differently when supplied with
        /// arguments (in a NameValueMap)
        /// than without. For example, a
        /// Rule which would normally add
        /// its results to the supplied
        /// NameValueMap might instead
        /// present them in a message box.
        /// 
        /// If such behavior is not desired
        /// in a call with no arguments,
        /// an empty NameValueMap or
        /// Dictionary may be supplied
        /// to avoid it, provided the
        /// Rule supports this.
        /// 
        Inventor.NameValueMap mp;
        Scripting.Dictionary rt;

        if (il == null)
        {
            rt = new Scripting.Dictionary();
            System.Diagnostics.Debugger.Break(); // and debug
        }
        else if (Args == null)
        {
            il.RunRule(md, ruleName);
            rt = new Scripting.Dictionary();
        }
        else if (Args is Inventor.NameValueMap)
        {
            il.RunRuleWithArguments(md, ruleName, Args);         /// !!!WARNING!!! This call MIGHT result
                                                                 /// in changes to supplied NameValueMap,
                                                                 /// which might or might not be a problem
                                                                 /// for the client process.
                                                                 /// 
                                                                 /// for now, will keep this way,
                                                                 /// but might reconsider in future
                                                                 /// for safety or other reasons.
                                                                 /// 
                                                                 /// UPDATE[2023.03.06.1401]
                                                                 /// actually, this serves vital role
                                                                 /// in receiving results FROM iLogic rule,
                                                                 /// so this is almost certainly going
                                                                 /// to remain as is

            rt = dcFromAiNameValMap(Args); // mp
        }
        else if (Args is Scripting.Dictionary)
            rt = Apply(ruleName, dc2aiNameValMap(Args)); // dcArgs, NVMap
        else
        {
        }

        Apply = rt;
    }

    /// 

    /// 
    public iLogicIfc Itself()
    {
        Itself = this;
    }

    private void Class_Initialize()
    {
        /// REV[2023.03.06.1504] added Dictionary
        /// to collect set of iLogic Rules
        dc = new Scripting.Dictionary();
        /// Initialization calls local method WithRulesIn
        /// to set up private objects and variables.
        /// 
        /// Originally passed ThisDocument to initialize
        /// from Inventor Document containing this Class.
        /// 
        /// However, for some unknown reason, a Type Mismatch
        /// Error is triggered by the value returned by
        /// WithRulesIn, even though it SHOULD be compatible.
        /// 
        /// After numerous failed attempts to correct the issue,
        /// it was decided to pass Nothing to the function, which
        /// interprets the null value as an indication to use
        /// ThisDocument internally. THIS seems to work
        /// 
        /// One prior "solution" was to enclose the call in an
        /// Error Trap that clears the Type Mismatch, however,
        /// this felt too much a kludge, so this alternative
        /// was chosen instead.
        /// 
        // WithRulesIn ThisDocument
        // WithRulesIn Nothing
        WithRulesIn(ThisApplication.Documents.ItemByName(ThisDocument.FullDocumentName));
    }
}