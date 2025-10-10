class SurroundingClass
{
    private Scripting.Dictionary dcWkg;
    private Scripting.Dictionary dcFiled;
    private Scripting.Dictionary dcIndex;

    private fmIfcTest05A fm;

    private void Class_Initialize()
    {
        /// 
        dcWkg = new Scripting.Dictionary();
        dcFiled = new Scripting.Dictionary();
        dcIndex = new Scripting.Dictionary();

        fm = new fmIfcTest05A();
    }

    private void Class_Terminate()
    {
        /// 
        fm = null;

        dcWkg.RemoveAll();
        dcFiled.RemoveAll();
        dcIndex.RemoveAll();

        dcWkg = null;
        dcFiled = null;
        dcIndex = null;
    }

    public wkgCls0 Itself()
    {
        Itself = this;
    }

    public wkgCls0 Using(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (AiDoc == null)
        {
            using ( == this)
            {
                ;/* Cannot convert ElseStatementSyntax, CONVERSION ERROR: Conversion for ElseStatement not implemented, please report this issue in 'Else' at character 734
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitElseStatement(ElseStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.ElseStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Else

 */
                using ( == Collect(AiDoc))
                {
                }
            }
        }
    }

    public Scripting.Dictionary Process(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;

        if (AiDoc == null)
        {
            /// THIS is where we start processing
            /// Inventor Documents collected
            /// in the internal Dictionary

            /// at the moment, we simply pull up
            /// the standard form and present it
            /// with the current collection
            /// of Inventor Documents
            {
                var withBlock = fm.Using(dcWkg);
                withBlock.Show(1);
                System.Diagnostics.Debugger.Break();
            }

            rt = dcCopy(dcWkg);
        }
        else
            rt = Collect(AiDoc).Process();

        Process = rt;
    }

    public wkgCls0 Collect(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */) // Scripting.Dictionary
    {
        // ,optional dcWkg as Scripting.Dictionary=nothing
        /// Method Function Collect
        /// 
        /// given a valid Inventor Document
        /// (usually an assembly), gather
        /// any and all Parts in it into
        /// the internal Dictionary dcWkg,
        /// and return a copy.
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        if (AiDoc == null)
            Collect(ThisApplication.ActiveDocument);
        else if (AiDoc == ThisDocument)
        {
        }
        else
            dcWkg = dcAiDocGrpsByForm(dcRemapByPtNum(dcAiDocComponents(AiDoc, null/* Conversion error: Set to default value for this argument */, 0)));

        Collect = this; // dcCopy(dcWkg)
    }
}