class fmEmpty
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Name = "fmEmpty";
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_GlobalNameSpace = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Creatable = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_PredeclaredId = true;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Exposed = false;

    private Scripting.Dictionary Watchers;

    public event CloseRequestedEventHandler CloseRequested;

    public delegate void CloseRequestedEventHandler(int CloseMode);

    public fmEmpty Itself()
    {
        Itself = this;
    }

    public Variant Notify(object ob, Variant ky = Empty
    )
    {
        long dx;

        {
            var withBlock = Watchers;
            if (IsEmpty(ky))
            {
                dx = withBlock.Count;
                while (withBlock.Exists(dx))
                    dx = 1 + dx;
                Notify = Notify(ob, dx);
            }
            else if (withBlock.Exists(ky))
                Notify = Empty;
            else
            {
                withBlock.Add(ky, ob);
                Notify = ky;
            }
        }
    }

    public Variant NoMsgs(Variant nm)
    {
        {
            var withBlock = Watchers;
            if (withBlock.Exists(nm))
            {
                withBlock.Remove(nm);
                NoMsgs = nm;
            }
            else
                NoMsgs = Empty;
        }
    }

    private void UserForm_Initialize()
    {
        Watchers = new Scripting.Dictionary();
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        VbMsgBoxResult ck;

        Cancel = 1;
        if (Watchers.Count > 0)
            CloseRequested?.Invoke(CloseMode);
        else
        {
            ck = MsgBox(Join(Array("Review any selections", "and select Yes if ready.", "Otherwise, select No."
       ), Constants.vbNewLine), Constants.vbYesNo, "Close Form?"
       );
            if (ck == Constants.vbYes)
                this.Hide();
            else
            {
            }
        }
    }

    private void UserForm_Terminate()
    {
        Watchers.RemoveAll();
        Watchers = null;
    }
}