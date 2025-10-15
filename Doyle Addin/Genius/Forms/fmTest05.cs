using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Forms;

class fmTest05 : Form
{
    private var VB_Name = "fmTest05";

    private var VB_GlobalNameSpace = false;

    private var VB_Creatable = false;

    private var VB_PredeclaredId = true;

    private var VB_Exposed = false;

    public event SentEventHandler Sent;

    public delegate void SentEventHandler(VbMsgBoxResult Signal);

    public event GroupIsEventHandler GroupIs;

    public delegate void GroupIsEventHandler(string Now);

    public event ItemIsEventHandler ItemIs;

    public delegate void ItemIsEventHandler(string Now);

    private Dictionary dcHolding;

    private const string txVersion = "";
    // 

    // 

    public fmTest05 Holding(dynamic Obj
    )
    {
        // Holding -- Hold onto supplied
        // dynamic until terminated,
        // or directed to drop it.
        // 
        // not sure about this one.
        // purpose is to keep a
        // client interface "alive"
        // while the form itself
        // remains active.
        // 
        {
            var withBlock = dcHolding;
            if (withBlock.Exists(Obj))
            {
            }
            else
                withBlock.Add(Obj, withBlock.Count);
        }

        return this;
    }

    public fmTest05 Dropping(dynamic Obj
    )
    {
        {
            var withBlock = dcHolding;
            if (withBlock.Exists(Obj))
                withBlock.Remove(Obj);
            else
            {
            }
        }

        return this;
    }

    public string GroupNow()
    {
        {
            var withBlock = tbsItemGrps;
            long dx = withBlock.Value;
            MSForms.Tab tb = withBlock.Tabs.get_Item(dx);
            return tb.Name;
        }
    }

    public fmTest05 InGroup(string GrpId
    ) // fmIfcTest05A
    {
        {
            var withBlock = tbsItemGrps;
            Information.Err().Clear();

            MSForms.Tab tb = withBlock.Tabs.get_Item(GrpId);
            if (Information.Err().Number == 0)
                withBlock.Value = tb.Index;
            else
            {
            }

            Information.Err().Clear();
        }

        return this;
    }

    public string ItemNow()
    {
        // With lbxItems
        return lbxItems.Value;
    }

    public fmTest05 OnItem(string ItemId
    ) // fmIfcTest05A
    {
        // Dim tb As MSForms.Tab

        {
            var withBlock = lbxItems;

            Information.Err().Clear();

            // tb = .Tabs.get_Item(ItemId)
            withBlock.Value = ItemId;
            if (Information.Err().Number == 0)
                // .Value = tb.Index
                // Stop
                Debug.Print(""); // Breakpoint Landing
            else
                Debugger.Break();

            Information.Err().Clear();
        }

        return this;
    }

    private void cmdEndCancel_Click()
    {
        Sent?.Invoke(Constants.vbCancel);
    }

    private void cmdEndSave_Click()
    {
        Sent?.Invoke(Constants.vbOK);
    }

    private void cmdOpenItem_Click()
    {
        Sent?.Invoke(Constants.vbRetry);
    }

    private void lbxItems_Change()
    {
        ItemIs?.Invoke(lbxItems.Value);
    }

    private void tbsItemGrps_Change()
    {
        GroupIs?.Invoke(GroupNow);
    }

    private void tbsItemGrps_BeforeDropOrPaste(long Index, MSForms.ReturnBoolean Cancel, MSForms.fmAction Action,
        MSForms.DataObject Data, float X, float Y, MSForms.ReturnEffect Effect, int Shift
    )
    {
        // will keep this one as is, for now
        // not sure what you can actually drop
        // onto a tab group
        Debugger.Break();
    }

    private void lbxItems_MouseMove(int Button, int Shift, float X, float Y
    )
    {
        // keeping this one here, since it basically governs
        // drag-and-drop behavior from a local control.
        // might try to see if this is actually needed.
        // one would think this kind of behavior
        // would occur automatically.

        if (Button != 1) return;
        var dt = new DataObject();
        dt.SetText(lbxItems.Value);
        int ef = dt.StartDrag();
    }

    private void InitializeComponent()
    {
        dcHolding = new Dictionary();
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        Sent?.Invoke(Constants.vbAbort);
    }

    private void UserForm_Terminate()
    {
        dcHolding.RemoveAll();
        dcHolding = null;
    }
}