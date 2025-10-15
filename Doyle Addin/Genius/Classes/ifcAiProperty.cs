using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class ifcAiProperty : ifcDatum
{
    // Private ps As Inventor.PropertySet
    private Property pr;

    // Private nm As String
    private dynamic vlWas;
    private dynamic vlNow;

    public ifcDatum Connect(Property ToProp)
    {
        pr = ToProp;
        if (ToProp == null)
            // ps = Nothing
            // nm = ""
            vlWas = null;
        else
        {
            var withBlock = pr;
            // ps = .Parent
            // nm = .Name
            vlWas = withBlock.Value;
        }

        vlNow = vlWas;

        return this;
    }
    // replaces disabled function below
    // 
    // Public Function AttachedTo(Name As String,' Optional InPropSet As Inventor.PropertySet = Nothing') As ifcAiProperty
    // '''
    // '''
    // '''
    // nm = Name
    // 
    // If Not InPropSet Is Nothing Then
    // ps = InPropSet
    // End If
    // 
    // If Not ps Is Nothing Then
    // 
    // pr = ps.get_Item(nm)
    // If Err().Number = 0 Then
    // vlWas = pr.Value
    // Else
    // pr = Nothing
    // vlWas = null
    // End If
    // 
    // End If
    // 
    // AttachedTo = Me
    // End Function

    public ifcDatum MakeValue(dynamic This)
    {
        if (This is not null)
        {
        }
        else
            vlNow = null;

        return this;
    }
    // replaces disabled function below
    // 
    // Public Function WithValue(' NewVal As dynamic') As ifcAiProperty
    // Me.Value = NewVal
    // 
    // WithValue = Me
    // End Function

    public ifcAiProperty Commit()
    {
        // Dim ps As Inventor.PropertySet
        // Dim ck As dynamic

        if (vlWas is null)
            Debugger.Break();
        else if (pr == null)
        {
        }
        else
        {
            vlWas = pr.Value; // ck

            if (vlNow == vlWas)
            {
            }
            else if (Convert.ToHexString(vlNow) == Convert.ToHexString(vlWas))
            {
            }
            else
            {
                pr.Value = vlNow; // vlWas
                if (Information.Err().Number == 0)
                {
                }
                else
                    Debugger.Break(); // and see what we can do
            }
        }

        return this;
    }

    private ifcDatum Itself()
    {
        return this;
    }
    // replaces disabled function below
    // 
    // Public Function Obj() As ifcAiProperty
    // Obj = Me
    // End Function

    private bool Connected(Property ToThis = null)
    {
        if (ToThis == null)
            return !pr == null;
        return pr == ToThis;
    }

    private dynamic Value()
    {
        // this should NOT ever happen
        // but just to be robust...
        return vlNow ?? vlNow;
    }

    public long Status()
    {
        return -1;
    }

    public dynamic Name()
    {
        return pr == null ? pr.Name : "";
    }

    // Public Property Value() As dynamic
    // Get
    // Value = vlWas
    // End Property
    // 
    // Public Property Value(NewVal As dynamic)
    // If IsEmpty(NewVal) Then
    // Stop
    // 'ElseIf IsNull(NewVal) Then
    // 'ElseIf IsMissing(NewVal) Then
    // ElseIf IsObject(NewVal) Then
    // Stop
    // Else
    // vlWas = NewVal
    // End If
    // End Property

    private void Class_Initialize()
    {
        // nm = ""
        vlWas = null;
        // ps = Nothing
        pr = null;
    }

    private void Class_Terminate()
    {
        // If ps Is Nothing Then 'nowhere to save
        // so nothing to do but drop it
        // Else
        if (pr == null)
            // to create Property, if desired
            // (and possible)
            Debugger.Break();
        else if (vlWas == pr.Value)
        {
        }
        else if (Convert.ToHexString(vlWas) == Convert.ToHexString(pr.Value))
        {
        }
        else
            // and MIGHT need to be committed
            Debugger.Break();
    }

    private ifcDatum ifcDatum_Commit()
    {
    }

    private ifcDatum ifcDatum_Connect(dynamic ToThis)
    {
        return Connect(obAiProp(ToThis));
    }

    private bool ifcDatum_Connected(dynamic ToThis = null)
    {
        return Connected(obAiProp(ToThis));
    }

    private ifcDatum ifcDatum_Itself()
    {
        return this;
    }

    private ifcDatum ifcDatum_MakeValue(dynamic This)
    {
        return MakeValue(This);
    }

    private dynamic ifcDatum_Value()
    {
        return vlNow ?? vlNow;
    }
}