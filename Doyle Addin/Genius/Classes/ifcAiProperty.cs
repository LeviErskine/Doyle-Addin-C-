class ifcAiProperty : ifcDatum
{

    // Private ps As Inventor.PropertySet
    private Inventor.Property pr;
    // Private nm As String
    private Variant vlWas;
    private Variant vlNow;

    public ifcDatum Connect(Inventor.Property ToProp)
    {
        pr = ToProp;
        if (ToProp == null)
            // ps = Nothing
            // nm = ""
            vlWas = Empty;
        else
        {
            var withBlock = pr;
            // ps = .Parent
            // nm = .Name
            vlWas = withBlock.Value;
        }
        vlNow = vlWas;

        Connect = this;
    }
    /// replaces disabled function below
    // 
    // Public Function AttachedTo(Name As String,'    Optional InPropSet As Inventor.PropertySet = Nothing') As ifcAiProperty
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
    // pr = ps.Item(nm)
    // If Err.Number = 0 Then
    // vlWas = pr.Value
    // Else
    // pr = Nothing
    // vlWas = Empty
    // End If
    // 
    // End If
    // 
    // AttachedTo = Me
    // End Function

    public ifcDatum MakeValue(Variant This)
    {
        if (IsObject(This))
        {
        }
        else
            vlNow = This;

        MakeValue = this;
    }
    /// replaces disabled function below
    // 
    // Public Function WithValue('    NewVal As Variant') As ifcAiProperty
    // Me.Value = NewVal
    // 
    // WithValue = Me
    // End Function

    public ifcAiProperty Commit()
    {
        // Dim ps As Inventor.PropertySet
        // Dim ck As Variant

        if (IsEmpty(vlWas))
            System.Diagnostics.Debugger.Break();
        else if (pr == null)
        {
        }
        else
        {
            vlWas = pr.Value; // ck

            if (vlNow == vlWas)
            {
            }
            else if (System.Convert.ToHexString(vlNow) == System.Convert.ToHexString(vlWas))
            {
            }
            else
            {
                pr.Value = vlNow; // vlWas
                if (Information.Err.Number == 0)
                {
                }
                else
                    System.Diagnostics.Debugger.Break();// and see what we can do
            }
        }

        Commit = this;
    }

    private ifcDatum Itself()
    {
        Itself = this;
    }
    /// replaces disabled function below
    // 
    // Public Function Obj() As ifcAiProperty
    // Obj = Me
    // End Function

    private bool Connected(Inventor.Property ToThis = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (ToThis == null)
            Connected = !pr == null;
        else
            Connected = pr == ToThis;
    }

    private Variant Value()
    {
        if (IsObject(vlNow))
            /// this should NOT ever happen
            /// but just to be robust...
            Value = vlNow;
        else
            Value = vlNow;
    }

    public long Status()
    {
        Status = -1;
    }

    public Variant Name()
    {
        if (pr == null)
            Name = pr.Name;
        else
            Name = "";
    }

    // Public Property Value() As Variant
    // Get
    // Value = vlWas
    // End Property
    // 
    // Public Property  Value(NewVal As Variant)
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
        vlWas = Empty;
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
            System.Diagnostics.Debugger.Break();
        else if (vlWas == pr.Value)
        {
        }
        else if (System.Convert.ToHexString(vlWas) == System.Convert.ToHexString(pr.Value))
        {
        }
        else
            // and MIGHT need to be committed
            System.Diagnostics.Debugger.Break();
    }

    private ifcDatum ifcDatum_Commit()
    {
    }

    private ifcDatum ifcDatum_Connect(object ToThis)
    {
        ifcDatum_Connect = Connect(obAiProp(ToThis));
    }

    private bool ifcDatum_Connected(object ToThis = null)
    {
        ifcDatum_Connected = Connected(obAiProp(ToThis));
    }

    private ifcDatum ifcDatum_Itself()
    {
        ifcDatum_Itself = this;
    }

    private ifcDatum ifcDatum_MakeValue(Variant This)
    {
        ifcDatum_MakeValue = MakeValue(This);
    }

    private Variant ifcDatum_Value()
    {
        if (IsObject(vlNow))
            /// this should NOT ever happen
            /// but just to be robust...
            ifcDatum_Value = vlNow;
        else
            ifcDatum_Value = vlNow;
    }
}