class jsonConverter
{
    // ============================================= '
    // Public Methods
    // ============================================= '

    // '
    // Convert JSON string to object (Dictionary/Collection)
    // 
    // @method ParseJson
    // @param {String} json_String
    // @return {Object} (Dictionary or Collection)
    // @throws 10001 - JSON parse error
    // '
    public object ParseJson(string JsonString)
    {
        long json_Index;
        json_Index = 1;

        // Remove vbCr, vbLf, and vbTab from json_String
        JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "");

        json_SkipSpaces(JsonString, ref json_Index);
        switch (VBA.Mid(JsonString, json_Index, 1))
        {
            case "{":
                {
                    ParseJson = json_ParseObject(JsonString, ref json_Index);
                    break;
                }

            case "[":
                {
                    ParseJson = json_ParseArray(JsonString, ref json_Index);
                    break;
                }

            default:
                {
                    // Error: Invalid JSON string
                    Information.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(JsonString, ref json_Index, "Expecting '{' or '['"));
                    break;
                }
        }
    }

    // '
    // Convert object (Dictionary/Collection/Array) to JSON
    // 
    // @method ConvertToJson
    // @param {Variant} JsonValue (Dictionary, Collection, or Array)
    // @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
    // @return {String}
    // '
    public string ConvertToJson(Variant JsonValue, Variant Whitespace = , long json_CurrentIndentation = 0)
    {
        string json_buffer;
        long json_BufferPosition;
        long json_BufferLength;
        long json_Index;
        long json_LBound;
        long json_UBound;
        bool json_IsFirstItem;
        long json_Index2D;
        long json_LBound2D;
        long json_UBound2D;
        bool json_IsFirstItem2D;
        Variant json_Key;
        Variant json_Value;
        string json_DateStr;
        string json_Converted;
        bool json_SkipItem;
        bool json_PrettyPrint;
        string json_Indentation;
        string json_InnerIndentation;

        json_LBound = -1;
        json_UBound = -1;
        json_IsFirstItem = true;
        json_LBound2D = -1;
        json_UBound2D = -1;
        json_IsFirstItem2D = true;
        json_PrettyPrint = !IsMissing(Whitespace);

        switch (VBA.VarType(JsonValue))
        {
            case object _ when VBA.vbNull:
                {
                    ConvertToJson = "null";
                    break;
                }

            case object _ when VBA.vbDate:
                {
                    // Date
                    json_DateStr = ConvertToIso(VBA.CDate(JsonValue));

                    ConvertToJson = "\"" + json_DateStr + "\"";
                    break;
                }

            case object _ when VBA.vbString:
                {
                    // String (or large number encoded as string)

                    /// NOTE[2021.08.04] -- Prep for modification to this section
                    /// 
                    /// This comment block, and whitespace immediately above and below,
                    /// are added TEMPORARILY to note ongoing work on String data handling.
                    /// They should be removed once modified code is verified correct.
                    /// 
                    /// This segment is presumed to handle conversion of VBA Strings
                    /// to JSON compatible form. However, strings containing double quotes
                    /// do NOT appear to be properly converted.
                    /// 
                    /// In order to prevent premature termination by double-quote characters,
                    /// it is necessary to "escape" these characters, using a preceding
                    /// backslash (\) character. This does not appear to happen, resulting
                    /// in the output of invalid JSON text when such Strings are encountered.
                    /// 
                    /// To address this issue, new code will be added in this Case section
                    /// to attempt to capture and modify strings containing double quotes.
                    /// That might take place here, or possibly in the json_Encode function,
                    /// called in the Else clause just below. Note that this revision will
                    /// almost certainly require processing existing backslashes, as well.
                    /// 
                    /// Any new code interposed in the following section is to be demarcated
                    /// with comments preceded with the tripled single quotes seen in this
                    /// comment section, and removed along with said comments when no longer
                    /// needed. All modifications retained for the final revision should,
                    /// however, retain accompanying comments documenting earch modification
                    /// and its purpose.
                    /// 

                    if (!JsonOptions.UseDoubleForLargeNumbers & json_StringIsLargeNumber(JsonValue))
                        ConvertToJson = JsonValue;
                    else
                    {
                        /// watch and stop for any string
                        /// containing one or more double quotes
                        if (InStr(1, JsonValue, "\"") > 0)
                            // Stop
                            Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing
                        /// end of watch section
                        ConvertToJson = "\"" + json_Encode(JsonValue) + "\"";
                    }

                    break;
                }

            case object _ when VBA.vbBoolean:
                {
                    if (JsonValue)
                        ConvertToJson = "true";
                    else
                        ConvertToJson = "false";
                    break;
                }

            case object _ when VBA.vbArray <= VBA.VarType(JsonValue) && VBA.VarType(JsonValue) <= VBA.vbArray + VBA.vbByte:
                {
                    if (json_PrettyPrint)
                    {
                        if (VBA.VarType(Whitespace) == VBA.vbString)
                        {
                            json_Indentation = VBA.String(json_CurrentIndentation + 1, Whitespace);
                            json_InnerIndentation = VBA.String(json_CurrentIndentation + 2, Whitespace);
                        }
                        else
                        {
                            json_Indentation = VBA.Space((json_CurrentIndentation + 1) * Whitespace);
                            json_InnerIndentation = VBA.Space((json_CurrentIndentation + 2) * Whitespace);
                        }
                    }

                    // Array
                    json_BufferAppend(ref json_buffer, ref "[", ref json_BufferPosition, ref json_BufferLength);
                    json_LBound = LBound(JsonValue, 1);
                    json_UBound = UBound(JsonValue, 1);
                    json_LBound2D = LBound(JsonValue, 2);
                    json_UBound2D = UBound(JsonValue, 2);

                    if (json_LBound >= 0 & json_UBound >= 0)
                    {
                        for (json_Index = json_LBound; json_Index <= json_UBound; json_Index++)
                        {
                            if (json_IsFirstItem)
                                json_IsFirstItem = false;
                            else
                                // Append comma to previous line
                                json_BufferAppend(ref json_buffer, ref ",", ref json_BufferPosition, ref json_BufferLength);

                            if (json_LBound2D >= 0 & json_UBound2D >= 0)
                            {
                                // 2D Array
                                if (json_PrettyPrint)
                                    json_BufferAppend(ref json_buffer, ref Constants.vbNewLine, ref json_BufferPosition, ref json_BufferLength);
                                json_BufferAppend(ref json_buffer, ref json_Indentation + "[", ref json_BufferPosition, ref json_BufferLength);
                                for (json_Index2D = json_LBound2D; json_Index2D <= json_UBound2D; json_Index2D++)
                                {
                                    if (json_IsFirstItem2D)
                                        json_IsFirstItem2D = false;
                                    else
                                        json_BufferAppend(ref json_buffer, ref ",", ref json_BufferPosition, ref json_BufferLength);

                                    json_Converted = ConvertToJson(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation + 2);

                                    // For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                                    if (json_Converted == "")
                                    {
                                        // (nest to only check if converted = "")
                                        if (json_IsUndefined(JsonValue(json_Index, json_Index2D)))
                                            json_Converted = "null";
                                    }

                                    if (json_PrettyPrint)
                                        json_Converted = Constants.vbNewLine + json_InnerIndentation + json_Converted;

                                    json_BufferAppend(ref json_buffer, ref json_Converted, ref json_BufferPosition, ref json_BufferLength);
                                }

                                if (json_PrettyPrint)
                                    json_BufferAppend(ref json_buffer, ref Constants.vbNewLine, ref json_BufferPosition, ref json_BufferLength);

                                json_BufferAppend(ref json_buffer, ref json_Indentation + "]", ref json_BufferPosition, ref json_BufferLength);
                                json_IsFirstItem2D = true;
                            }
                            else
                            {
                                // 1D Array
                                json_Converted = ConvertToJson(JsonValue(json_Index), Whitespace, json_CurrentIndentation + 1);

                                // For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                                if (json_Converted == "")
                                {
                                    // (nest to only check if converted = "")
                                    if (json_IsUndefined(JsonValue(json_Index)))
                                        json_Converted = "null";
                                }

                                if (json_PrettyPrint)
                                    json_Converted = Constants.vbNewLine + json_Indentation + json_Converted;

                                json_BufferAppend(ref json_buffer, ref json_Converted, ref json_BufferPosition, ref json_BufferLength);
                            }
                        }
                    }



                    if (json_PrettyPrint)
                    {
                        json_BufferAppend(ref json_buffer, ref Constants.vbNewLine, ref json_BufferPosition, ref json_BufferLength);
                        if (VBA.VarType(Whitespace) == VBA.vbString)
                            json_Indentation = VBA.String(json_CurrentIndentation, Whitespace);
                        else
                            json_Indentation = VBA.Space(json_CurrentIndentation * Whitespace);
                    }

                    json_BufferAppend(ref json_buffer, ref json_Indentation + "]", ref json_BufferPosition, ref json_BufferLength);
                    ConvertToJson = json_BufferToString(ref json_buffer, json_BufferPosition, json_BufferLength);
                    break;
                }

            case object _ when VBA.vbObject:
                {
                    if (json_PrettyPrint)
                    {
                        if (VBA.VarType(Whitespace) == VBA.vbString)
                            json_Indentation = VBA.String(json_CurrentIndentation + 1, Whitespace);
                        else
                            json_Indentation = VBA.Space((json_CurrentIndentation + 1) * Whitespace);
                    }

                    // Dictionary
                    if (VBA.TypeName(JsonValue) == "Dictionary")
                    {
                        json_BufferAppend(ref json_buffer, ref "{", ref json_BufferPosition, ref json_BufferLength);
                        foreach (var json_Key in JsonValue.Keys)
                        {
                            // For Objects, undefined (Empty/Nothing) is not added to object
                            json_Converted = ConvertToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1);
                            if (json_Converted == "")
                                json_SkipItem = json_IsUndefined(JsonValue(json_Key));
                            else
                                json_SkipItem = false;

                            if (!json_SkipItem)
                            {
                                if (json_IsFirstItem)
                                    json_IsFirstItem = false;
                                else
                                    json_BufferAppend(ref json_buffer, ref ",", ref json_BufferPosition, ref json_BufferLength);

                                /// NOTE[2021.08.05]: Code to watch for Dictionary Keys
                                /// containing double quote marks, and check whether
                                /// those characters are properly escaped.
                                /// 
                                /// Initial run seems to indicate String values ARE
                                /// processed correctly in their Case section above.
                                /// 
                                /// Review of prior output seems to show incorrect
                                /// output only of Dictionary Keys with double quotes,
                                /// suggesting the issue might lie in THIS section.
                                /// 
                                /// Believe this might be it.
                                /// 
                                /// The original If statement below appears to be where
                                /// the JSON key/value expression is generated, and both
                                /// tracks appear to emit the key value UNPROCESSED,
                                /// when in fact, it SHOULD be processed recursively
                                /// through ConvertToJson, just like any (string) value.
                                /// 
                                /// There also appears an opportunity here to tighten up
                                /// the code a bit, generating the key/value expression
                                /// PRIOR to the PrettyPrint check, and then use that
                                /// to determine whether to prefix the indentation.
                                /// 
                                /// Will proceed with proposed revisions, and see how this goes.
                                /// 
                                /// The following code watches for such Keys,
                                /// and stops for review when encountered.
                                /// 
                                if (InStr(1, json_Key, "\"") > 0)
                                    Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing
                                /// 
                                // json_Converted = """" & ConvertToJson(json_Key, Whitespace, json_CurrentIndentation + 1) & """:" & json_Converted
                                if (IsNull(json_Key))
                                {
                                    json_Converted = "\"<NULL>\":" + json_Converted;
                                    System.Diagnostics.Debugger.Break(); // because we DON'T want this happening!
                                }
                                else
                                    json_Converted = "\"" + json_Encode(json_Key) + "\":" + json_Converted;
                                /// Preceding code is added to prepare key/value expression
                                /// unconditionally, as it's included in both cases.
                                /// 
                                if (json_PrettyPrint)
                                {
                                    // json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                                    /// Preceding is original expression,
                                    /// disabled in favor of the following,
                                    /// which will replace it, if successful
                                    /// 
                                    json_Converted = Constants.vbNewLine + json_Indentation + json_Converted;
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                }
                                else
                                {
                                }

                                json_BufferAppend(ref json_buffer, ref json_Converted, ref json_BufferPosition, ref json_BufferLength);
                            }
                        }

                        if (json_PrettyPrint)
                        {
                            json_BufferAppend(ref json_buffer, ref Constants.vbNewLine, ref json_BufferPosition, ref json_BufferLength);
                            if (VBA.VarType(Whitespace) == VBA.vbString)
                                json_Indentation = VBA.String(json_CurrentIndentation, Whitespace);
                            else
                                json_Indentation = VBA.Space(json_CurrentIndentation * Whitespace);
                        }

                        json_BufferAppend(ref json_buffer, ref json_Indentation + "}", ref json_BufferPosition, ref json_BufferLength);
                    }
                    else if (VBA.TypeName(JsonValue) == "Collection")
                    {
                        json_BufferAppend(ref json_buffer, ref "[", ref json_BufferPosition, ref json_BufferLength);
                        foreach (var json_Value in JsonValue)
                        {
                            if (json_IsFirstItem)
                                json_IsFirstItem = false;
                            else
                                json_BufferAppend(ref json_buffer, ref ",", ref json_BufferPosition, ref json_BufferLength);

                            json_Converted = ConvertToJson(json_Value, Whitespace, json_CurrentIndentation + 1);

                            // For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                            if (json_Converted == "")
                            {
                                // (nest to only check if converted = "")
                                if (json_IsUndefined(json_Value))
                                    json_Converted = "null";
                            }

                            if (json_PrettyPrint)
                                json_Converted = Constants.vbNewLine + json_Indentation + json_Converted;

                            json_BufferAppend(ref json_buffer, ref json_Converted, ref json_BufferPosition, ref json_BufferLength);
                        }

                        if (json_PrettyPrint)
                        {
                            json_BufferAppend(ref json_buffer, ref Constants.vbNewLine, ref json_BufferPosition, ref json_BufferLength);
                            if (VBA.VarType(Whitespace) == VBA.vbString)
                                json_Indentation = VBA.String(json_CurrentIndentation, Whitespace);
                            else
                                json_Indentation = VBA.Space(json_CurrentIndentation * Whitespace);
                        }

                        json_BufferAppend(ref json_buffer, ref json_Indentation + "]", ref json_BufferPosition, ref json_BufferLength);
                    }
                    else
                    {
                        if (JsonValue == null)
                            json_Value = "<Nothing>";
                        else
                            json_Value = "<Unhandled Object: " + VBA.TypeName(JsonValue) + ">";
                        json_Converted = ConvertToJson(json_Value, Whitespace, json_CurrentIndentation + 1);
                        json_BufferAppend(ref json_buffer, ref json_Converted, ref json_BufferPosition, ref json_BufferLength);
                    }

                    ConvertToJson = json_BufferToString(ref json_buffer, json_BufferPosition, json_BufferLength);
                    break;
                }

            case object _ when VBA.vbInteger:
            case object _ when VBA.vbLong:
            case object _ when VBA.vbSingle:
            case object _ when VBA.vbDouble:
            case object _ when VBA.vbCurrency:
            case object _ when VBA.vbDecimal:
                {
                    // Number (use decimals for numbers)
                    ConvertToJson = VBA.Replace(JsonValue, ",", ".");
                    break;
                }

            default:
                {
                    // vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
                    // Use VBA's built-in to-string

                    ConvertToJson = JsonValue;
                    break;
                }
        }
    }

    // ============================================= '
    // Private Functions
    // ============================================= '

    private Dictionary json_ParseObject(string json_String, ref long json_Index)
    {
        string json_Key;
        string json_NextChar;

        json_ParseObject = new Dictionary();
        json_SkipSpaces(json_String, ref json_Index);
        if (VBA.Mid(json_String, json_Index, 1) != "{")
            Information.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, ref json_Index, "Expecting '{'"));
        else
        {
            json_Index = json_Index + 1;

            do
            {
                json_SkipSpaces(json_String, ref json_Index);
                if (VBA.Mid(json_String, json_Index, 1) == "}")
                {
                    json_Index = json_Index + 1;
                    return;
                }
                else if (VBA.Mid(json_String, json_Index, 1) == ",")
                {
                    json_Index = json_Index + 1;
                    json_SkipSpaces(json_String, ref json_Index);
                }

                json_Key = json_ParseKey(json_String, ref json_Index);
                json_NextChar = json_Peek(json_String, json_Index);
                if (json_NextChar == "[" | json_NextChar == "{")
                    json_ParseObject.Item(json_Key) = json_ParseValue(json_String, ref json_Index);
                else
                    json_ParseObject.Item(json_Key) = json_ParseValue(json_String, ref json_Index);
            }
            while (true);
        }
    }

    private Collection json_ParseArray(string json_String, ref long json_Index)
    {
        json_ParseArray = new Collection();

        json_SkipSpaces(json_String, ref json_Index);
        if (VBA.Mid(json_String, json_Index, 1) != "[")
            Information.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, ref json_Index, "Expecting '['"));
        else
        {
            json_Index = json_Index + 1;

            do
            {
                json_SkipSpaces(json_String, ref json_Index);
                if (VBA.Mid(json_String, json_Index, 1) == "]")
                {
                    json_Index = json_Index + 1;
                    return;
                }
                else if (VBA.Mid(json_String, json_Index, 1) == ",")
                {
                    json_Index = json_Index + 1;
                    json_SkipSpaces(json_String, ref json_Index);
                }

                json_ParseArray.Add(json_ParseValue(json_String, ref json_Index));
            }
            while (true);
        }
    }

    private Variant json_ParseValue(string json_String, ref long json_Index)
    {
        json_SkipSpaces(json_String, ref json_Index);
        switch (VBA.Mid(json_String, json_Index, 1))
        {
            case "{":
                {
                    json_ParseValue = json_ParseObject(json_String, ref json_Index);
                    break;
                }

            case "[":
                {
                    json_ParseValue = json_ParseArray(json_String, ref json_Index);
                    break;
                }

            case "\"":
            case "'":
                {
                    json_ParseValue = json_ParseString(json_String, ref json_Index);
                    break;
                }

            default:
                {
                    if (VBA.Mid(json_String, json_Index, 4) == "true")
                    {
                        json_ParseValue = true;
                        json_Index = json_Index + 4;
                    }
                    else if (VBA.Mid(json_String, json_Index, 5) == "false")
                    {
                        json_ParseValue = false;
                        json_Index = json_Index + 5;
                    }
                    else if (VBA.Mid(json_String, json_Index, 4) == "null")
                    {
                        json_ParseValue = Null;
                        json_Index = json_Index + 4;
                    }
                    else if (VBA.InStr("+-0123456789", VBA.Mid(json_String, json_Index, 1)))
                        json_ParseValue = json_ParseNumber(json_String, ref json_Index);
                    else
                        Information.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, ref json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['"));
                    break;
                }
        }
    }

    private string json_ParseString(string json_String, ref long json_Index)
    {
        string json_Quote;
        string json_Char;
        string json_Code;
        string json_buffer;
        long json_BufferPosition;
        long json_BufferLength;

        json_SkipSpaces(json_String, ref json_Index);

        // Store opening quote to look for matching closing quote
        json_Quote = VBA.Mid(json_String, json_Index, 1);
        json_Index = json_Index + 1;

        while (json_Index > 0 & json_Index <= Strings.Len(json_String))
        {
            json_Char = VBA.Mid(json_String, json_Index, 1);

            switch (json_Char)
            {
                case @"\":
                    {
                        // Escaped string, \\, or \/
                        json_Index = json_Index + 1;
                        json_Char = VBA.Mid(json_String, json_Index, 1);

                        switch (json_Char)
                        {
                            case "\"":
                            case @"\":
                            case "/":
                            case "'":
                                {
                                    json_BufferAppend(ref json_buffer, ref json_Char, ref json_BufferPosition, ref json_BufferLength);
                                    json_Index = json_Index + 1;
                                    break;
                                }

                            case "b":
                                {
                                    json_BufferAppend(ref json_buffer, ref Constants.vbBack, ref json_BufferPosition, ref json_BufferLength);
                                    json_Index = json_Index + 1;
                                    break;
                                }

                            case "f":
                                {
                                    json_BufferAppend(ref json_buffer, ref Constants.vbFormFeed, ref json_BufferPosition, ref json_BufferLength);
                                    json_Index = json_Index + 1;
                                    break;
                                }

                            case "n":
                                {
                                    json_BufferAppend(ref json_buffer, ref Constants.vbCrLf, ref json_BufferPosition, ref json_BufferLength);
                                    json_Index = json_Index + 1;
                                    break;
                                }

                            case "r":
                                {
                                    json_BufferAppend(ref json_buffer, ref Constants.vbCr, ref json_BufferPosition, ref json_BufferLength);
                                    json_Index = json_Index + 1;
                                    break;
                                }

                            case "t":
                                {
                                    json_BufferAppend(ref json_buffer, ref Constants.vbTab, ref json_BufferPosition, ref json_BufferLength);
                                    json_Index = json_Index + 1;
                                    break;
                                }

                            case "u":
                                {
                                    // Unicode character escape (e.g. \u00a9 = Copyright)
                                    json_Index = json_Index + 1;
                                    json_Code = VBA.Mid(json_String, json_Index, 4);
                                    json_BufferAppend(ref json_buffer, ref VBA.ChrW(VBA.Val("&h" + json_Code)), ref json_BufferPosition, ref json_BufferLength);
                                    json_Index = json_Index + 4;
                                    break;
                                }
                        }

                        break;
                    }

                case object _ when json_Quote:
                    {
                        json_ParseString = json_BufferToString(ref json_buffer, json_BufferPosition, json_BufferLength);
                        json_Index = json_Index + 1;
                        return;
                    }

                default:
                    {
                        json_BufferAppend(ref json_buffer, ref json_Char, ref json_BufferPosition, ref json_BufferLength);
                        json_Index = json_Index + 1;
                        break;
                    }
            }
        }
    }

    private Variant json_ParseNumber(string json_String, ref long json_Index)
    {
        string json_Char;
        string json_Value;
        bool json_IsLargeNumber;

        json_SkipSpaces(json_String, ref json_Index);
        while (json_Index > 0 & json_Index <= Strings.Len(json_String))
        {
            json_Char = VBA.Mid(json_String, json_Index, 1);

            if (VBA.InStr("+-0123456789.eE", json_Char))
            {
                // Unlikely to have massive number, so use simple append rather than buffer here
                json_Value = json_Value + json_Char;
                json_Index = json_Index + 1;
            }
            else
            {
                // Excel only stores 15 significant digits, so any numbers larger than that are truncated
                // This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
                // See: http://support.microsoft.com/kb/269370
                // 
                // Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
                // (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
                json_IsLargeNumber = IIf(InStr(json_Value, "."), Strings.Len(json_Value) >= 17, Strings.Len(json_Value) >= 16);
                if (!JsonOptions.UseDoubleForLargeNumbers & json_IsLargeNumber)
                    json_ParseNumber = json_Value;
                else
                    // VBA.Val does not use regional settings, so guard for comma is not needed
                    json_ParseNumber = VBA.Val(json_Value);
                return;
            }
        }
    }

    private string json_ParseKey(string json_String, ref long json_Index)
    {
        // Parse key with single or double quotes
        if (VBA.Mid(json_String, json_Index, 1) == "\"" | VBA.Mid(json_String, json_Index, 1) == "'")
            json_ParseKey = json_ParseString(json_String, ref json_Index);
        else if (JsonOptions.AllowUnquotedKeys)
        {
            string json_Char;
            while (json_Index > 0 & json_Index <= Strings.Len(json_String))
            {
                json_Char = VBA.Mid(json_String, json_Index, 1);
                if ((json_Char != " ") & (json_Char != ":"))
                {
                    json_ParseKey = json_ParseKey + json_Char;
                    json_Index = json_Index + 1;
                }
                else
                    break;
            }
        }
        else
            Information.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, ref json_Index, "Expecting '\"' or '''"));

        // Check for colon and skip if present or throw if not present
        json_SkipSpaces(json_String, ref json_Index);
        if (VBA.Mid(json_String, json_Index, 1) != ":")
            Information.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, ref json_Index, "Expecting ':'"));
        else
            json_Index = json_Index + 1;
    }

    private bool json_IsUndefined(Variant json_Value)
    {
        // Empty / Nothing -> undefined
        switch (VBA.VarType(json_Value))
        {
            case object _ when VBA.vbEmpty:
                {
                    json_IsUndefined = true;
                    break;
                }

            case object _ when VBA.vbObject:
                {
                    switch (VBA.TypeName(json_Value))
                    {
                        case "Empty":
                        case "Nothing":
                            {
                                json_IsUndefined = true;
                                break;
                            }
                    }

                    break;
                }
        }
    }

    private string json_Encode(Variant json_Text)
    {
        // Reference: http://www.ietf.org/rfc/rfc4627.txt
        // Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
        long json_Index;
        string json_Char;
        long json_AscCode;
        string json_buffer;
        long json_BufferPosition;
        long json_BufferLength;

        for (json_Index = 1; json_Index <= VBA.Len(json_Text); json_Index++)
        {
            json_Char = VBA.Mid(json_Text, json_Index, 1);
            json_AscCode = VBA.AscW(json_Char);

            // When AscW returns a negative number, it returns the twos complement form of that number.
            // To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
            // https://support.microsoft.com/en-us/kb/272138
            if (json_AscCode < 0)
                json_AscCode = json_AscCode + 65536;

            // From spec, ", \, and control characters must be escaped (solidus is optional)

            switch (json_AscCode)
            {
                case 34:
                    {
                        // " -> 34 -> \"
                        json_Char = @"\""";
                        break;
                    }

                case 92:
                    {
                        // \ -> 92 -> \\
                        json_Char = @"\\";
                        break;
                    }

                case 47:
                    {
                        // / -> 47 -> \/ (optional)
                        if (JsonOptions.EscapeSolidus)
                            json_Char = @"\/";
                        break;
                    }

                case 8:
                    {
                        // backspace -> 8 -> \b
                        json_Char = @"\b";
                        break;
                    }

                case 12:
                    {
                        // form feed -> 12 -> \f
                        json_Char = @"\f";
                        break;
                    }

                case 10:
                    {
                        // line feed -> 10 -> \n
                        json_Char = @"\n";
                        break;
                    }

                case 13:
                    {
                        // carriage return -> 13 -> \r
                        json_Char = @"\r";
                        break;
                    }

                case 9:
                    {
                        // tab -> 9 -> \t
                        json_Char = @"\t";
                        break;
                    }

                case object _ when 0 <= json_AscCode && json_AscCode <= 31:
                case object _ when 127 <= json_AscCode && json_AscCode <= 65535:
                    {
                        // Non-ascii characters -> convert to 4-digit hex
                        json_Char = @"\u" + VBA.Right("0000" + VBA.Hex(json_AscCode), 4);
                        break;
                    }
            }

            json_BufferAppend(ref json_buffer, ref json_Char, ref json_BufferPosition, ref json_BufferLength);
        }

        json_Encode = json_BufferToString(ref json_buffer, json_BufferPosition, json_BufferLength);
    }

    private string json_Peek(string json_String, long json_Index, long json_NumberOfCharacters = 1)
    {
        // "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
        json_SkipSpaces(json_String, ref json_Index);
        json_Peek = VBA.Mid(json_String, json_Index, json_NumberOfCharacters);
    }

    private void json_SkipSpaces(string json_String, ref long json_Index)
    {
        // Increment index to skip over spaces
        while (json_Index > 0 & json_Index <= VBA.Len(json_String) & VBA.Mid(json_String, json_Index, 1) == " ")
            json_Index = json_Index + 1;
    }

    private bool json_StringIsLargeNumber(Variant json_String)
    {
        // Check if the given string is considered a "large number"
        // (See json_ParseNumber)

        long json_Length;
        long json_CharIndex;
        json_Length = VBA.Len(json_String);

        // Length with be at least 16 characters and assume will be less than 100 characters
        if (json_Length >= 16 & json_Length <= 100)
        {
            string json_CharCode;
            long json_Index;

            json_StringIsLargeNumber = true;

            for (json_CharIndex = 1; json_CharIndex <= json_Length; json_CharIndex++)
            {
                json_CharCode = VBA.Asc(VBA.Mid(json_String, json_CharIndex, 1));
                switch (json_CharCode)
                {
                    case 46:
                    case object _ when 48 <= json_CharCode && json_CharCode <= 57:
                    case 69:
                    case 101:
                        {
                            break;
                        }

                    default:
                        {
                            json_StringIsLargeNumber = false;
                            return;
                        }
                }
            }
        }
    }

    private void json_ParseErrorMessage(string json_String, ref long json_Index, string ErrorMessage)
    {
        // Provide detailed parse error message, including details of where and what occurred
        // 
        // Example:
        // Error parsing JSON:
        // {"abcde":True}
        // ^
        // Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

        long json_StartIndex;
        long json_StopIndex;

        // Include 10 characters before and after error (if possible)
        json_StartIndex = json_Index - 10;
        json_StopIndex = json_Index + 10;
        if (json_StartIndex <= 0)
            json_StartIndex = 1;
        if (json_StopIndex > VBA.Len(json_String))
            json_StopIndex = VBA.Len(json_String);

        json_ParseErrorMessage = "Error parsing JSON:" + VBA.vbNewLine + VBA.Mid(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) + VBA.vbNewLine + VBA.Space(json_Index - json_StartIndex) + "^" + VBA.vbNewLine + ErrorMessage;
    }

    private void json_BufferAppend(ref string json_buffer, ref Variant json_Append, ref long json_BufferPosition, ref long json_BufferLength)
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */    // VBA can be slow to append strings due to allocating a new string for each append
                                                                                                                                                // Instead of using the traditional append, allocate a large empty string and then copy string at append position
                                                                                                                                                // 
                                                                                                                                                // Example:
                                                                                                                                                // Buffer: "abc  "
                                                                                                                                                // Append: "def"
                                                                                                                                                // Buffer Position: 3
                                                                                                                                                // Buffer Length: 5
                                                                                                                                                // 
                                                                                                                                                // Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
                                                                                                                                                // Buffer: "abc       "
                                                                                                                                                // Buffer Length: 10
                                                                                                                                                // 
                                                                                                                                                // Copy memory for "def" into buffer at position 3 (0-based)
                                                                                                                                                // Buffer: "abcdef    "
                                                                                                                                                // 
                                                                                                                                                // Approach based on cStringBuilder from vbAccelerator
                                                                                                                                                // http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

        long json_AppendLength;
        long json_LengthPlusPosition;

        json_AppendLength = VBA.LenB(json_Append);
        json_LengthPlusPosition = json_AppendLength + json_BufferPosition;

        if (json_LengthPlusPosition > json_BufferLength)
        {
            // Appending would overflow buffer, add chunks until buffer is long enough
            long json_TemporaryLength;

            json_TemporaryLength = json_BufferLength;
            while (json_TemporaryLength < json_LengthPlusPosition)
            {
                // Initially, initialize string with 255 characters,
                // then add large chunks (8192) after that
                // 
                // Size: # Characters x 2 bytes / character
                if (json_TemporaryLength == 0)
                    json_TemporaryLength = json_TemporaryLength + 510;
                else
                    json_TemporaryLength = json_TemporaryLength + 16384;
            }

            json_buffer = json_buffer + VBA.Space((json_TemporaryLength - json_BufferLength) / 2);
            json_BufferLength = json_TemporaryLength;
        }

        // Copy memory from append to buffer at buffer position
        json_CopyMemory();
        json_BufferPosition = json_BufferPosition + json_AppendLength;
    }

    private string json_BufferToString(ref string json_buffer, long json_BufferPosition, long json_BufferLength)
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
        if (json_BufferPosition > 0)
            json_BufferToString = VBA.Left(json_buffer, json_BufferPosition / 2);
    }

    /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
    private long json_UnsignedAdd(long json_Start, long json_Increment)
    {
        /* TODO ERROR: Skipped EndIfDirectiveTrivia */
        if (json_Start & 0x80000000)
            json_UnsignedAdd = json_Start + json_Increment;
        else if ((json_Start | 0x80000000) < -json_Increment)
            json_UnsignedAdd = json_Start + json_Increment;
        else
            json_UnsignedAdd = (json_Start + 0x80000000) + (json_Increment + 0x80000000);
    }

    // '
    // VBA-UTC v1.0.3
    // (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
    // 
    // UTC/ISO 8601 Converter for VBA
    // 
    // Errors:
    // 10011 - UTC parsing error
    // 10012 - UTC conversion error
    // 10013 - ISO 8601 parsing error
    // 10014 - ISO 8601 conversion error
    // 
    // @module UtcConverter
    // @author tim.hall.engr@gmail.com
    // @license MIT (http://www.opensource.org/licenses/mit-license.php)
    // ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

    // (Declarations moved to top)

    // ============================================= '
    // Public Methods
    // ============================================= '

    // '
    // Parse UTC date to local date
    // 
    // @method ParseUtc
    // @param {Date} UtcDate
    // @return {Date} Local date
    // @throws 10011 - UTC parsing error
    // '
    public DateTime ParseUtc(DateTime utc_UtcDate)
    {


        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
        utc_TIME_ZONE_INFORMATION utc_TimeZoneInfo;
        utc_SYSTEMTIME utc_LocalDate;

        utc_GetTimeZoneInformation(utc_TimeZoneInfo);
        utc_SystemTimeToTzSpecificLocalTime(utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate);
        ParseUtc = utc_SystemTimeToDate(utc_LocalDate);
        /* TODO ERROR: Skipped EndIfDirectiveTrivia */
        return;

    utc_ErrorHandling:
        ;
        Information.Err.Raise(10011, "UtcConverter.ParseUtc", "UTC parsing error: " + Information.Err.Number + " - " + Information.Err.Description);
    }

    // '
    // Convert local date to UTC date
    // 
    // @method ConvertToUrc
    // @param {Date} utc_LocalDate
    // @return {Date} UTC date
    // @throws 10012 - UTC conversion error
    // '
    public DateTime ConvertToUtc(DateTime utc_LocalDate)
    {


        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
        utc_TIME_ZONE_INFORMATION utc_TimeZoneInfo;
        utc_SYSTEMTIME utc_UtcDate;

        utc_GetTimeZoneInformation(utc_TimeZoneInfo);
        utc_TzSpecificLocalTimeToSystemTime(utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate);
        ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate);
        /* TODO ERROR: Skipped EndIfDirectiveTrivia */
        return;

    utc_ErrorHandling:
        ;
        Information.Err.Raise(10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " + Information.Err.Number + " - " + Information.Err.Description);
    }

    // '
    // Parse ISO 8601 date string to local date
    // 
    // @method ParseIso
    // @param {Date} utc_IsoString
    // @return {Date} Local date
    // @throws 10013 - ISO 8601 parsing error
    // '
    public DateTime ParseIso(string utc_IsoString)
    {
        string[] utc_Parts;
        string[] utc_DateParts;
        string[] utc_TimeParts;
        long utc_OffsetIndex;
        bool utc_HasOffset;
        bool utc_NegativeOffset;
        string[] utc_OffsetParts;
        DateTime utc_Offset;

        utc_Parts = VBA.Split(utc_IsoString, "T");
        utc_DateParts = VBA.Split(utc_Parts[0], "-");
        ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts[0]), VBA.CInt(utc_DateParts[1]), VBA.CInt(utc_DateParts[2]));

        if (UBound(utc_Parts) > 0)
        {
            if (VBA.InStr(utc_Parts[1], "Z"))
                utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts[1], "Z", ""), ":");
            else
            {
                utc_OffsetIndex = VBA.InStr(1, utc_Parts[1], "+");
                if (utc_OffsetIndex == 0)
                {
                    utc_NegativeOffset = true;
                    utc_OffsetIndex = VBA.InStr(1, utc_Parts[1], "-");
                }

                if (utc_OffsetIndex > 0)
                {
                    utc_HasOffset = true;
                    utc_TimeParts = VBA.Split(VBA.Left(utc_Parts[1], utc_OffsetIndex - 1), ":");
                    utc_OffsetParts = VBA.Split(VBA.Right(utc_Parts[1], Strings.Len(utc_Parts[1]) - utc_OffsetIndex), ":");

                    switch (UBound(utc_OffsetParts))
                    {
                        case 0:
                            {
                                utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts[0]), 0, 0);
                                break;
                            }

                        case 1:
                            {
                                utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts[0]), VBA.CInt(utc_OffsetParts[1]), 0);
                                break;
                            }

                        case 2:
                            {
                                // VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                                utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts[0]), VBA.CInt(utc_OffsetParts[1]), Int(VBA.Val(utc_OffsetParts[2])));
                                break;
                            }
                    }

                    if (utc_NegativeOffset)
                        utc_Offset = -utc_Offset;
                    else
                        utc_TimeParts = VBA.Split(utc_Parts[1], ":");
                }

                switch (UBound(utc_TimeParts))
                {
                    case 0:
                        {
                            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts[0]), 0, 0);
                            break;
                        }

                    case 1:
                        {
                            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts[0]), VBA.CInt(utc_TimeParts[1]), 0);
                            break;
                        }

                    case 2:
                        {
                            // VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts[0]), VBA.CInt(utc_TimeParts[1]), Int(VBA.Val(utc_TimeParts[2])));
                            break;
                        }
                }

                ParseIso = ParseUtc(ParseIso);

                if (utc_HasOffset)
                    ParseIso = ParseIso + utc_Offset;
            }

            return;

        utc_ErrorHandling:
            ;
            Information.Err.Raise(10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " + utc_IsoString + ": " + Information.Err.Number + " - " + Information.Err.Description);
        }
    }

    // '
    // Convert local date to ISO 8601 string
    // 
    // @method ConvertToIso
    // @param {Date} utc_LocalDate
    // @return {Date} ISO 8601 string
    // @throws 10014 - ISO 8601 conversion error
    // '
    public string ConvertToIso(DateTime utc_LocalDate)
    {
        ConvertToIso = VBA.Format(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z");

        return;

    utc_ErrorHandling:
        ;
        Information.Err.Raise(10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " + Information.Err.Number + " - " + Information.Err.Description);
    }

    // ============================================= '
    // Private Functions
    // ============================================= '

    /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
    private utc_SYSTEMTIME utc_DateToSystemTime(DateTime utc_Value)
    {
        utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value);
        utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value);
        utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value);
        utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value);
        utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value);
        utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value);
        utc_DateToSystemTime.utc_wMilliseconds = 0;
    }

    private DateTime utc_SystemTimeToDate(utc_SYSTEMTIME utc_Value)
    {
        utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond);
    }
}