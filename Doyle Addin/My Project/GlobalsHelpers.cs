#region

global using System;
global using System.Collections.Generic;
global using System.Diagnostics;
global using System.Drawing;
global using System.Drawing.Drawing2D;
global using System.Drawing.Imaging;
global using System.IO;
global using System.Linq;
global using System.Reflection;
global using System.Runtime.InteropServices;
global using System.Text.Json;
global using System.Windows.Forms;
global using System.Xml.Serialization;
global using Docnet.Core;
global using Docnet.Core.Models;
global using Inventor;
global using Svg;
global using static Doyle_Addin.My_Project.GlobalsHelpers;
global using static Inventor.DocumentTypeEnum;
global using Application = Inventor.Application;
global using Color = System.Drawing.Color;
global using Document = Inventor.Document;
global using Environment = System.Environment;
global using File = System.IO.File;
global using IPictureDisp = Inventor.IPictureDisp;
global using Path = System.IO.Path;

#endregion

namespace Doyle_Addin.My_Project;

internal static class GlobalsHelpers
{
    // Inventor application object.
    public static Application ThisApplication;
    // public static Document ThisDocument;
}