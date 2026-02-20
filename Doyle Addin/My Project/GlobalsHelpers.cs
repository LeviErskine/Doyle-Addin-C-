#region

global using static DoyleAddin.My_Project.GlobalsHelpers;
global using static Inventor.DocumentTypeEnum;
global using Application = Inventor.Application;
global using Color = System.Drawing.Color;
global using Document = Inventor.Document;
global using Environment = System.Environment;
global using File = System.IO.File;
global using IPictureDisp = Inventor.IPictureDisp;
global using Path = System.IO.Path;
using System;

#endregion

namespace DoyleAddin.My_Project;

internal static class GlobalsHelpers
{
	// Inventor application object.
	internal static Application ThisApplication { get; private set; }

	/// <summary>
	///     Initializes the ThisApplication static field with the Inventor application instance
	/// </summary>
	/// <param name="application">The Inventor application instance</param>
	internal static void Initialize(Application application)
	{
		ThisApplication = application ?? throw new ArgumentNullException(nameof(application));
	}
}