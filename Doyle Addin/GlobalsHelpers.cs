global using System;
global using System.IO;
global using static DoyleAddin.GlobalsHelpers;
global using static Inventor.DocumentTypeEnum;
global using Application = Inventor.Application;
global using Environment = System.Environment;
global using File = System.IO.File;
global using Path = System.IO.Path;

namespace DoyleAddin;

internal static class GlobalsHelpers
{
	// Inventor application object.
	internal static Application ThisApplication { get; set; }

	/// <summary>
	///     Initializes the ThisApplication static field with the Inventor application instance
	/// </summary>
	/// <param name="application">The Inventor application instance</param>
	internal static void Initialize(Application application)
	{
		ThisApplication = application;
	}
}