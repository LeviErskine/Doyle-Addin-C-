global using System;
global using System.Diagnostics;
global using System.IO;
global using System.Linq;
global using Inventor;
global using static System.Environment;
global using static System.IO.File;
global using static System.IO.Path;
global using static DoyleAddin.GlobalUsings;
global using static Inventor.DocumentTypeEnum;
global using Application = Inventor.Application;
global using Environment = System.Environment;
global using File = System.IO.File;
global using Path = System.IO.Path;

namespace DoyleAddin;

internal static class GlobalUsings
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