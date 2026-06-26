namespace DoyleAddin.Genius;

using System.Configuration;

/// <summary>
///     Provides the connection string for the DoyleDB database.
///     Reads from app.config first, falling back to a default value.
/// </summary>
internal static class ConnectionStringProvider
{
	private const string ConfigName = "DoyleDB";

	/// <summary>
	///     Gets the connection string for the DoyleDB database.
	///     Tries app.config connectionStrings section first,
	///     then falls back to the default embedded connection string.
	/// </summary>
	internal static string GetConnectionString()
	{
		var config = ConfigurationManager.ConnectionStrings[ConfigName]?.ConnectionString;
		if (!string.IsNullOrWhiteSpace(config))
		{
			Debug.WriteLine("ConnectionStringProvider: Using connection string from config.");
			return config;
		}

		Debug.WriteLine("ConnectionStringProvider: WARNING — No config found, using default connection string." +
		                " Configure by adding a connectionStrings section to the add-in config file.");
		return
			"Data Source=DOYLE-ERP02;Initial Catalog=DoyleDB;User ID=geniusreporting;Password=geniusreporting;";
	}
}