using System;
using System.IO;
using System.Xml.Serialization;

namespace Doyle_Addin.Options;

/// <summary>
/// Represents a collection of user-configurable options and preferences related to
/// the Doyle Add-in. This class provides methods to save and load these options
/// as an XML file in the user's application data directory.
/// </summary>
public class UserOptions
{
    /// <summary>
    /// Specifies the directory path where PDF files will be exported during the print update process.
    /// </summary>
    public string PrintExportLocation { get; set; } = "";
    /// <summary>
    /// Specifies the file system path where generated DXF files will be exported.
    /// </summary>
    public string DxfExportLocation { get; set; } = "";
    /// <summary>
    /// Feature toggles
    /// </summary>
    public bool EnableObsoletePrint { get; set; } = true;

    /// <summary>
    /// The file path where the user options data is stored.
    /// This file contains serialized user settings and preferences.
    /// </summary>
    public static readonly string OptionsFilePath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DoyleAddinOptions.xml");


    /// <summary>
    /// Saves the current configuration values of the <see cref="UserOptions"/> object
    /// to a file in XML format. The file is written to a predefined path within the user's
    /// application data directory, overwriting any existing file.
    /// </summary>
    public void Save()
    {
        var serializer = new XmlSerializer(typeof(UserOptions));
        using var writer = new StreamWriter(OptionsFilePath);
        serializer.Serialize(writer, this);
    }
    /// <summary>
    /// Loads the user options settings from a file if it exists. If the file is not found,
    /// a new instance of the UserOptions class is returned with default values.
    /// </summary>
    /// <returns>A <see cref="UserOptions"/> instance containing the loaded settings or default values if no file exists.</returns>
    public static UserOptions Load()
    {
        if (File.Exists(OptionsFilePath))
        {
            var serializer = new XmlSerializer(typeof(UserOptions));
            using var reader = new StreamReader(OptionsFilePath);
            return (UserOptions)serializer.Deserialize(reader);
        }
        else
        {
            return new UserOptions();
        }
    }
}