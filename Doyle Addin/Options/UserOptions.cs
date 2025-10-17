using System;
using System.ComponentModel;
using System.IO;
using System.Xml.Serialization;

namespace Doyle_Addin.Options;

/// <summary>
/// Represents a collection of user-configurable options and preferences related to
/// the Doyle Add-in. This class provides methods to save and load these options
/// as an XML file in the user's application data directory.
/// </summary>
public class UserOptions : INotifyPropertyChanged
{
    /// <summary>
    /// Raised when a property value changes; used by WPF data binding (INotifyPropertyChanged).
    /// </summary>
    public event PropertyChangedEventHandler PropertyChanged;

    private void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    private readonly string printExportLocation = "";
    private readonly string dxfExportLocation = "";
    private bool enableObsoletePrint = true;

    /// <summary>
    /// Specifies the directory path where PDF files will be exported during the print update process.
    /// </summary>
    public string PrintExportLocation
    {
        get => printExportLocation;
        init
        {
            if (value == printExportLocation) return;
            printExportLocation = value;
            OnPropertyChanged(nameof(PrintExportLocation));
        }
    }
    /// <summary>
    /// Specifies the file system path where generated DXF files will be exported.
    /// </summary>
    public string DxfExportLocation
    {
        get => dxfExportLocation;
        init
        {
            if (value == dxfExportLocation) return;
            dxfExportLocation = value;
            OnPropertyChanged(nameof(DxfExportLocation));
        }
    }
    /// <summary>
    /// Feature toggles
    /// </summary>
    public bool EnableObsoletePrint
    {
        get => enableObsoletePrint;
        set
        {
            if (value == enableObsoletePrint) return;
            enableObsoletePrint = value;
            OnPropertyChanged(nameof(EnableObsoletePrint));
        }
    }

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
        try
        {
            var serializer = new XmlSerializer(typeof(UserOptions));
            using var writer = new StreamWriter(OptionsFilePath);
            serializer.Serialize(writer, this);
        }
        catch (Exception)
        {
            // Intentionally swallow exceptions to avoid crashing the host application.
            // In a future change we could log this to a telemetry sink or show a user-facing message.
        }
    }
    /// <summary>
    /// Loads the user options settings from a file if it exists. If the file is not found,
    /// a new instance of the UserOptions class is returned with default values.
    /// </summary>
    /// <returns>A <see cref="UserOptions"/> instance containing the loaded settings or default values if no file exists or on error.</returns>
    public static UserOptions Load()
    {
        try
        {
            if (!File.Exists(OptionsFilePath)) return new UserOptions();
            var serializer = new XmlSerializer(typeof(UserOptions));
            using var reader = new StreamReader(OptionsFilePath);
            return (UserOptions)serializer.Deserialize(reader);
        }
        catch (Exception)
        {
            // If loading fails for any reason (corrupt file, permission issue, etc.)
            // return default options to keep the application usable.
            return new UserOptions();
        }

    }
}