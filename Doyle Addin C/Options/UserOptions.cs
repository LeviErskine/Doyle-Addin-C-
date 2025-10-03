using System;
using System.IO;
using System.Xml.Serialization;

namespace Doyle_Addin.Options
{
    public class UserOptions
    {
        public string PrintExportLocation { get; set; } = "";
        public string DxfExportLocation { get; set; } = "";

        // Feature toggles
        public bool EnableObsoletePrint { get; set; } = true;

        public static readonly string OptionsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DoyleAddinOptions.xml");


        public void Save()
        {
            var serializer = new XmlSerializer(typeof(UserOptions));
            using (var writer = new StreamWriter(OptionsFilePath))
            {
                serializer.Serialize(writer, this);
            }
        }

        public static UserOptions Load()
        {
            if (File.Exists(OptionsFilePath))
            {
                var serializer = new XmlSerializer(typeof(UserOptions));
                using (var reader = new StreamReader(OptionsFilePath))
                {
                    return (UserOptions)serializer.Deserialize(reader);
                }
            }
            else
            {
                return new UserOptions();
            }
        }
    }
}