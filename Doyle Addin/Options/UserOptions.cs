using System.Xml.Serialization;

namespace Doyle_Addin.Options
{
    /// <summary>
    /// 
    /// </summary>
    public class UserOptions
    {
        /// <summary>
        /// 
        /// </summary>
        public string PrintExportLocation { get; set; } = "";

        /// <summary>
        /// 
        /// </summary>
        public string DxfExportLocation { get; set; } = "";

        // Feature toggles
        /// <summary>
        /// 
        /// </summary>
        public bool EnableObsoletePrint { get; set; } = true;

        /// <summary>
        /// 
        /// </summary>
        public static readonly string OptionsFilePath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DoyleAddinOptions.xml");


        /// <summary>
        /// 
        /// </summary>
        public void Save()
        {
            var serializer = new XmlSerializer(typeof(UserOptions));
            using var writer = new StreamWriter(OptionsFilePath);
            serializer.Serialize(writer, this);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static UserOptions Load()
        {
            if (File.Exists(OptionsFilePath))
            {
                var serializer = new XmlSerializer(typeof(UserOptions));
                using var reader = new StreamReader(OptionsFilePath);
                return ((UserOptions)serializer.Deserialize(reader)!);
            }
            else
            {
                return new UserOptions();
            }
        }
    }
}