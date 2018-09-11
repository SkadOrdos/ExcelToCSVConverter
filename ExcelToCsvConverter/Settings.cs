using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace ExcelToCsvConverter
{
    [Serializable]
    public class Settings
    {
        public const string DefaultFileExtension = "csv";

        public Settings()
        {
            WokrbookFile = String.Empty;
            ExportBookFile = "locales.xlsx";
            ProcessedSheetCount = 1;
            OutFileExtension = "csv";
            OutFileSeparator = "=";
        }


        public string WokrbookFile { get; set; }

        public string ExportBookFile { get; set; }

        public IEnumerable<string> WorkFiles
        {
            get
            {
                if (!String.IsNullOrEmpty(WokrbookFile)) return new List<string> { WokrbookFile };
                return Directory.GetFiles(Environment.CurrentDirectory, "*.xlsx").Where(f => !f.StartsWith("$"));
            }
        }


        public int ProcessedSheetCount { get; set; }

        public string OutFileExtension { get; set; }

        public string GetFormatExtension
        {
            get
            {
                if (!String.IsNullOrEmpty(OutFileExtension))
                {
                    if (!OutFileExtension.StartsWith("."))
                        return "." + OutFileExtension;
                    return OutFileExtension;
                }

                return DefaultFileExtension;
            }
        }

        public string OutFileSeparator { get; set; }



        public static void SaveToXml(String fileName, Settings serializableObject)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Settings));
            using (TextWriter textWriter = new StreamWriter(fileName))
            {
                serializer.Serialize(textWriter, serializableObject);
            }
        }

        public static Settings LoadFromXml(String fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Settings));
            using (TextReader textReader = new StreamReader(fileName))
            {
                return (Settings)serializer.Deserialize(textReader);
            }
        }
    }
}
