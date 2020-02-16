using System.IO;
using System.Linq;

namespace ExcelToJson
{
    public class Configuration
    {
        static readonly string configFileName = ".config";

        public readonly string ImportDirectory;
        public readonly string ExportDirectory;

        private Configuration(string import, string export)
        {
            ImportDirectory = import;
            ExportDirectory = export;
        }

        public static bool HasConfigFile()
        {
            var currentDir = Directory.GetCurrentDirectory();
            return Directory.GetFiles(currentDir).Contains($"{currentDir}\\{configFileName}");
        }

        public static void Write(string importDir, string exportDir)
        {
            using (var configStream = new FileStream(configFileName, FileMode.Create, FileAccess.Write))
            {
                using (var streamWriter = new StreamWriter(configStream))
                {
                    streamWriter.WriteLine(importDir);
                    streamWriter.WriteLine(exportDir);
                }
            }
        }

        public static Configuration Read()
        {
            using (var configStream = new FileStream(configFileName, FileMode.Open, FileAccess.Read))
            {
                using (var streamReader = new StreamReader(configStream))
                {
                    var import = streamReader.ReadLine();
                    var export = streamReader.ReadLine();

                    return new Configuration(import, export);
                }
            }
        }

    }
}
