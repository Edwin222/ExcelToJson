using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelToJson.ExcelControl;
using ExcelToJson.JsonConverter;

namespace ExcelToJson
{
    class Program
    {
        static readonly string configFileName = ".config";

        static void Main(string[] args)
        {
            var currentDir = Directory.GetCurrentDirectory();
            var hasConfigFile = Directory.GetFiles(currentDir).Contains($"{currentDir}\\{configFileName}");

            if (!hasConfigFile)
            {
                CreateConfigFile(".", ".");
            }

            using (var configStream = new FileStream(configFileName, FileMode.Open, FileAccess.Read))
            {
                using (var streamReader = new StreamReader(configStream))
                {
                    var importDir = streamReader.ReadLine();
                    var exportDir = streamReader.ReadLine();

                    ContertExcelDirectoryToJson(importDir, exportDir);
                }
            }    
        }

        static void CreateConfigFile(string importDir, string exportDir)
        {
            using (var configWritingStream = new FileStream(configFileName, FileMode.Create, FileAccess.Write))
            {
                using (var streamWriter = new StreamWriter(configWritingStream))
                {
                    streamWriter.WriteLine(importDir);
                    streamWriter.WriteLine(exportDir);
                }
            }
        }

        static void ContertExcelDirectoryToJson(string importDir, string exportDir)
        {
            var filePaths = Directory.GetFiles(importDir, "*.xlsx");

            foreach(var filePath in filePaths)
            {
                Console.WriteLine($"Import: {filePath}");
                ConvertExcelToJson(filePath, exportDir);
            }
        }

        static void ConvertExcelToJson(string excelFilePath, string exportDir)
        {
            using (var excelFile = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                var excelLoader = new ExcelLoader(excelFile);

                while (excelLoader.HasSheetToRead())
                {
                    var excelSheet = excelLoader.LoadNextSheet();

                    var sheetFileName = $"{GetNameFromPath(excelFile.Name)}.{excelSheet.Name}.json";
                    using (var jsonFile = new FileStream($"{exportDir}/{sheetFileName}", FileMode.Create, FileAccess.Write))
                    {
                        using (var streamWriter = new StreamWriter(jsonFile))
                        {
                            var collectionToJsonConverter = new CollectionToJsonConverter(excelSheet.Properties);

                            streamWriter.Write(collectionToJsonConverter.CreateJSONLiteral(excelSheet.Sheet));
                        }

                        Console.WriteLine($"Exported: {jsonFile.Name}");
                    }
                }
            }
        }

        static string GetNameFromPath(string path)
        {
            var regEx = new Regex(@"\\([a-zA-z0-9]+)\.xlsx$");

            var matches = regEx.Matches(path);

            return matches.Last().Groups[1].Value;
        }
    }
}
