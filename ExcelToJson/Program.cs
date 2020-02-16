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
        static void Main(string[] args)
        {
            switch (args.Count())
            {
                case 0:
                    DefaultBehavior();
                    return;
                
                case 1:
                    if(args[0] == "-h")
                    {
                        Console.WriteLine("-i <importDirectory> : Change Excel Importing Directory.");
                        Console.WriteLine("-e <exportDirectory> : Change JSON Exporting Directory.");
                        return;
                    }

                    throw new ArgumentException("Invalid Argument");
                
                case 2:
                    if (args[0] == "-i")
                    {
                        ChangeConfigImportDirectory(args[1]);
                        return;
                    }
                    else if (args[0] == "-e")
                    {
                        ChangeConfigExportDirectory(args[1]);
                        return;
                    }

                    throw new ArgumentException("Invalid Argument");
                
                default:
                    throw new ArgumentException("Invalid Argument");
            }
        }

        static void DefaultBehavior()
        {
            if (!Configuration.HasConfigFile())
            {
                Configuration.Write(".", ".");
            }

            var config = Configuration.Read();
            ContertExcelDirectoryToJson(config.ImportDirectory, config.ExportDirectory);
        }

        static void ChangeConfigImportDirectory(string importDir)
        {
            var config = Configuration.Read();

            Configuration.Write(importDir, config.ExportDirectory);
        }

        static void ChangeConfigExportDirectory(string exportDir)
        {
            var config = Configuration.Read();

            Configuration.Write(config.ImportDirectory, exportDir);
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
