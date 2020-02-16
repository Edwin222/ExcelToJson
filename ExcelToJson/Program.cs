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
            ConvertExcelToJson("mockupData.xlsx", ".");
        }

        static void ConvertExcelToJson(string excelFilePath, string resultDir)
        {
            using (var excelFile = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                var excelLoader = new ExcelLoader(excelFile);

                while (excelLoader.HasSheetToRead())
                {
                    var excelSheet = excelLoader.LoadNextSheet();

                    var sheetFileName = $"{GetNameFromPath(excelFile.Name)}.{excelSheet.Name}.json";
                    using (var jsonFile = new FileStream($"{resultDir}/{sheetFileName}", FileMode.Create, FileAccess.Write))
                    {
                        using (var streamWriter = new StreamWriter(jsonFile))
                        {
                            var collectionToJsonConverter = new CollectionToJsonConverter(excelSheet.Properties);

                            streamWriter.Write(collectionToJsonConverter.CreateJSONLiteral(excelSheet.Sheet));
                        }
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
