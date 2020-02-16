using System;
using System.IO;
using ExcelToJson.ExcelControl;
using ExcelToJson.JsonConverter;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var fileStream = new FileStream("mockupData.xlsx", FileMode.Open, FileAccess.Read))
            {
            }
        }
    }
}
