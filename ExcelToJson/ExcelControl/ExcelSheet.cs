using System.Collections.Generic;

namespace ExcelToJson.ExcelControl
{
    public class ExcelSheet
    {
        public readonly string Name;
        public readonly List<string> Properties;
        public readonly List<List<string>> Sheet;

        public ExcelSheet(string name, List<string> properties, List<List<string>> sheet)
        {
            Name = name;
            Properties = properties;
            Sheet = sheet;
        }
    }
}
