using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToJson.JsonConverter
{
    public class JsonConverter
    {
        static readonly string indentation = "    ";
        
        readonly List<string> properties;
        readonly List<IndentationSection> indentationSectionCollection;

        int indentationDepth => indentationSectionCollection.Count;

        public JsonConverter(List<string> properties) 
        {
            this.properties = properties;
            indentationSectionCollection = new List<IndentationSection>();
        }
        
        public string CreateJSONLiteral(List<List<string>> sheetData)
        {
            var result = AttachIndentation("{\n");

            foreach(var row in sheetData)
            {
                using(new IndentationSection(indentationSectionCollection))
                {
                    result += CreateJSONObject(row);

                    if(row != sheetData.Last())
                    {
                        result += ",\n";
                    }
                }
            }

            result += AttachIndentation("\n}");

            return result;
        }

        public string CreateJSONObject(List<string> dataRow)
        {
            var result = AttachIndentation("{\n");
            
            for(var i = 0; i < properties.Count; i++)
            {
                using (new IndentationSection(indentationSectionCollection))
                {
                    var rowString = AttachIndentation($"\"{properties[i]}\": \"{dataRow[i]}\"");
                    result += AttachIndentation($"{indentation}{rowString}\n");
                }
            }

            result += AttachIndentation("}");

            return result;
        }

        string AttachIndentation(string value)
        {
            var result = "";

            for(var i = 0; i < indentationDepth; i++)
            {
                result += indentation;
            }

            result += value;

            return result;
        }
    }
}
