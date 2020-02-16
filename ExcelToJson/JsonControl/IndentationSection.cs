using System;
using System.Collections.Generic;

namespace ExcelToJson.JsonConverter
{
    public class IndentationSection : IDisposable
    {
        readonly List<IndentationSection> sectionCollection;

        public IndentationSection(List<IndentationSection> collection)
        {
            sectionCollection = collection;
            sectionCollection.Add(this);
        }

        public void Dispose()
        {
            sectionCollection.Remove(this);
        }
    }
}
