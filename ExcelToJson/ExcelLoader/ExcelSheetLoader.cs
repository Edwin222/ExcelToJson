using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToJson.ExcelLoader
{
    class ExcelSheetLoader
    {
        readonly ISheet sheet;
        readonly int rowStart;
        readonly int rowEnd;
        readonly int maxColumnWidth;

        public ExcelSheetLoader(ISheet sheet)
        {
            this.sheet = sheet;

            rowStart = sheet.FirstRowNum;
            rowEnd = sheet.LastRowNum;

            maxColumnWidth = 0;

            for (var i = rowStart; i <= rowEnd; i++)
            {
                var currentColumn = sheet.GetColumnWidth(i);
                maxColumnWidth = Math.Max(maxColumnWidth, currentColumn);
            }
        }

        string ReadRow(IRow row)
        {

        }
    }
}
