using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

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

        public List<List<string>> Load()
        {
            var sheetData = new List<List<string>>();

            for(var i = rowStart;i<= rowEnd; i++)
            {
                var row = sheet.GetRow(i);
                var parsedRow = ReadRow(row);

                sheetData.Add(parsedRow);
            }

            return sheetData;
        }

        List<string> ReadRow(IRow row)
        {
            var rowData = new List<string>();
            var cursor = 0;
            
            foreach(var cell in row.Cells)
            {
                rowData.Add(cell.StringCellValue);
                ++cursor;
            }

            for (var i = cursor; i < maxColumnWidth; i++)
            {
                rowData.Add(string.Empty);
            }

            return rowData;
        }
    }
}
