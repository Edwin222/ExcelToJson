using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelToJson.ExcelControl
{
    public class ExcelSheetLoader
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

        public ExcelSheet Load()
        {
            var sheetData = new List<List<string>>();
            var properties = ReadRow(sheet.GetRow(rowStart));

            for (var i = rowStart + 1;i<= rowEnd; i++)
            {
                var row = sheet.GetRow(i);
                var parsedRow = ReadRow(row);

                sheetData.Add(parsedRow);
            }

            return new ExcelSheet(sheet.SheetName, properties, sheetData);
        }

        List<string> ReadRow(IRow row)
        {
            var rowData = new List<string>();
            var cursor = 0;
            
            foreach(var cell in row.Cells)
            {
                rowData.Add(cell.ToString());
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
