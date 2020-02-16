using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelToJson.ExcelLoader
{
    class ExcelLoader
    {
        readonly IWorkbook excelWorkbook;
        int currentSheetNumber;
        int maxSheetNumber => excelWorkbook.NumberOfSheets;

        public ExcelLoader(FileStream fileStream)
        {
            excelWorkbook = new HSSFWorkbook(fileStream);
            currentSheetNumber = 0;
        }

        public string ReadNextSheet()
        {
            if(currentSheetNumber < maxSheetNumber)
            {
                var sheet = excelWorkbook.GetSheetAt(currentSheetNumber);

                ++currentSheetNumber;
            }

            return null;
        }
    }
}
