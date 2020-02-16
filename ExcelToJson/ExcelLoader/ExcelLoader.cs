using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;

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

        public List<List<string>> LoadNextSheet()
        {
            if(currentSheetNumber < maxSheetNumber)
            {
                var sheet = excelWorkbook.GetSheetAt(currentSheetNumber);
                ++currentSheetNumber;

                var sheetLoader = new ExcelSheetLoader(sheet);

                return sheetLoader.Load();
            }

            return null;
        }
    }
}
