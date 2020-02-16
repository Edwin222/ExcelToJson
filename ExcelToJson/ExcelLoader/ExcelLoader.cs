using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace ExcelToJson.ExcelLoader
{
    public class ExcelLoader
    {
        readonly IWorkbook excelWorkbook;
        int currentSheetNumber;
        int maxSheetNumber => excelWorkbook.NumberOfSheets;

        public ExcelLoader(FileStream fileStream)
        {
            excelWorkbook = new XSSFWorkbook(fileStream);
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
