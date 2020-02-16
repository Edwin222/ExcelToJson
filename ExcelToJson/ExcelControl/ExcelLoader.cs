using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelToJson.ExcelControl
{
    public class ExcelLoader
    {
        readonly IWorkbook excelWorkbook;
        int currentSheetNumber;
        int maxSheetNumber => excelWorkbook.NumberOfSheets;
        public bool HasSheetToRead() => currentSheetNumber < maxSheetNumber;

        public ExcelLoader(FileStream fileStream)
        {
            excelWorkbook = new XSSFWorkbook(fileStream);
            currentSheetNumber = 0;
        }

        public ExcelSheet LoadNextSheet()
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
