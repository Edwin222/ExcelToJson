using ExcelToJson.ExcelControl;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace ExcelToJsonTests
{
    [TestClass]
    public class ExcelLoaderTests
    {
        List<string> firstSheetProperties = new List<string> { "Key", "Property1", "Property2", "Property3", "Property4" };
        List<List<string>> firstSheetExpected = new List<List<string>>
        {
            new List<string> { "1번", "가", "나", "다", "" },
            new List<string> { "2번", "", "라", "마", "바" },
            new List<string> { "3번", "사", "", "아", "자" },
            new List<string> { "4번", "차", "카", "", "타" },
        };

        List<string> secondSheetProperties = new List<string> { "Key", "Value" };
        List<List<string>> secondSheetExpected = new List<List<string>>
        {
            new List<string> { "가", "1"},
            new List<string> { "나", "2"},
            new List<string> { "다", ""},
            new List<string> { "라", "4"},
            new List<string> { "마", ""},
            new List<string> { "바", "6"},
        };

        [TestMethod]
        public void 기본_동작_테스트()
        {
            using(var fileStream = new FileStream("mockupData.xlsx", FileMode.Open, FileAccess.Read))
            {
                var excelLoader = new ExcelLoader(fileStream);

                var firstSheet = excelLoader.LoadNextSheet();
                var secondSheet = excelLoader.LoadNextSheet();

                // NOTE @sleepyrainyday 200216: 3번째 시트는 없으므로 null
                Assert.IsNull(excelLoader.LoadNextSheet());

                AreEqual(firstSheetProperties, firstSheet.Properties);
                AreEqual(firstSheetExpected, firstSheet.Sheet);

                AreEqual(secondSheetProperties, secondSheet.Properties);
                AreEqual(secondSheetExpected, secondSheet.Sheet);
            }
        }

        public void AreEqual(List<List<string>> expected, List<List<string>> actual)
        {
            Assert.AreEqual(expected.Count, actual.Count);

            for (var i = 0; i < expected.Count; i++)
            {
                AreEqual(expected[i], actual[i]);
            }
        }

        public void AreEqual(List<string> expected, List<string> actual)
        {
            Assert.AreEqual(expected.Count, actual.Count);

            for(var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], actual[i]);
            }
        }
    }
}
