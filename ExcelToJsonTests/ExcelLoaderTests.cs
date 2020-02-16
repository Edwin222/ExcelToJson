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
            new List<string> { "1��", "��", "��", "��", "" },
            new List<string> { "2��", "", "��", "��", "��" },
            new List<string> { "3��", "��", "", "��", "��" },
            new List<string> { "4��", "��", "ī", "", "Ÿ" },
        };

        List<string> secondSheetProperties = new List<string> { "Key", "Value" };
        List<List<string>> secondSheetExpected = new List<List<string>>
        {
            new List<string> { "��", "1"},
            new List<string> { "��", "2"},
            new List<string> { "��", ""},
            new List<string> { "��", "4"},
            new List<string> { "��", ""},
            new List<string> { "��", "6"},
        };

        [TestMethod]
        public void �⺻_����_�׽�Ʈ()
        {
            using(var fileStream = new FileStream("mockupData.xlsx", FileMode.Open, FileAccess.Read))
            {
                var excelLoader = new ExcelLoader(fileStream);

                var firstSheet = excelLoader.LoadNextSheet();
                var secondSheet = excelLoader.LoadNextSheet();

                // NOTE @sleepyrainyday 200216: 3��° ��Ʈ�� �����Ƿ� null
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
