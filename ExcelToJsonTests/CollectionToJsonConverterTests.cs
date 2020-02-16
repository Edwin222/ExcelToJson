using ExcelToJson.JsonConverter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ExcelToJsonTests
{
    [TestClass]
    public class CollectionToJsonConverterTests
    {
        [TestMethod]
        public void JSON_객체생성_테스트()
        {
            var properties = new List<string> { "1", "2", "3" };
            var data = new List<string> { "가", "나", "다" };

            var collectionToJsonConverter = new CollectionToJsonConverter(properties);
            Assert.AreEqual(
                "{\n" +
                "    \"1\": \"가\",\n" +
                "    \"2\": \"나\",\n" +
                "    \"3\": \"다\"\n" +
                "}",
                collectionToJsonConverter.CreateJSONObject(data));
        }

        [TestMethod]
        public void JSON_리터럴생성_테스트()
        {
            var properties = new List<string> { "1", "2", "3" };
            var data = new List<List<string>>
            {
                new List<string> { "가", "나", "다" },
                new List<string> { "라", "마", "바" },
                new List<string> { "사", "아", "자" },
            };

            var collectionToJsonConverter = new CollectionToJsonConverter(properties);

            Assert.AreEqual(
                "{\n" +
                "    \"Data\": [\n" +
                "        {\n" +
                "            \"1\": \"가\",\n" +
                "            \"2\": \"나\",\n" +
                "            \"3\": \"다\"\n" +
                "        },\n" +
                "        {\n" +
                "            \"1\": \"라\",\n" +
                "            \"2\": \"마\",\n" +
                "            \"3\": \"바\"\n" +
                "        },\n" +
                "        {\n" +
                "            \"1\": \"사\",\n" +
                "            \"2\": \"아\",\n" +
                "            \"3\": \"자\"\n" +
                "        }\n" +
                "    ]\n" +
                "}", 
                collectionToJsonConverter.CreateJSONLiteral(data));
        }
    }
}
