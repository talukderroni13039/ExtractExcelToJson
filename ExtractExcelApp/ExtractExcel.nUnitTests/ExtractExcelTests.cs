using ExtractExcelApp.Services;
namespace ExtractExcel.nUnitTests
{
    public class ExtractExcelTests
    {
        public ExtractExccl _extractExcel { get; set; } = null;

        [SetUp]
        public void Setup()
        {
            string filePath = "D:\\ExtractExcelApp\\Input\\Bayer Accounts.xlsx";
            _extractExcel =new ExtractExccl();
        }

        [TestCase(new string[] { "*"}, 1, "D:\\ExtractExcelApp\\Input\\Bayer Accounts.xlsx")]
        [TestCase(new string[] { "", "" }, 1, "D:\\ExtractExcelApp\\Input\\Bayer Accounts.xlsx")]
        [TestCase(new string[] { "Sheet1" }, 1 ,"D:\\ExtractExcelApp\\Input\\Bayer Account1.xlsx")]
        [TestCase(new string[] { "*" }, 1, "D:\\ExtractExcelApp\\Input\\Bayer Accounts.xlsx")]
        public void ExtractExcelTest(string[] sheetName, int columnIndex,string filePath)
        {
            _extractExcel.ExtractDatafromFastExcel(sheetName, columnIndex, filePath);

            Assert.Pass();
        }
    }
}