using DocumentFormat.OpenXml.Packaging;

namespace ExcelTools
{
    public sealed class Program
    {
        private const string SampleExcelFile = "SampleData.xlsx";

        public static void Main(string[] args)
        {
            using (var document = SpreadsheetDocument.Open(SampleExcelFile, isEditable: true))
            {
            }
        }
    }
}
