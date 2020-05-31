using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Diagnostics;
using System.IO;
using Xunit;

namespace NPOI.Extensions.XUnitTest.Test_GetValueExtensions
{
    [CollectionDefinition("Test_GetValueExtensions")]
    public class DatabaseCollection : ICollectionFixture<CollectionFixture>
    {
        // This class has no code, and is never created.
        // Its purpose is simply to be the place to apply [CollectionDefinition] and all the ICollectionFixture<> interfaces.
    }

    /// <summary>
    /// https://xunit.net/docs/shared-context
    /// </summary>
    public class CollectionFixture : IDisposable
    {
        protected const string testExcelFileName = "./Test_GetValueExtensions/Test.xlsx";

        public XSSFWorkbook Workbook { get; }
        public ISheet Sheet { get; }

        public CollectionFixture()
        {
            // For .NET Core (& xUnit)
            //
            // Add: System.Text.Encoding.CodePages
            // Extend: System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            //
            // You will also need: System.Configuration.ConfigurationManager
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var fileReader = new FileStream(testExcelFileName, FileMode.Open, FileAccess.Read))
                Workbook = new XSSFWorkbook(fileReader);

            Sheet = Workbook.GetSheetAt(0);

            var textCell = Sheet.GetRow((int)RowIndexEnum.Text).GetCell(1);
            var numberCell = Sheet.GetRow((int)RowIndexEnum.Number).GetCell(1);
            var boolCell = Sheet.GetRow((int)RowIndexEnum.Bool).GetCell(1);
            var numberFormattedCell = Sheet.GetRow((int)RowIndexEnum.NumberFormatted).GetCell(1);
            var numberFormattedAsCurrencyCell = Sheet.GetRow((int)RowIndexEnum.NumberFormattedAsCurrency).GetCell(1);
            var dateCell = Sheet.GetRow((int)RowIndexEnum.Date).GetCell(1);
            var timeCell = Sheet.GetRow((int)RowIndexEnum.Time).GetCell(1);
            var blankCell = Sheet.GetRow((int)RowIndexEnum.Blank).GetCell(1, MissingCellPolicy.RETURN_NULL_AND_BLANK);
            var richTextCell = Sheet.GetRow((int)RowIndexEnum.RichText).GetCell(1);
            var formulaCellNumeric = Sheet.GetRow((int)RowIndexEnum.FormulaNumeric).GetCell(1);
            var formulaCellString = Sheet.GetRow((int)RowIndexEnum.FormulaString).GetCell(1);

            Debugger.Break();
        }

        public void Dispose()
        {

        }

    }

}
