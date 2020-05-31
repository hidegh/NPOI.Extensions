using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using Xunit;

namespace NPOI.Extensions.XUnitTest.Test_GetValueExtensions
{
    /// <summary>
    /// See also: CollectionFixture.cs
    /// </summary>
    public class ClassFixture : IClassFixture<ClassFixture.Setup>
    {
        public class Setup : IDisposable
        {
            protected const string testExcelFileName = "./Npoi/Test.xlsx";

            public XSSFWorkbook Workbook { get; }
            public ISheet Sheet { get; }

            public Setup()
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
            }

            public void Dispose()
            {

            }
        }

        protected readonly Setup Fixture;

        [Fact(Skip = "This is just an IClassFixture sample.")]
        public void Test_CanAccessFixtureData()
        {
            Assert.NotNull(Fixture.Workbook);
        }

    }
}
