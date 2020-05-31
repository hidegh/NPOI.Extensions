using Xunit;

namespace NPOI.Extensions.XUnitTest.Test_GetValueExtensions
{
    [Collection("Test_GetValueExtensions")]
    public class Test_Formulas
    {
        protected CollectionFixture Fixture { get; }

        public Test_Formulas(CollectionFixture fixture)
        {
            this.Fixture = fixture;
        }

        [Fact]
        public void Test_CanRead_NumericResult_FromExcelCellWithFormula()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.FormulaNumeric).GetValue<decimal>(1);
            Assert.Equal(2501.25m, result);
        }

        [Fact]
        public void Test_CanRead_StringResult_FromExcelCellWithFormula()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.FormulaString).GetValue<string>(1);
            Assert.Equal("Date-Time", result);
        }

    }
}
