using System;
using Xunit;

namespace NPOI.Extensions.XUnitTest.Test_GetValueExtensions
{
    [Collection("Test_GetValueExtensions")]
    public class Test_ReadAsString
    {
        protected CollectionFixture Fixture { get; }

        public Test_ReadAsString(CollectionFixture fixture)
        {
            this.Fixture = fixture;
        }

        //
        //
        //

        [Fact]
        public void Test_CanRead_Number_AsString()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<string>(1);
            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.Equal(19.2m.ToString(), result);
        }

        [Fact(Skip = "Results depend on language settings")]
        public void Test_CanRead_Bool_AsString()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Bool).GetValue<string>(1);
            Assert.NotNull(result);
            Assert.NotEmpty(result);
            // NOTE: result might vary due to windows language settings
            Assert.Equal("TRUE", result);
        }

        [Fact]
        public void Test_CanRead_Date_AsString()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Date).GetValue<string>(1);
            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.Equal(new DateTime(2001, 2, 22).ToString("dddd, MMMM dd, yyyy"), result);
        }

        [Fact()]
        public void Test_CanRead_Time_AsString()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Time).GetValue<string>(1);
            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.Equal(new DateTime().Add(TimeSpan.FromHours(23)).ToString("hh:mm tt"), result);
        }

        [Fact]
        public void Test_CanRead_FormattedNumeric_AsString()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.NumberFormatted).GetValue<string>(1);
            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.Equal(1000.5m.ToString("N2"), result);
        }

        [Fact]
        public void Test_CanRead_FormattedCurrency_AsString()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.NumberFormattedAsCurrency).GetValue<string>(1);
            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.Equal(1500.75m.ToString("N2") + " €", result);
        }

    }
}
