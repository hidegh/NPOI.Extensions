using System;
using Xunit;

namespace NPOI.Extensions.XUnitTest.Test_GetValueExtensions
{
    [Collection("Test_GetValueExtensions")]
    public class Test_Blanks
    {
        protected CollectionFixture Fixture { get; }

        public Test_Blanks(CollectionFixture fixture)
        {
            this.Fixture = fixture;
        }

        [Fact]
        public void Test_CanReadBlankCell_IntoString()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<string>(1);
            Assert.Null(result);
        }

        //
        //
        //

        [Fact]
        public void Test_CanReadBlankCell_IntoNullableInt()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<int?>(1);
            Assert.Null(result);
        }

        [Fact]
        public void Test_CanReadBlankCell_IntoNullableDecimal()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<decimal?>(1);
            Assert.Null(result);
        }

        [Fact]
        public void Test_CanReadBlankCell_IntoNullableBool()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<bool?>(1);
            Assert.Null(result);
        }

        [Fact]
        public void Test_CanReadBlankCell_IntoNullableNullableDate()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<DateTime?>(1);
            Assert.Null(result);
        }

        [Fact]
        public void Test_CanReadBlankCell_IntoNullableTimeSpan()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<TimeSpan?>(1);
            Assert.Null(result);
        }


        //
        //
        //

        [Fact]
        public void Test_CantReadBlankCell_IntoNonNullableInt()
        {
            Assert.Throws<NotSupportedException>(() => Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<int>(1));
        }

        [Fact]
        public void Test_CantReadBlankCell_IntoNonNullableDecimal()
        {
            Assert.Throws<NotSupportedException>(() => Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<decimal>(1));
        }

        [Fact]
        public void Test_CantReadBlankCell_IntoNonNullableBool()
        {
            Assert.Throws<NotSupportedException>(() => Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<bool>(1));
        }

        [Fact]
        public void Test_CantReadBlankCell_IntoNonNullableNullableDate()
        {
            Assert.Throws<NotSupportedException>(() => Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<DateTime>(1));
        }

        [Fact]
        public void Test_CantReadBlankCell_IntoNonNullableTimeSpan()
        {
            Assert.Throws<NotSupportedException>(() => Fixture.Sheet.GetRow((int)RowIndexEnum.Blank).GetValue<TimeSpan>(1));
        }

    }
}
