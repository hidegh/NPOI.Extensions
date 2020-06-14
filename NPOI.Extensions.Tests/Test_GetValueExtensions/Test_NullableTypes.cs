using System;
using Xunit;

namespace NPOI.Extensions.XUnitTest.Test_GetValueExtensions
{
    [Collection("Test_GetValueExtensions")]
    public class Test_NullableTypes
    {
        protected CollectionFixture Fixture { get; }

        public Test_NullableTypes(CollectionFixture fixture)
        {
            this.Fixture = fixture;
        }

        //
        //
        //

        [Fact]
        public void Test_Read_Char()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<char?>(1);
            Assert.Equal((char?)19, result);
        }

        //
        //
        //

        [Fact]
        public void Test_Read_SByte()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<sbyte?>(1);
            Assert.Equal((sbyte?)19, result);
        }

        [Fact]
        public void Test_Read_Byte()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<byte?>(1);
            Assert.Equal((byte?)19, result);
        }

        [Fact]
        public void Test_Read_Int16()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<Int16?>(1);
            Assert.Equal((Int16?)19, result);
        }

        [Fact]
        public void Test_Read_UInt16()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<UInt16?>(1);
            Assert.Equal((UInt16?)19, result);
        }

        [Fact]
        public void Test_Read_Int32()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<Int32?>(1);
            Assert.Equal(19, result);
        }

        [Fact]
        public void Test_Read_UInt32()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<UInt32?>(1);
            Assert.Equal((uint)19, result);
        }

        [Fact]
        public void Test_Read_Int64()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<Int64?>(1);
            Assert.Equal(19, result);
        }

        [Fact]
        public void Test_Read_UInt64()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<UInt64?>(1);
            Assert.Equal((ulong)19, result);
        }

        //
        //
        //

        [Fact]
        public void Test_Read_Single()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<Single?>(1);
            Assert.Equal((Single)19.2, result);
        }

        [Fact]
        public void Test_Read_Double()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<Double?>(1);
            Assert.Equal((Double)19.2, result);
        }

        [Fact]
        public void Test_Read_Decimal()
        {
            var result = Fixture.Sheet.GetRow((int)RowIndexEnum.Number).GetValue<Decimal?>(1);
            Assert.Equal(19.2m, result);
        }

    }

}
