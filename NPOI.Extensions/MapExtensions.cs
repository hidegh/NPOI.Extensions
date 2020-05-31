using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.Extensions
{
    public static class MapExtensions
    {
        public class ExcelSheetHeaderDetail
        {
            public string Title { get; set; }
            public int ColumnIndex { get; set; }
        }

        public class ExcelRowMapper
        {
            public IRow Row { get; }

            protected IList<ExcelSheetHeaderDetail> HeaderDetails { get; }

            public ExcelRowMapper(IRow row, IList<ExcelSheetHeaderDetail> headerDetails = null)
            {
                this.Row = row;
                this.HeaderDetails = headerDetails;
            }

            public T GetValue<T>(string title)
            {
                if (this.HeaderDetails == null)
                    throw new NotSupportedException();

                var headerInfo = this.HeaderDetails.First(i => i.Title == title);
                return this.Row.GetValue<T>(headerInfo.ColumnIndex);
            }

            public T GetValue<T>(int columnIndex)
            {
                return this.Row.GetValue<T>(columnIndex);
            }

        }

        public static List<T> MapTo<T>(this ISheet sheet, bool firstRowIsHeader, Func<ExcelRowMapper, T> rowMapperFunction)
            where T : class
        {
            var result = new List<T>();

            var header = sheet.GetRow(sheet.FirstRowNum);
            var headerDetails = firstRowIsHeader
                ? header.Cells.Select(c => new ExcelSheetHeaderDetail() { Title = c.StringCellValue, ColumnIndex = c.ColumnIndex }).ToList()
                : null as IList<ExcelSheetHeaderDetail>;

            for (var rowIndex = sheet.FirstRowNum + (firstRowIsHeader ? 1 : 0); rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                var rowMapper = new ExcelRowMapper(row, headerDetails);

                var mappedRow = rowMapperFunction(rowMapper);
                result.Add(mappedRow);
            }

            return result;
        }
    }
}
