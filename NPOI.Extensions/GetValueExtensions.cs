using NPOI.Extensions.Utils;
using NPOI.SS.UserModel;
using System;

namespace NPOI.Extensions
{
    public static class GetValueExtensions
    {
        public static T GetValue<T>(this IRow row, int columnIndex)
        {
            var type = typeof(T);
            var cell = row.GetCell(columnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);

            // null-handling
            if (cell == null)
            {
                if (type.IsNullableType() || type == typeof(string)) 
                    return (T) (object) null;
                else
                    throw new NotSupportedException($"Can't assign null value into '{type.FullName}'!");
            }

            // string handling
            if (type == typeof(string))
            {
                // regular string entry
                if (cell.CellType == CellType.String)
                    return (T)(object)cell.StringCellValue;

                // formula with string result
                if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.String)
                    return (T)(object)cell.StringCellValue;

                // format value as string
                var formatter = new DataFormatter();
                var formattedValue = formatter.FormatCellValue(cell);

                // replace non-breakign space with regular space (nbsp were uses with numeric formatting)
                formattedValue = formattedValue.Replace(Convert.ToChar(160), ' ');

                return (T) (object) formattedValue;
            }

            // fetching other value types
            if (type == typeof(char)) return (T)(object)(char)cell.NumericCellValue;
            else if (type.IsNumeric()) return (T)Convert.ChangeType(cell.NumericCellValue, type);
            else if (type.IsDateTime()) return (T)(object)cell.DateCellValue;
            else if (type.IsTimeSpan()) return (T)(object) TimeSpan.FromDays(cell.NumericCellValue);
            else if (type.IsBool()) return (T)(object)cell.BooleanCellValue;
            else throw new NotSupportedException($"Can't handle type'{type.FullName}'!");
        }

    }
}
