using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Linq;
using System.IO;

namespace NPOI.Extensions.Samples.MapExtensions
{
    public class Sample
    {
        protected const string testExcelFileName = "./MapExtensions/Sample.xlsx";

        public class SheetColumnTitles
        {
            public const string Date = "Order Date";
            public const string Region = "Region";
            public const string Representative = "Representative";
            public const string Item = "Item";
            public const string Units = "Units";
            public const string UnitCost = "Unit Cost";
            public const string Total = "Total";
        }

        public enum RegionEnum
        {
            East,
            Central,
            West
        }

        public class OrderDetails
        {
            public DateTime Date { get; set; }
            public RegionEnum Region { get; set; }
            public string Representative { get; set; }
            public string Item { get; set; }
            public int Units { get; set; }
            public decimal UnitCost { get; set; }
            public decimal Total { get; set; }

            public string DateFormatted { get; set; }
            public string TotalFormatted { get; set; }
        }

        public void Run_MapToClass()
        {
            IWorkbook workbook;

            using (var fileReader = new FileStream(testExcelFileName, FileMode.Open, FileAccess.Read))
                workbook = new XSSFWorkbook(fileReader);

            var sheet = workbook.GetSheetAt(0);

            // reusable mapping function from value to enum...
            var regionMapper = (Func<string, RegionEnum>)((string region) =>
            {
                switch ((region ?? "").ToUpper())
                {
                    case "EAST": return RegionEnum.East;
                    case "CENTRAL": return RegionEnum.Central;
                    case "WEST": return RegionEnum.West;
                    default: throw new NotSupportedException();
                }
            });

            // map to list
            var data = sheet.MapTo<OrderDetails>(true, rowMapper =>
            {
                return new OrderDetails()
                {
                    Date = rowMapper.GetValue<DateTime>(SheetColumnTitles.Date),
                    Region = regionMapper(rowMapper.GetValue<string>(SheetColumnTitles.Region)),
                    Representative = rowMapper.GetValue<string>(SheetColumnTitles.Representative),
                    Item = rowMapper.GetValue<string>(SheetColumnTitles.Item),
                    Units = rowMapper.GetValue<int>(SheetColumnTitles.Units),
                    UnitCost = rowMapper.GetValue<decimal>(SheetColumnTitles.UnitCost),
                    Total = rowMapper.GetValue<decimal>(SheetColumnTitles.Total),

                    DateFormatted = rowMapper.GetValue<string>(SheetColumnTitles.Date),
                    TotalFormatted = rowMapper.GetValue<string>(SheetColumnTitles.Total)
                };
            });

            // write to console
            Console.WriteLine("Date       | Region  | Representative  | Item            | Units | Unit cost  | Total      |");
            Console.WriteLine("-------------------------------------------------------------------------------------------- ");

            data.ForEach(data =>
            {
                Console.WriteLine("{0:yyyy-MM-dd} | {1,-7} | {2,-15} | {3,-15} | {4,5:N0} | {5,10:N2} | {6,10:N2} | {7,10}, {8}", data.Date, data.Region, data.Representative, data.Item, data.Units, data.UnitCost, data.Total, data.DateFormatted, data.TotalFormatted);
            });

        }

    }

}
