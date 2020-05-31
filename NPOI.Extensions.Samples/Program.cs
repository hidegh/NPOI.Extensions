using System;

namespace NPOI.Extensions.Samples
{
    class Program
    {
        static void Main(string[] args)
        {
            new MapExtensions.Sample().Run_MapToClass();

            Console.WriteLine();
            Console.WriteLine("Press any key to exit!");
            Console.ReadKey();
        }
    }
}
