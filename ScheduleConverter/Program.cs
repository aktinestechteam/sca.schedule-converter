using System;

namespace ScheduleConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Converter converter = new Converter();
            converter.ConvertSchedule(@"C:\POC\Book1.xlsx",string.Empty);


            Console.ReadKey();
        }
    }
}
