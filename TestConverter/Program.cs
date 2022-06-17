using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ConverterExsel;

namespace TestConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ConverterExsel.ExcelConverter test = new ExcelConverter();
            //string test_ansver = test.Converter("Port 8, Subchannel 0 - Пиранометр (СФ-06)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 10)/*DateTime.UtcNow*/, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test");
            //Console.WriteLine(test_ansver);

            //string test_ansver_2 = test.Converter("Порт 8 - Датчики аналоговые ZONE", new DateTime(2022, 05, 06), new DateTime(2022, 05, 25)/*DateTime.UtcNow*/, "ArchivePob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_2");
            //Console.WriteLine(test_ansver_2);

            //string test_ansver_3 = test.Converter("Port 12 - Датчики цифровые  (ZONE)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 12)/*DateTime.UtcNow*/, "ArchiveXPob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_3");
            //Console.WriteLine(test_ansver_3);

            //string test_ansver_4 = test.Converter("Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 13)/*DateTime.UtcNow*/, "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_4", "M05");
            //Console.WriteLine(test_ansver_4);

            //string test_ansver_5 = test.Converter("Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 13)/*DateTime.UtcNow*/, "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_5", "M05");
            //Console.WriteLine(test_ansver_5);

            string test_ansver_6 = test.Converter("Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 13)/*DateTime.UtcNow*/, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_6", "M05");
            Console.WriteLine(test_ansver_6);

            Console.ReadLine();
        }
    }
}
