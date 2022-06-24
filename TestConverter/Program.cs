using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ConverterExsel;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Drawing;

namespace TestConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ConverterExsel.ExcelConverter test = new ExcelConverter();
            ConverterExsel.PDFConverter test2 = new PDFConverter();
            //string test_ansver = test.Converter("Port 8, Subchannel 0 - Пиранометр (СФ-06)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 10), "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test");
            //Console.WriteLine(test_ansver);

            //string test_ansver_2 = test.Converter("Порт 8 - Датчики аналоговые ZONE", new DateTime(2022, 05, 06), new DateTime(2022, 05, 25), "ArchivePob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_2");
            //Console.WriteLine(test_ansver_2);

            //string test_ansver_3 = test.Converter("Port 12 - Датчики цифровые  (ZONE)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 12), "ArchiveXPob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_3");
            //Console.WriteLine(test_ansver_3);

            //string test_ansver_4 = test.Converter("Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 13), "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_4", "M05", "ActinometryArchive");
            //Console.WriteLine(test_ansver_4);

            //string test_ansver_5 = test.Converter("Port 12 - Датчики цифровые  (ZONE)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 13), "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_5", "M05", "ArchiveXPob");
            //Console.WriteLine(test_ansver_5);

            //string test_ansver_6 = test.Converter("Порт 8 - Датчики аналоговые ZONE", new DateTime(2022, 05, 06), new DateTime(2022, 05, 25), "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_6", "M05", "ArchivePob");
            //Console.WriteLine(test_ansver_6);

            //string test_ansver_7 = test.Converter("Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 13), "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_7", "M05");
            //Console.WriteLine(test_ansver_7);

            //string test_ansver_8 = test.Converter("Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_8", "M05");
            //Console.WriteLine(test_ansver_8);

            //tester();
            //PDF_conv();
            //SavePDF();

            //string test_ansver_9 = test2.Converter("Port 12 - Датчики цифровые  (ZONE)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 12), "ArchiveXPob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_9");
            //Console.WriteLine(test_ansver_9);

            //string test_ansver_10 = test2.Converter("Port 8, Subchannel 0 - Пиранометр (СФ-06)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 10), "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_10");
            //Console.WriteLine(test_ansver_10);

            //string test_ansver_11 = test2.Converter("Порт 8 - Датчики аналоговые ZONE", new DateTime(2022, 05, 06), new DateTime(2022, 05, 25), "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_11", "M05", "ArchivePob");
            //Console.WriteLine(test_ansver_11);

            //string test_ansver_12 = test2.Converter("Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "ArchivePob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test_12");
            //Console.WriteLine(test_ansver_12);

            tester_2();

            Console.ReadLine();
        }

        static void tester()
        {
            List<List<object>> test = new List<List<object>>() {
                new List<object>() { "Port 8, Subchannel 0 - Пиранометр (СФ-06)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },
                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },
                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },
                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 8, Subchannel 2 - Актинометр (СФ-12)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 8, Subchannel 2 - Актинометр СФ-12", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 10, Subchannel 0 - Пиранометр (S)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 10, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 10, Subchannel 2 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "ArchiveXPob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },
                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },
                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },
                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },

                new List<object>() { "Port 12 - Цифровой пиранометр СФ-06", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 12, Channel 1 - I_Пиранометр (SF-06)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 12, Channel 1 - Цифровой пиранометр СФ-06", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Порт 6 - Балансомер СФ-08", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "ArchivePob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },
                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },
                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },
                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },

                new List<object>() { "Порт 12, Идентификатор 1 - Цифровой пиранометр СФ-06", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" }
            };
            ConverterExsel.ExcelConverter converter = new ExcelConverter();
            for (int i = 0; i < test.Count; i++)
            {
                string test_ansver = converter.Converter((string)test[i][0], (DateTime)test[i][1], (DateTime)test[i][2], (string)test[i][3], (string)test[i][4], (string)test[i][0] + i/*(string)test[i][5]+"_"+i*/, (string)test[i][6], (string)test[i][7]);

                //Console.WriteLine("Continue?");
                //Console.ReadLine();
                //System.Threading.Thread.Sleep(5000);
                Console.WriteLine((string)test[i][0] + ": " + test_ansver);
            }
            Console.WriteLine("test compliate");
        }
        static void tester_2()
        {
            List<List<object>> test = new List<List<object>>() {
                new List<object>() { "Port 8, Subchannel 0 - Пиранометр (СФ-06)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },
                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },
                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },
                new List<object>() { "Port 8, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 8, Subchannel 2 - Актинометр (СФ-12)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 8, Subchannel 2 - Актинометр СФ-12", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 10, Subchannel 0 - Пиранометр (S)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 10, Subchannel 1 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 10, Subchannel 2 - Балансомер (СФ-08)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "ArchiveXPob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },
                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },
                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },
                new List<object>() { "Port 12 - Датчики цифровые  (ZONE)", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchiveXPob" },

                new List<object>() { "Port 12 - Цифровой пиранометр СФ-06", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 12, Channel 1 - I_Пиранометр (SF-06)", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Port 12, Channel 1 - Цифровой пиранометр СФ-06", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Порт 6 - Балансомер СФ-08", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" },

                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "ArchivePob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },
                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "MINpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },
                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "Hpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },
                new List<object>() { "Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob" },

                new List<object>() { "Порт 12, Идентификатор 1 - Цифровой пиранометр СФ-06", new DateTime(2021, 05, 06), DateTime.Now, "ActinometryArchive", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ActinometryArchive" }
            };
            ConverterExsel.PDFConverter converter = new PDFConverter();
            for (int i = 0; i < test.Count; i++)
            {
                string test_ansver = converter.Converter((string)test[i][0], (DateTime)test[i][1], (DateTime)test[i][2], (string)test[i][3], (string)test[i][4], (string)test[i][0] + i/*(string)test[i][5]+"_"+i*/, (string)test[i][6], (string)test[i][7]);

                //Console.WriteLine("Continue?");
                //Console.ReadLine();
                //System.Threading.Thread.Sleep(5000);
                Console.WriteLine((string)test[i][0] + ": " + test_ansver);
            }
            Console.WriteLine("test compliate");
        }
    }
}
