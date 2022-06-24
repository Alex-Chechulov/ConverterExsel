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

            tester();

            //PDF_conv();
            //SavePDF();

            Console.ReadLine();
        }

        static void tester()
        {
            List<List <object>> test = new List<List<object>>() { 
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
            for(int i = 0; i < test.Count; i++)
            {
                string test_ansver = converter.Converter((string)test[i][0], (DateTime)test[i][1], (DateTime)test[i][2], (string)test[i][3], (string)test[i][4], (string)test[i][0]+i/*(string)test[i][5]+"_"+i*/, (string)test[i][6], (string)test[i][7]);

                //Console.WriteLine("Continue?");
                //Console.ReadLine();
                //System.Threading.Thread.Sleep(5000);
                Console.WriteLine((string)test[i][0] + ": "+test_ansver);
            }
            Console.WriteLine("test compliate");
        }
        static void PDF_conv()
        {
            var doc = new Document();
            PdfWriter.GetInstance(doc, new FileStream("Document.pdf", FileMode.Create, FileAccess.Write));

            doc.Open();
            PdfPTable table = new PdfPTable(4);
            PdfPCell cell = new PdfPCell(new Phrase("Simple table",
              new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 16,
              iTextSharp.text.Font.NORMAL, new BaseColor(Color.Orange))));
            cell.BackgroundColor = new BaseColor(Color.Wheat);
            cell.Padding = 5;
            cell.Colspan = 4;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(cell);
            table.AddCell("Col 1 Row 1");
            table.AddCell("Col 2 Row 1");
            table.AddCell("Col 3 Row 1");
            table.AddCell("Col 1 Row 2");
            table.AddCell("Col 2 Row 2");
            table.AddCell("Col 3 Row 2");
            table.AddCell("Col 2 Row 2");
            table.AddCell("Col 3 Row 2");
            cell.Padding = 5;
            cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            table.AddCell(cell);
            cell = new PdfPCell(new Phrase("Col 2 Row 3"));
            cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            table.AddCell(cell);
            cell.Padding = 5;
            cell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
            table.AddCell(cell);
            doc.Add(table);
            doc.Close();
        }
    }
}
