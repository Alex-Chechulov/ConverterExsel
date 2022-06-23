using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ConverterExsel;

//using iTextSharp.text;
//using iTextSharp.text.pdf;
//using iTextSharp.text.pdf.parser;
//using System.Drawing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

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

            //string test_ansver_7 = test.Converter("Порт 8 - Датчики аналоговые ZONE", new DateTime(2021, 05, 06), DateTime.Now, "VODpel", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test", "M05", "ArchivePob");
            //Console.WriteLine(test_ansver_7);

            //tester();

            //PDF_conv();
            SavePDF();

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
        //static void PDF_conv()
        //{
        //    var doc = new Document();
        //    PdfWriter.GetInstance(doc, new FileStream("Document.pdf", FileMode.Create, FileAccess.Write));

        //    doc.Open();
        //    PdfPTable table = new PdfPTable(3);
        //    PdfPCell cell = new PdfPCell(new Phrase("Simple table",
        //      new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 16,
        //      iTextSharp.text.Font.NORMAL, new BaseColor(Color.Orange))));
        //    cell.BackgroundColor = new BaseColor(Color.Wheat);
        //    cell.Padding = 5;
        //    cell.Colspan = 3;
        //    cell.HorizontalAlignment = Element.ALIGN_CENTER;
        //    table.AddCell(cell);
        //    table.AddCell("Col 1 Row 1");
        //    table.AddCell("Col 2 Row 1");
        //    table.AddCell("Col 3 Row 1");
        //    table.AddCell("Col 1 Row 2");
        //    table.AddCell("Col 2 Row 2");
        //    table.AddCell("Col 3 Row 2");
        //    cell.Padding = 5;
        //    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
        //    table.AddCell(cell);
        //    cell = new PdfPCell(new Phrase("Col 2 Row 3"));
        //    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
        //    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
        //    table.AddCell(cell);
        //    cell.Padding = 5;
        //    cell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
        //    table.AddCell(cell);
        //    doc.Add(table);
        //    doc.Close();
        //}

        static void SavePDF()
        {
            //System.IO.FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\" + "First PDF document.pdf", FileMode.Create);
            //// Create an instance of the document class which represents the PDF document itself.  
            //Document document = new Document(PageSize.A4, 25, 25, 30, 30);
            //// Create an instance to the PDF file by creating an instance of the PDF   
            //// Writer class using the document and the filestrem in the constructor.  

            //PdfWriter writer = PdfWriter.GetInstance(document, fs);
            //// Add meta information to the document  
            //document.AddAuthor("Micke Blomquist");
            //document.AddCreator("Sample application using iTextSharp");
            //document.AddKeywords("PDF tutorial education");
            //document.AddSubject("Document subject - Describing the steps creating a PDF document");
            //document.AddTitle("The document title - PDF creation using iTextSharp");
            //// Open the document to enable you to write to the document  
            //document.Open();
            //// Add a simple and wellknown phrase to the document in a flow layout manner  
            //document.Add(new Paragraph("Hello World!"));
            //// Close the document  
            //document.Close();
            //// Close the writer instance  
            //writer.Close();
            //// Always close open filehandles explicity  
            //fs.Close();

            //PdfDocument document = new PdfDocument();
            //PdfPage page = document.AddPage();
            //XGraphics gfx = XGraphics.FromPdfPage(page);
            //XFont font = new XFont("arial", 20);
            //gfx.DrawString("filst", font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.Center);
            //document.Save("D:/Work_25/Projects_VS/C#/ConverterExsel/TestConverter/bin/Debug/1.pdf");
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Table Example";

            for (int p = 0; p < 2; p++)
            {
                // Page Options
                PdfPage pdfPage = document.AddPage();
                pdfPage.Height = 842;//842
                pdfPage.Width = 590;
                //pdfPage.Height = 590;//842
                //pdfPage.Width = 842;

                // Get an XGraphics object for drawing
                XGraphics graph = XGraphics.FromPdfPage(pdfPage);

                // Text format
                XStringFormat format = new XStringFormat();
                format.LineAlignment = XLineAlignment.Near;
                format.Alignment = XStringAlignment.Near;
                var tf = graph;

                XFont fontParagraph = new XFont("Verdana", 8, XFontStyle.Regular);

                // Row elements
                int el_width = 31;

                // page structure options
                double lineHeight = 20;
                int marginLeft = 20;
                int marginTop = 20;

                int el_height = 30;
                int rect_height = 17;

                int interLine_X_1 = 2;
                int interLine_X_2 = 2 * interLine_X_1;

                int offSetX_1 = el_width;
                int offSetX_2 = el_width + el_width;

                XSolidBrush rect_style1 = new XSolidBrush(XColors.LightGray);
                XSolidBrush rect_style2 = new XSolidBrush(XColors.DarkGreen);
                XSolidBrush rect_style3 = new XSolidBrush(XColors.Red);

                for (int i = 0; i < 60; i++)
                {
                    double dist_Y = lineHeight * (i + 1);
                    double dist_Y2 = dist_Y - 2;

                    // header della G
                    if (i == 0)
                    {
                        //graph.DrawRectangle(rect_style2, marginLeft, marginTop, pdfPage.Width - 2 * marginLeft, rect_height);

                        tf.DrawString("column1", fontParagraph, XBrushes.Black,
                                      new XRect(marginLeft, marginTop, el_width, el_height), format);

                        tf.DrawString("column2", fontParagraph, XBrushes.Black,
                                      new XRect(marginLeft + offSetX_1 + interLine_X_1, marginTop, el_width, el_height), format);

                        tf.DrawString("column3", fontParagraph, XBrushes.Black,
                                      new XRect(marginLeft + offSetX_2 + 2 * interLine_X_2, marginTop, el_width, el_height), format);

                        // stampo il primo elemento insieme all'header
                        graph.DrawRectangle(rect_style1, marginLeft, dist_Y2 + marginTop, el_width, rect_height);
                        tf.DrawString("text1", fontParagraph, XBrushes.Black,
                                      new XRect(marginLeft, dist_Y + marginTop, el_width, el_height), format);

                        //ELEMENT 2 - BIG 380
                        graph.DrawRectangle(rect_style1, marginLeft + offSetX_1 + interLine_X_1, dist_Y2 + marginTop, el_width, rect_height);
                        tf.DrawString(
                            "text2",
                            fontParagraph,
                            XBrushes.Black,
                            new XRect(marginLeft + offSetX_1 + interLine_X_1, dist_Y + marginTop, el_width, el_height),
                            format);


                        //ELEMENT 3 - SMALL 80

                        graph.DrawRectangle(rect_style1, marginLeft + offSetX_2 + interLine_X_2, dist_Y2 + marginTop, el_width, rect_height);
                        tf.DrawString(
                            "text3",
                            fontParagraph,
                            XBrushes.Black,
                            new XRect(marginLeft + offSetX_2 + 2 * interLine_X_2, dist_Y + marginTop, el_width, el_height),
                            format);


                    }
                    else
                    {

                        //if (i % 2 == 1)
                        //{
                        //  graph.DrawRectangle(TextBackgroundBrush, marginLeft, lineY - 2 + marginTop, pdfPage.Width - marginLeft - marginRight, lineHeight - 2);
                        //}

                        //ELEMENT 1 - SMALL 80
                        graph.DrawRectangle(rect_style1, marginLeft, marginTop + dist_Y2, el_width, rect_height);
                        tf.DrawString(

                            "text1",
                            fontParagraph,
                            XBrushes.Black,
                            new XRect(marginLeft, marginTop + dist_Y, el_width, el_height),
                            format);

                        //ELEMENT 2 - BIG 380
                        graph.DrawRectangle(rect_style1, marginLeft + offSetX_1 + interLine_X_1, dist_Y2 + marginTop, el_width, rect_height);
                        tf.DrawString(
                            "text2",
                            fontParagraph,
                            XBrushes.Black,
                            new XRect(marginLeft + offSetX_1 + interLine_X_1, marginTop + dist_Y, el_width, el_height),
                            format);


                        //ELEMENT 3 - SMALL 80

                        graph.DrawRectangle(rect_style1, marginLeft + offSetX_2 + interLine_X_2, dist_Y2 + marginTop, el_width, rect_height);
                        tf.DrawString(
                            "text3",
                            fontParagraph,
                            XBrushes.Black,
                            new XRect(marginLeft + offSetX_2 + 2 * interLine_X_2, marginTop + dist_Y, el_width, el_height),
                            format);

                    }

                }


            }


            const string filename = "D:/Work_25/Projects_VS/C#/ConverterExsel/TestConverter/bin/Debug/2.pdf";
            document.Save(filename);

            //byte[] bytes = null;
            //using (MemoryStream stream = new MemoryStream())
            //{
            //    document.Save(stream, true);
            //    bytes = stream.ToArray();
            //}

            //SendFileToResponse(bytes, "HelloWorld_test.pdf");
        }

    }
}
