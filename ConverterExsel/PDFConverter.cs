using System;
using System.Collections.Generic;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing;
namespace ConverterExsel
{
    public class PDFConverter
    {
        public string Converter(string pathBegin, DateTime dateBegin, DateTime dateEnd, string parsingFormat, string pathEnd, string name, string stationIndex = "", string additionalParsingFormat = "")
        {
            string answer = null;


            switch (parsingFormat)
            {
                case "ArchivePob":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile);

                        var doc = new Document(PageSize.A4.Rotate(), 10, 10, 10, 10);
                        PdfWriter.GetInstance(doc, new FileStream(Path.Combine(pathEnd, name + ".pdf"), FileMode.Create, FileAccess.Write));

                        doc.Open();
                        PdfPTable table = new PdfPTable(9);
                        table.AddCell("Date Time");
                        table.AddCell("RadiationS, Vt/m2");
                        table.AddCell("RadiationD, Vt/m2");
                        table.AddCell("RadiationQ, Vt/m2");
                        table.AddCell("RadiationR, Vt/m2");
                        table.AddCell("RadiationB, Vt/m2");
                        table.AddCell("RadiationQk, Vt/m2");
                        table.AddCell("RadiationQet, Vt/m2");
                        table.AddCell("RadiationX, Vt/m2");
                        for (int i = 0; i < myCollection.Count; i++)
                        {
                            if (myCollection[i] is AuxiliaryFiles.ArchivePob Collection)
                            {
                                table.AddCell(Convert.ToString(Collection.Time));
                                table.AddCell(Collection.RadiationS);
                                table.AddCell(Collection.RadiationD);
                                table.AddCell(Collection.RadiationQ);
                                table.AddCell(Collection.RadiationR);
                                table.AddCell(Collection.RadiationB);
                                table.AddCell(Collection.RadiationQk);
                                table.AddCell(Collection.RadiationQet);
                                table.AddCell(Collection.RadiationX);
                            }
                        }
                        doc.Add(table);
                        doc.Close();

                        answer = "Convertation ArchivePob sucsess";
                    }

                    break;


                case "ArchiveXPob":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile);

                        var doc = new Document(PageSize.A4, 10, 10, 10, 10);
                        PdfWriter.GetInstance(doc, new FileStream(Path.Combine(pathEnd, name + ".pdf"), FileMode.Create, FileAccess.Write));

                        doc.Open();
                        PdfPTable table = new PdfPTable(3);
                        table.AddCell("Date Time");
                        table.AddCell("Radiation, Vt/m2");
                        table.AddCell("Millivolt, mV");
                        for (int i = 0; i < myCollection.Count; i++)
                        {
                            if (myCollection[i] is AuxiliaryFiles.ArchiveXPob Collection)
                            {
                                table.AddCell(Convert.ToString(Collection.Time));
                                table.AddCell(Collection.Radiation);
                                table.AddCell(Collection.Millivolt);
                            }
                        }
                        doc.Add(table);
                        doc.Close();

                        answer = "Convertation ArchiveXPob sucsess";
                    }

                    break;


                case "ActinometryArchive":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile);

                        var doc = new Document(PageSize.A4, 10, 10, 10, 10);
                        PdfWriter.GetInstance(doc, new FileStream(Path.Combine(pathEnd, name + ".pdf"), FileMode.Create, FileAccess.Write));

                        doc.Open();
                        PdfPTable table = new PdfPTable(3);
                        table.AddCell("Date Time");
                        table.AddCell("Radiation, Vt/m2");
                        table.AddCell("Millivolt, mV");
                        for (int i = 0; i < myCollection.Count; i++)
                        {
                            if (myCollection[i] is AuxiliaryFiles.ActinometryArchive Collection)
                            {
                                table.AddCell(Convert.ToString(Collection.Time));
                                table.AddCell(Collection.Radiation);
                                table.AddCell(Collection.Millivolt);
                            }
                        }
                        doc.Add(table);
                        doc.Close();
                        answer = "Convertation ActinometryArchive sucsess";
                    }
                    break;


                case "Hpel":
                case "MINpel":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile, additionalParsingFormat);

                        if (myCollection[0] is AuxiliaryFiles.ActinometryArchive)
                        {
                            var doc = new Document(PageSize.A4, 10, 10, 10, 10);
                            PdfWriter.GetInstance(doc, new FileStream(Path.Combine(pathEnd, name + ".pdf"), FileMode.Create, FileAccess.Write));

                            doc.Open();
                            PdfPTable table = new PdfPTable(3);
                            table.AddCell("Date Time");
                            table.AddCell("Radiation, Vt/m2");
                            table.AddCell("Millivolt, mV");
                            for (int i = 0; i < myCollection.Count; i++)
                            {
                                if (myCollection[i] is AuxiliaryFiles.ActinometryArchive Collection)
                                {
                                    table.AddCell(Convert.ToString(Collection.Time));
                                    table.AddCell(Collection.Radiation);
                                    table.AddCell(Collection.Millivolt);
                                }
                            }
                            doc.Add(table);
                            doc.Close();
                        }
                        if (myCollection[0] is AuxiliaryFiles.Pob_HpelMINpel)
                        {
                            var doc = new Document(PageSize.A4.Rotate(), 0, 0, 10, 10);
                            PdfWriter.GetInstance(doc, new FileStream(Path.Combine(pathEnd, name + ".pdf"), FileMode.Create, FileAccess.Write));                            
                            doc.Open();
                            PdfPTable table = new PdfPTable(17);

                            table.AddCell("Date Time");
                            table.AddCell("RadiationS, Vt/m2");
                            table.AddCell("RadiationD, Vt/m2");
                            table.AddCell("RadiationQ, Vt/m2");
                            table.AddCell("RadiationR, Vt/m2");
                            table.AddCell("RadiationB, Vt/m2");
                            table.AddCell("RadiationQk, Vt/m2");
                            table.AddCell("RadiationQet, Vt/m2");
                            table.AddCell("RadiationX, Vt/m2");
                            table.AddCell("MillivoltS, mV");
                            table.AddCell("MillivoltD, mV");
                            table.AddCell("MillivoltQ, mV");
                            table.AddCell("MillivoltR, mV");
                            table.AddCell("MillivoltB, mV");
                            table.AddCell("MillivoltQk, mV");
                            table.AddCell("MillivoltQet, mV");
                            table.AddCell("MillivoltX, mV");

                            for (int i = 0; i < myCollection.Count; i++)
                            {
                                if (myCollection[i] is AuxiliaryFiles.Pob_HpelMINpel Collection)
                                {
                                    table.AddCell(Convert.ToString(Collection.Time));
                                    table.AddCell(Collection.RadiationS);
                                    table.AddCell(Collection.RadiationD);
                                    table.AddCell(Collection.RadiationQ);
                                    table.AddCell(Collection.RadiationR);
                                    table.AddCell(Collection.RadiationB);
                                    table.AddCell(Collection.RadiationQk);
                                    table.AddCell(Collection.RadiationQet);
                                    table.AddCell(Collection.RadiationX);
                                    table.AddCell(Collection.MillivoltS);
                                    table.AddCell(Collection.MillivoltD);
                                    table.AddCell(Collection.MillivoltQ);
                                    table.AddCell(Collection.MillivoltR);
                                    table.AddCell(Collection.MillivoltB);
                                    table.AddCell(Collection.MillivoltQk);
                                    table.AddCell(Collection.MillivoltQet);
                                    table.AddCell(Collection.MillivoltX);
                                }
                            }
                            doc.Add(table);
                            doc.Close();
                        }
                        answer = "Convertation HpelMINpel sucsess";
                    }
                    break;


                case "VODpel":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile, additionalParsingFormat);

                        var doc = new Document(PageSize.A4, 10, 10, 10, 10);
                        PdfWriter.GetInstance(doc, new FileStream(Path.Combine(pathEnd, name + ".pdf"), FileMode.Create, FileAccess.Write));

                        doc.Open();
                        PdfPTable table = new PdfPTable(2);
                        table.AddCell("Date Time");
                        table.AddCell("Radiation, Vt/m2");
                        for (int i = 0; i < myCollection.Count; i++)
                        {
                            if (myCollection[i] is AuxiliaryFiles.VODpel Collection)
                            {
                                table.AddCell(Convert.ToString(Collection.Time));
                                table.AddCell(Collection.Radiation1);
                            }
                        }
                        doc.Add(table);
                        doc.Close();

                        answer = "Convertation VODpel sucsess";
                    }

                    break;

                default:
                    break;
            }


            return answer;
        }
    }
}
