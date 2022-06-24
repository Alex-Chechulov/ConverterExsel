using System;
using System.Collections.Generic;
using System.IO;

namespace ConverterExsel
{
    public class ExcelConverter
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

                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")))
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "RadiationS, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "RadiationD, Вт/м2");
                                exsel.Set(column: 4, row: 1, data: "RadiationQ, Вт/м2");
                                exsel.Set(column: 5, row: 1, data: "RadiationR, Вт/м2");
                                exsel.Set(column: 6, row: 1, data: "RadiationB, Вт/м2");
                                exsel.Set(column: 7, row: 1, data: "RadiationQk, Вт/м2");
                                exsel.Set(column: 8, row: 1, data: "RadiationQet, Вт/м2");
                                exsel.Set(column: 9, row: 1, data: "RadiationX, Вт/м2");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    if (myCollection[i] is AuxiliaryFiles.ArchivePob Collection)
                                    {
                                        exsel.Set(column: 1, row: i + 2, data: Collection.Time);
                                        exsel.Set(column: 2, row: i + 2, data: Collection.RadiationS);
                                        exsel.Set(column: 3, row: i + 2, data: Collection.RadiationD);
                                        exsel.Set(column: 4, row: i + 2, data: Collection.RadiationQ);
                                        exsel.Set(column: 5, row: i + 2, data: Collection.RadiationR);
                                        exsel.Set(column: 6, row: i + 2, data: Collection.RadiationB);
                                        exsel.Set(column: 7, row: i + 2, data: Collection.RadiationQk);
                                        exsel.Set(column: 8, row: i + 2, data: Collection.RadiationQet);
                                        exsel.Set(column: 9, row: i + 2, data: Collection.RadiationX);
                                    }
                                }
                                exsel.Save();
                            }
                        }
                        answer = "Convertation ArchivePob sucsess";
                    }

                    break;


                case "ArchiveXPob":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile);
                        
                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")))
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "Radiation, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "Millivolt, Вт/м2");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    if (myCollection[i] is AuxiliaryFiles.ArchiveXPob Collection)
                                    {
                                        exsel.Set(column: 1, row: i + 2, data: Collection.Time);
                                        exsel.Set(column: 2, row: i + 2, data: Collection.Radiation);
                                        exsel.Set(column: 3, row: i + 2, data: Collection.Millivolt);
                                    }
                                }
                                exsel.Save();
                            }
                        }
                        answer = "Convertation ArchiveXPob sucsess";
                    }

                    break;


                case "ActinometryArchive":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile);

                        //Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        //List<AuxiliaryFiles.ActinometryArchive> myCollection = new List<AuxiliaryFiles.ActinometryArchive>();
                        
                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")))
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "Radiation, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "Millivolt, мВ");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    if (myCollection[i] is AuxiliaryFiles.ActinometryArchive Collection)
                                    {
                                        exsel.Set(column: 1, row: i + 2, data: Collection.Time);
                                        exsel.Set(column: 2, row: i + 2, data: Collection.Radiation);
                                        exsel.Set(column: 3, row: i + 2, data: Collection.Millivolt);
                                    }
                                }
                                exsel.Save();
                            }
                        }
                        answer = "Convertation ActinometryArchive sucsess";
                    }
                    break;


                case "Hpel":
                case "MINpel":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile, additionalParsingFormat);

                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")) && myCollection[0] is AuxiliaryFiles.ActinometryArchive)
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "RadiationD, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "Millivolt, мВ");

                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    if (myCollection[i] is AuxiliaryFiles.ActinometryArchive aee)
                                    {
                                        exsel.Set(column: 1, row: i + 2, data: aee.Time);
                                        exsel.Set(column: 2, row: i + 2, data: aee.Radiation);
                                        exsel.Set(column: 3, row: i + 2, data: aee.Millivolt);
                                    }
                                }
                                exsel.Save();
                            }
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")) && myCollection[0] is AuxiliaryFiles.Pob_HpelMINpel)
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "RadiationS, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "RadiationD, Вт/м2");
                                exsel.Set(column: 4, row: 1, data: "RadiationQ, Вт/м2");
                                exsel.Set(column: 5, row: 1, data: "RadiationR, Вт/м2");
                                exsel.Set(column: 6, row: 1, data: "RadiationB, Вт/м2");
                                exsel.Set(column: 7, row: 1, data: "RadiationQk, Вт/м2");
                                exsel.Set(column: 8, row: 1, data: "RadiationQet, Вт/м2");
                                exsel.Set(column: 9, row: 1, data: "RadiationX, Вт/м2");
                                exsel.Set(column: 10, row: 1, data: "MillivoltS, мВ");
                                exsel.Set(column: 11, row: 1, data: "MillivoltD, мВ");
                                exsel.Set(column: 12, row: 1, data: "MillivoltQ, мВ");
                                exsel.Set(column: 13, row: 1, data: "MillivoltR, мВ");
                                exsel.Set(column: 14, row: 1, data: "MillivoltB, мВ");
                                exsel.Set(column: 15, row: 1, data: "MillivoltQk, мВ");
                                exsel.Set(column: 16, row: 1, data: "MillivoltQet, мВ");
                                exsel.Set(column: 17, row: 1, data: "MillivoltX, мВ");

                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    if (myCollection[i] is AuxiliaryFiles.Pob_HpelMINpel aee)
                                    {
                                        exsel.Set(column: 1, row: i + 2, data: aee.Time);
                                        exsel.Set(column: 2, row: i + 2, data: aee.RadiationS);
                                        exsel.Set(column: 3, row: i + 2, data: aee.RadiationD);
                                        exsel.Set(column: 4, row: i + 2, data: aee.RadiationQ);
                                        exsel.Set(column: 5, row: i + 2, data: aee.RadiationR);
                                        exsel.Set(column: 6, row: i + 2, data: aee.RadiationB);
                                        exsel.Set(column: 7, row: i + 2, data: aee.RadiationQk);
                                        exsel.Set(column: 8, row: i + 2, data: aee.RadiationQet);
                                        exsel.Set(column: 9, row: i + 2, data: aee.RadiationX);
                                        exsel.Set(column: 10, row: i + 2, data: aee.MillivoltS);
                                        exsel.Set(column: 11, row: i + 2, data: aee.MillivoltD);
                                        exsel.Set(column: 12, row: i + 2, data: aee.MillivoltQ);
                                        exsel.Set(column: 13, row: i + 2, data: aee.MillivoltR);
                                        exsel.Set(column: 14, row: i + 2, data: aee.MillivoltB);
                                        exsel.Set(column: 15, row: i + 2, data: aee.MillivoltQk);
                                        exsel.Set(column: 16, row: i + 2, data: aee.MillivoltQet);
                                        exsel.Set(column: 17, row: i + 2, data: aee.MillivoltX);
                                    }
                                }
                                exsel.Save();
                            }
                        }
                        answer = "Convertation HpelMINpel sucsess";
                    }
                    break;


                case "VODpel":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat);
                        List<object> myCollection = Functions.GetSuitableCollection(parsingFormat, dataFile, additionalParsingFormat);

                        
                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")))
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "Radiation, Вт/м2");
                                //exsel.Set(column: 3, row: 1, data: "Millivolt, мВ");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    if (myCollection[i] is AuxiliaryFiles.VODpel Collection)
                                    {
                                        exsel.Set(column: 1, row: i + 2, data: Collection.Time);
                                        exsel.Set(column: 2, row: i + 2, data: Collection.Radiation1);
                                    }
                                    //exsel.Set(column: 3, row: i + 2, data: myCollection[i].Radiation2);
                                }
                                exsel.Save();
                            }
                        }
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
