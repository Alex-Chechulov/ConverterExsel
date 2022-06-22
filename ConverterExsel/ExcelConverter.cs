using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
                        List<AuxiliaryFiles.ArchivePob> myCollection = new List<AuxiliaryFiles.ArchivePob>();
                        foreach (var file in dataFile)
                        {
                            var fileData = file.Value.Split(new char[] { '\n' });
                            foreach (string lineOfFile in fileData)
                            {
                                if (lineOfFile != "")
                                {
                                    AuxiliaryFiles.ArchivePob archivePob = new AuxiliaryFiles.ArchivePob();
                                    var words = lineOfFile.Split(new char[] { ' ' });
                                    var time = words[1].Substring(1, words[1].Length - 3).Split(new char[] { ':' });
                                    archivePob.Time = file.Key.AddHours(Convert.ToInt32(time[0])).AddMinutes(Convert.ToInt32(time[1])).AddSeconds(Convert.ToInt32(time[2]));

                                    words[2] = words[2].Substring(1);
                                    for (int i = 0; i < words.Length; i++)
                                    {
                                        switch(words[i].Substring(0, words[i].Length - 1))
                                        {
                                            case "RadiationS":
                                                archivePob.RadiationS = words[i+1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            case "RadiationD":
                                                archivePob.RadiationD = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            case "RadiationQ":
                                                archivePob.RadiationQ = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            case "RadiationR":
                                                archivePob.RadiationR = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            case "RadiationB":
                                                archivePob.RadiationB = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            case "RadiationQk":
                                                archivePob.RadiationQk = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            case "RadiationQet":
                                                archivePob.RadiationQet = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            case "RadiationX":
                                                archivePob.RadiationX = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
                                                break;
                                            default:
                                                break;

                                        }
                                    };
                                    myCollection.Add(archivePob);
                                }
                            }

                        }
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
                                    exsel.Set(column: 1, row: i + 2, data: myCollection[i].Time);
                                    exsel.Set(column: 2, row: i + 2, data: myCollection[i].RadiationS);
                                    exsel.Set(column: 3, row: i + 2, data: myCollection[i].RadiationD);
                                    exsel.Set(column: 4, row: i + 2, data: myCollection[i].RadiationQ);
                                    exsel.Set(column: 5, row: i + 2, data: myCollection[i].RadiationR);
                                    exsel.Set(column: 6, row: i + 2, data: myCollection[i].RadiationB);
                                    exsel.Set(column: 7, row: i + 2, data: myCollection[i].RadiationQk);
                                    exsel.Set(column: 8, row: i + 2, data: myCollection[i].RadiationQet);
                                    exsel.Set(column: 9, row: i + 2, data: myCollection[i].RadiationX);
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
                        List<AuxiliaryFiles.ArchiveXPob> myCollection = new List<AuxiliaryFiles.ArchiveXPob>();
                        foreach (var file in dataFile)
                        {
                            var fileData = file.Value.Split(new char[] { '\n' });
                            foreach (string lineOfFile in fileData)
                            {
                                if (lineOfFile != "")
                                {
                                    AuxiliaryFiles.ArchiveXPob archiveXPobe = new AuxiliaryFiles.ArchiveXPob();
                                    var words = lineOfFile.Split(new char[] { ' ' });
                                    //if (words[2].Substring(1, words[2].Length - 2) == "1[S]")
                                    if(words[2].Contains("1["))
                                    {
                                        var time = words[1].Substring(1, words[1].Length - 3).Split(new char[] { ':' });
                                        archiveXPobe.Time = file.Key.AddHours(Convert.ToInt32(time[0])).AddMinutes(Convert.ToInt32(time[1])).AddSeconds(Convert.ToInt32(time[2]));
                                        archiveXPobe.Radiation = words[3];
                                        archiveXPobe.Millivolt = words[4].Substring(0, words[4].Length - 2);
                                        myCollection.Add(archiveXPobe);
                                    }
                                }
                            }

                        }
                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")))
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "Radiation, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "Millivolt, Вт/м2");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    exsel.Set(column: 1, row: i + 2, data: myCollection[i].Time);
                                    exsel.Set(column: 2, row: i + 2, data: myCollection[i].Radiation);
                                    exsel.Set(column: 3, row: i + 2, data: myCollection[i].Millivolt);
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
                        List<AuxiliaryFiles.ActinometryArchive> myCollection = new List<AuxiliaryFiles.ActinometryArchive>();
                        foreach (var file in dataFile)
                        {
                            var fileData = file.Value.Split(new char[] { '\n' });
                            foreach (string lineOfFile in fileData)
                            {
                                if (lineOfFile != "")
                                {
                                    AuxiliaryFiles.ActinometryArchive actinometryArchive = new AuxiliaryFiles.ActinometryArchive();
                                    var words = lineOfFile.Split(new char[] { ' ' });
                                    var time = words[1].Substring(1, words[1].Length - 3).Split(new char[] { ':' });
                                    actinometryArchive.Time = file.Key.AddHours(Convert.ToInt32(time[0])).AddMinutes(Convert.ToInt32(time[1])).AddSeconds(Convert.ToInt32(time[2]));
                                    actinometryArchive.Radiation = words[3].Substring(0, words[3].Length - 1);
                                    actinometryArchive.Millivolt = words[5].Substring(0, words[5].Length - 2);
                                    myCollection.Add(actinometryArchive);
                                }
                            }

                        }
                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")))
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "Radiation, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "Millivolt, мВ");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    exsel.Set(column: 1, row: i + 2, data: myCollection[i].Time);
                                    exsel.Set(column: 2, row: i + 2, data: myCollection[i].Radiation);
                                    exsel.Set(column: 3, row: i + 2, data: myCollection[i].Millivolt);
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
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat, stationIndex, additionalParsingFormat);
                        //List<AuxiliaryFiles.HpelMINpel> myCollection = new List<AuxiliaryFiles.HpelMINpel>();
                        List<object> myCollection = new List<object>();

                        //myCollection__.Add(new AuxiliaryFiles.HpelMINpel());


                        foreach (var file in dataFile)
                        {
                            int limitEnd = 1;
                            int limitStart = 0;
                            var fileData = file.Value.Split(new char[] { '\n' });
                            foreach (string lineOfFile in fileData)
                            {
                                if (limitEnd > limitStart++)
                                    continue;
                                if (lineOfFile != "")
                                {

                                    //for (int i = 0; i < words.Length; i++)
                                    //{
                                    //    if (words.Length != 6)
                                    //        Console.WriteLine(words[i]/*.Substring(1, words[2].Length - 2)*/);
                                    //}
                                    //Console.WriteLine("\nnext\n");
                                    if (additionalParsingFormat == "ActinometryArchive"|| additionalParsingFormat == "ArchiveXPob")
                                    {
                                        AuxiliaryFiles.ActinometryArchive HpelMINpel = new AuxiliaryFiles.ActinometryArchive();
                                        var words = lineOfFile.Split(new char[] { ' ', ';' });
                                        var time = words[1].Substring(0, words[1].Length - 1).Split(new char[] { ':' });
                                        HpelMINpel.Time = file.Key.AddHours(Convert.ToInt32(time[0])).AddMinutes(Convert.ToInt32(time[1]));
                                        if (words.Length == 7)
                                        {
                                            HpelMINpel.Radiation = words[3];
                                            HpelMINpel.Millivolt = words[5];
                                        } else
                                        {
                                            HpelMINpel.Radiation = words[2];
                                            HpelMINpel.Millivolt = words[3];
                                        }
                                        myCollection.Add(HpelMINpel);
                                    }
                                    if (additionalParsingFormat == "ArchivePob")
                                    {
                                        AuxiliaryFiles.Pob_HpelMINpel HpelMINpel = new AuxiliaryFiles.Pob_HpelMINpel();
                                        var words = lineOfFile.Split(new char[] { ' ', ';' });
                                        var time = words[1].Substring(0, words[1].Length - 1).Split(new char[] { ':' });
                                        HpelMINpel.Time = file.Key.AddHours(Convert.ToInt32(time[0])).AddMinutes(Convert.ToInt32(time[1]));
                                        HpelMINpel.RadiationD = words[2];
                                        HpelMINpel.RadiationS = words[3];
                                        HpelMINpel.RadiationQ = words[4];
                                        HpelMINpel.RadiationR = words[5];
                                        HpelMINpel.RadiationB = words[6];
                                        HpelMINpel.RadiationQk = words[7];
                                        HpelMINpel.RadiationQet = words[8];
                                        HpelMINpel.RadiationX = words[9];
                                        HpelMINpel.MillivoltD = words[10];
                                        HpelMINpel.MillivoltS = words[11];
                                        HpelMINpel.MillivoltQ = words[12];
                                        HpelMINpel.MillivoltR = words[13];
                                        HpelMINpel.MillivoltB = words[14];
                                        HpelMINpel.MillivoltQk = words[15];
                                        HpelMINpel.MillivoltQet = words[16];
                                        HpelMINpel.MillivoltX = words[17];
                                        myCollection.Add(HpelMINpel);
                                    }
                                }
                            }

                        }
                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx"))&& myCollection[0] is AuxiliaryFiles.ActinometryArchive)
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "RadiationD, Вт/м2");
                                exsel.Set(column: 3, row: 1, data: "Millivolt, мВ");
                                
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    if(myCollection[i] is AuxiliaryFiles.ActinometryArchive aee)
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
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd, parsingFormat, stationIndex);
                        List<AuxiliaryFiles.VODpel> myCollection = new List<AuxiliaryFiles.VODpel>();
                        foreach (var file in dataFile)
                        {
                            var fileData = file.Value.Split(new char[] { '\n' });
                            int day = 0, column = 1;
                            foreach (string lineOfFile in fileData)
                            {
                                if (System.Text.RegularExpressions.Regex.IsMatch(lineOfFile, @"[a-zA-Z]{1,3}\r",System.Text.RegularExpressions.RegexOptions.Compiled))
                                {
                                    day = 0;
                                    column++;
                                    //Console.WriteLine("new");
                                    continue;
                                }
                                //if (lineOfFile != "D\r")
                                //{
                                //if(words[2].Contains("1["))

                                if (lineOfFile != "")
                                {
                                    for (int quarterAnHour = 0; quarterAnHour < lineOfFile.Length - 1;)
                                    {
                                        AuxiliaryFiles.VODpel VODpel = new AuxiliaryFiles.VODpel();
                                        string s_data = lineOfFile.Substring(quarterAnHour, 4).Replace(" ", "0");
                                        int data = Convert.ToInt32(s_data);
                                        VODpel.Time = file.Key.AddHours(quarterAnHour / 4 + 1).AddDays(day);
                                        if (data != 0)
                                        {
                                            VODpel.Radiation1 = Convert.ToString(data);
                                        }
                                        quarterAnHour += 4;
                                        myCollection.Add(VODpel);
                                    }
                                }
                                day++;

                                //}
                            }

                        }
                        using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                        {
                            if (exsel.Open(filePath: Path.Combine(pathEnd, name + ".xlsx")))
                            {
                                exsel.Set(column: 1, row: 1, data: "Date Time");
                                exsel.Set(column: 2, row: 1, data: "Radiation, Вт/м2");
                                //exsel.Set(column: 3, row: 1, data: "Millivolt, мВ");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    exsel.Set(column: 1, row: i + 2, data: myCollection[i].Time);
                                    exsel.Set(column: 2, row: i + 2, data: myCollection[i].Radiation1);
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
