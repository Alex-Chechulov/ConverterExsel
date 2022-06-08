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
        public string Converter(string pathBegin, DateTime dateBegin, DateTime dateEnd, string parsingFormat, string pathEnd, string name)
        {
            string answer = null;

   
            switch (parsingFormat)
            {
                case "ArchivePob":
                    {

                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd);
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
                                exsel.Set(column: "A", row: 1, data: "Date Time");
                                exsel.Set(column: "B", row: 1, data: "RadiationS, Вт/м2");
                                exsel.Set(column: "C", row: 1, data: "RadiationD, Вт/м2");
                                exsel.Set(column: "D", row: 1, data: "RadiationQ, Вт/м2");
                                exsel.Set(column: "E", row: 1, data: "RadiationR, Вт/м2");
                                exsel.Set(column: "F", row: 1, data: "RadiationB, Вт/м2");
                                exsel.Set(column: "G", row: 1, data: "RadiationQk, Вт/м2");
                                exsel.Set(column: "H", row: 1, data: "RadiationQet, Вт/м2");
                                exsel.Set(column: "I", row: 1, data: "RadiationX, Вт/м2");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    exsel.Set(column: "A", row: i + 2, data: myCollection[i].Time);
                                    exsel.Set(column: "B", row: i + 2, data: myCollection[i].RadiationS);
                                    exsel.Set(column: "C", row: i + 2, data: myCollection[i].RadiationD);
                                    exsel.Set(column: "D", row: i + 2, data: myCollection[i].RadiationQ);
                                    exsel.Set(column: "E", row: i + 2, data: myCollection[i].RadiationR);
                                    exsel.Set(column: "F", row: i + 2, data: myCollection[i].RadiationB);
                                    exsel.Set(column: "G", row: i + 2, data: myCollection[i].RadiationQk);
                                    exsel.Set(column: "H", row: i + 2, data: myCollection[i].RadiationQet);
                                    exsel.Set(column: "I", row: i + 2, data: myCollection[i].RadiationX);
                                }
                                exsel.Save();
                            }
                        }
                        answer = "Convertation sucsess";
                    }

                    break;


                case "ArchiveXPob":

                    break;


                case "ActinometryArchive":
                    {
                        Dictionary<DateTime, string> dataFile = Functions.GetSuitableData(pathBegin, dateBegin, dateEnd);
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
                                exsel.Set(column: "A", row: 1, data: "Date Time");
                                exsel.Set(column: "B", row: 1, data: "RadiationS, Вт/м2");
                                exsel.Set(column: "C", row: 1, data: "Millivolt, мВ");
                                for (int i = 0; i < myCollection.Count; i++)
                                {
                                    exsel.Set(column: "A", row: i + 2, data: myCollection[i].Time);
                                    exsel.Set(column: "B", row: i + 2, data: myCollection[i].Radiation);
                                    exsel.Set(column: "C", row: i + 2, data: myCollection[i].Millivolt);
                                }
                                exsel.Save();
                            }
                        }
                        answer = "Convertation sucsess";
                    }
                    break;


                case "MINpel":
                case "Hpel":

                    break;


                case "VODpel":

                    break;

                default:
                    break;
            }


            return answer;
        }
    }
}
