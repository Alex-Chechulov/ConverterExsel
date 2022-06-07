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
                                archivePob.Radiation = words[3].Substring(0, words[3].Length - 1);
                                archivePob.Millivolt = words[5].Substring(0, words[5].Length - 2);
                                myCollection.Add(archivePob);
                            }
                        }

                    }
                    using (AuxiliaryFiles.ExcelHelper exsel = new AuxiliaryFiles.ExcelHelper())
                    {
                        if (exsel.Open(filePath: Path.Combine(pathEnd, name+".xlsx")))
                        {
                            exsel.Set(column: "A", row: 1, data: "Дата время");
                            exsel.Set(column: "B", row: 1, data: "Радиация, Вт/м2");
                            exsel.Set(column: "C", row: 1, data: "Миливольты, мВ");
                            for (int i = 0; i < myCollection.Count; i++)
                            {
                                exsel.Set(column: "A", row: i+2, data: myCollection[i].Time);
                                exsel.Set(column: "B", row: i+2, data: myCollection[i].Radiation);
                                exsel.Set(column: "C", row: i+2, data: myCollection[i].Millivolt);
                            }
                            exsel.Save();
                        }
                    }
                    answer = "Convertation sucsess";
                    break;


                case "ArchiveXPob":

                    break;


                case "ActinometryArchive":

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
