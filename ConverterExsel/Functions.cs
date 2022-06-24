using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ConverterExsel
{
    class Functions
    {
        public static Dictionary<DateTime, string> GetSuitableData(string pathBegin, DateTime dateBegin, DateTime dateEnd, string parsingFormat, string stationIndex = "M05", string additionalParsingFormat = "")
        {
            Dictionary<DateTime, string> allRelevantData = new Dictionary<DateTime, string>();

            switch (parsingFormat)
            {
                case "ActinometryArchive":
                case "ArchiveXPob":
                case "ArchivePob":
                    {
                        string pathToFiles = Directory.GetCurrentDirectory() + "\\DATA\\" + pathBegin + "\\";
                        foreach (string path in Directory.GetFiles(pathToFiles))
                        {
                            string[] fileName = path.Replace(pathToFiles, "").Split(new char[] { '-', '.' });

                            DateTime dataFromName = new DateTime(Convert.ToInt32(fileName[0]), Convert.ToInt32(fileName[1]), Convert.ToInt32(fileName[2]));
                            if (dataFromName >= dateBegin && dataFromName <= dateEnd)
                            {
                                StreamReader reader = new StreamReader(path);
                                allRelevantData.Add(dataFromName, reader.ReadToEnd());
                            }
                        }
                    }
                    break;
                case "MINpel":
                case "Hpel":
                    {
                        string pathToFolder = Directory.GetCurrentDirectory() + "\\DATA\\" + pathBegin + "\\";
                        foreach (string path in Directory.GetDirectories(pathToFolder))
                        {
                            string folderName = path.Replace(pathToFolder, "");
                            if(folderName.Substring(0, 3) == stationIndex)
                            {
                                DateTime folderData = new DateTime(Convert.ToInt32(folderName.Substring(3, 4)), Convert.ToInt32(folderName.Substring(7, 2)), 1);

                                if (folderData.Year >= dateBegin.Year && folderData.Month >= dateBegin.Month && folderData.Year <= dateEnd.Year && folderData.Month <= dateEnd.Month)
                                {
                                    foreach (string pathToFiles in Directory.GetFiles(path + "\\" + parsingFormat + "\\"))
                                    {
                                        string[] fileName = pathToFiles.Replace(path + "\\" + parsingFormat + "\\", "").Split(new char[] { '-', '.' });
                                        DateTime dataFromName = new DateTime(Convert.ToInt32(fileName[0]), Convert.ToInt32(fileName[1]), Convert.ToInt32(fileName[2]));
                                        if (dataFromName >= dateBegin && dataFromName <= dateEnd)
                                        {
                                            StreamReader reader = new StreamReader(pathToFiles);
                                            allRelevantData.Add(dataFromName, reader.ReadToEnd());
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "VODpel":
                    {
                        string pathToFolder = Directory.GetCurrentDirectory() + "\\DATA\\" + pathBegin + "\\";
                        int i = 1;
                        foreach (string path in Directory.GetDirectories(pathToFolder))
                        {
                            string folderName = path.Replace(pathToFolder, "");
                            if (folderName.Substring(0, 3) == stationIndex)
                            {
                                DateTime folderData = new DateTime(Convert.ToInt32(folderName.Substring(3, 4)), Convert.ToInt32(folderName.Substring(7, 2)), 1);

                                if (folderData.Year >= dateBegin.Year && folderData.Month >= dateBegin.Month && folderData.Year <= dateEnd.Year && folderData.Month <= dateEnd.Month)
                                {
                                    foreach (string pathToFiles in Directory.GetFiles(path + "\\" + parsingFormat + "\\"))
                                    {
                                        string[] fileName = pathToFiles.Replace(path + "\\" + parsingFormat + "\\", "").Split(new char[] { '-', '.' });
                                        //DateTime dataFromName = new DateTime(Convert.ToInt32(fileName[0]), Convert.ToInt32(fileName[1]), Convert.ToInt32(fileName[2]));
                                        //if (dataFromName >= dateBegin && dataFromName <= dateEnd)
                                        //{
                                        
                                        StreamReader reader = new StreamReader(pathToFiles);
                                        if(allRelevantData.ContainsKey(folderData))
                                        {
                                            folderData = folderData.AddMilliseconds(i);
                                            i++;
                                        }
                                        allRelevantData.Add(folderData, reader.ReadToEnd());
                                        //}
                                    }
                                }
                            }
                        }
                    }
                    break;
                default:
                    break;
            }

            return allRelevantData;
        }
        public static List<object> GetSuitableCollection(string parsingFormat, Dictionary<DateTime, string> dataFile,string additionalParsingFormat = "")
        {
            List<object> myCollection = new List<object>();
            switch (parsingFormat)
            {
                case "ArchivePob":
                    {
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
                                        switch (words[i].Substring(0, words[i].Length - 1))
                                        {
                                            case "RadiationS":
                                                archivePob.RadiationS = words[i + 1].Substring(0, words[i + 1].Length - (i == (words.Length - 2) ? 2 : 1));
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
                    }
                    break;
                case "ArchiveXPob":
                    {
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
                                    if (words[2].Contains("1["))
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
                    }
                    break;
                case "ActinometryArchive":
                    {
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
                    }
                    break;
                case "Hpel":
                case "MINpel":
                    {
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
                                    if (additionalParsingFormat == "ActinometryArchive" || additionalParsingFormat == "ArchiveXPob")
                                    {
                                        AuxiliaryFiles.ActinometryArchive HpelMINpel = new AuxiliaryFiles.ActinometryArchive();
                                        var words = lineOfFile.Split(new char[] { ' ', ';' });
                                        var time = words[1].Substring(0, words[1].Length - 1).Split(new char[] { ':' });
                                        HpelMINpel.Time = file.Key.AddHours(Convert.ToInt32(time[0])).AddMinutes(Convert.ToInt32(time[1]));
                                        if (words.Length == 7)
                                        {
                                            HpelMINpel.Radiation = words[3];
                                            HpelMINpel.Millivolt = words[5];
                                        }
                                        else
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
                    }
                    break;

                case "VODpel":
                    {
                        foreach (var file in dataFile)
                        {
                            var fileData = file.Value.Split(new char[] { '\n' });
                            int day = 0, column = 1;
                            foreach (string lineOfFile in fileData)
                            {
                                if (System.Text.RegularExpressions.Regex.IsMatch(lineOfFile, @"[a-zA-Z]{1,3}\r", System.Text.RegularExpressions.RegexOptions.Compiled))
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
                    }
                    break;

                default:
                    break;


            }




            return myCollection;
        }
    }
}
