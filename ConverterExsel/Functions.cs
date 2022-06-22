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
    }
}
