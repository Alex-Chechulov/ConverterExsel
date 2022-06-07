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
        public static Dictionary<DateTime, string> GetSuitableData(string pathBegin, DateTime dateBegin, DateTime dateEnd)
        {
            Dictionary<DateTime, string> allRelevantData =new Dictionary<DateTime, string>();
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

            return allRelevantData;
        }
    }
}
