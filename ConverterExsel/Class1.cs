using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ConverterExsel
{
    public class Class1
    {
        public string converter(string path_begin, DateTime date_begin, DateTime date_end, string parsing_format, string path_end, string name)
        {
            string answer = "";
            string path = Directory.GetCurrentDirectory() + "\\DATA\\" + path_begin+"\\";
            string date_month = date_begin.Month < 10 ? "0" + date_begin.Month : date_begin.Month.ToString();
            string date_day = date_begin.Day < 10 ? "0" + date_begin.Day : date_begin.Day.ToString();
            string file_name = date_begin.Year + "-" +date_month + "-" + date_day + ".YML";
            switch (parsing_format)
            {
                case "ArchivePob":
                    using (StreamReader reader = new StreamReader(path+file_name))
                    {
                        List<string> radiation_data = new List<string>();
                        List<string> time = new List<string>();
                        var file = reader.ReadToEnd().ToString().Split(new char[] { '\n' });
                        foreach (string line_of_file in file)
                        {
                            if (line_of_file != "")
                            {
                                var words = line_of_file.Split(new char[] { ' ' });
                                time.Add(words[1].Substring(1, words[1].Length - 3));
                                radiation_data.Add(words[3].Substring(0, words[3].Length - 2));
                            }
                        }
                        //answer = line_of_file;


                        for (int i = 0; i < time.Count; i++)
                        {
                            Console.WriteLine(time[i] + " radiation: " + radiation_data[i]);
                        }

                    }

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
