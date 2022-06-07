using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ConverterExsel;

namespace TestConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ConverterExsel.ExcelConverter test = new ExcelConverter();
            string test_ansver = test.Converter("Port 8, Subchannel 0 - Пиранометр (СФ-06)", new DateTime(2022, 05, 06), new DateTime(2022, 05, 10)/*DateTime.UtcNow*/, "ArchivePob", "D:\\Work_25\\Projects_VS\\C#\\ConverterExsel", "test");
            Console.WriteLine(test_ansver);
            //var words = test_ansver.Split(new char[] { ' ' });
            //Console.WriteLine(words.Length);

            //foreach (string word in words)
            //{
            //    Console.WriteLine(word);
            //}
            //Console.WriteLine(words[1]);
            //string time = words[1].Substring(1, words[1].Length - 3);
            //Console.WriteLine(time);

            Console.ReadLine();
        }
    }
}
