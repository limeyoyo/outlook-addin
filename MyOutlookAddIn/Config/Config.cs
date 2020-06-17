using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MyOutlookAddIn
{
    class Config
    {
        private static string keyWord = "";
        private static string folderPath = "";
        private static string keyConf = @"config\key.txt";
        private static string pathConf = @"config\path.txt";

        public static void SetKeyWord(string keyWord)
        {
            Config.keyWord = keyWord;
            using (StreamWriter sw = new StreamWriter(keyConf, false))
            {
                sw.WriteLine(keyWord);
            }
        }

        public static void SetFolderPath(string folderPath)
        {
            Config.folderPath = folderPath;
            using (StreamWriter sw = new StreamWriter(pathConf, false))
            {
                sw.WriteLine(folderPath);
            }
        }

        public static string GetKeyWord()
        {
            using (StreamReader sr = new StreamReader(keyConf, Encoding.Default))
            {
                if (sr.Peek() >= 0)
                {
                    Config.keyWord = sr.ReadLine().Trim();
                }
            }
            return Config.keyWord;
        }

        public static string GetFolderPath()
        {
            using (StreamReader sr = new StreamReader(pathConf, Encoding.Default))
            {
                if (sr.Peek() >= 0)
                {
                    Config.folderPath = sr.ReadLine().Trim();
                }
            }
            return Config.folderPath;
        }
    }
}
